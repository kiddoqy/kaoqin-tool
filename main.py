"""
考勤工具 Web 服务
"""
import io
import os
import sys
import uuid
import re
import traceback
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse

BASE_DIR = Path(__file__).resolve().parent

# 尝试导入核心模块
try:
    from attendance_core import process, generate_excel
    CORE_OK = True
except Exception as e:
    CORE_OK = False
    CORE_ERR = f"核心模块加载失败: {e}\n{traceback.format_exc()}"

app = FastAPI(title="考勤生成工具")
cache = {}

TEMPLATE_PATH = BASE_DIR / "templates" / "index.html"


@app.get("/", response_class=HTMLResponse)
async def index():
    if not CORE_OK:
        return HTMLResponse(content=f"<pre>{CORE_ERR}</pre>", status_code=500)
    html = TEMPLATE_PATH.read_text(encoding="utf-8") if TEMPLATE_PATH.exists() else "<h1>模板文件未找到</h1>"
    return HTMLResponse(content=html)


@app.get("/health")
async def health():
    return {"status": "ok", "core": CORE_OK}


@app.post("/api/upload")
async def upload(
    record: UploadFile = File(...),
    monthly: UploadFile = File(...),
):
    if not CORE_OK:
        raise HTTPException(500, CORE_ERR)

    for f, label in [(record, "打卡时间记录"), (monthly, "月报")]:
        if not f.filename or not f.filename.endswith('.xlsx'):
            raise HTTPException(400, f"{label} 必须是 .xlsx 文件")

    try:
        record_bytes = await record.read()
        monthly_bytes = await monthly.read()
    except Exception:
        raise HTTPException(400, "文件读取失败")

    try:
        result = process(io.BytesIO(record_bytes), io.BytesIO(monthly_bytes))
    except Exception as e:
        raise HTTPException(500, f"处理失败: {e}")

    session_id = str(uuid.uuid4())[:8]
    cache[session_id] = {
        'result': result,
        'record_name': record.filename,
        'monthly_name': monthly.filename,
    }

    persons = result['persons']
    num_days = result['num_days']
    daily_abnormal = []
    for d in range(num_days):
        late_c = sum(1 for p in persons if d < len(p['daily_statuses']) and p['daily_statuses'][d] == '≦')
        miss_c = sum(1 for p in persons if d < len(p['daily_statuses']) and p['daily_statuses'][d] == '缺卡')
        daily_abnormal.append({'day': d + 1, 'late': late_c, 'miss': miss_c})

    return {
        'session_id': session_id,
        'persons': persons[:50],
        'weekdays': result['weekdays'],
        'num_days': num_days,
        'summary': result['summary'],
        'daily_abnormal': daily_abnormal,
        'month_rest_days': result['month_rest_days'],
    }


@app.get("/api/download/{session_id}")
async def download(session_id: str):
    if not CORE_OK:
        raise HTTPException(500, CORE_ERR)
    if session_id not in cache:
        raise HTTPException(404, "会话已过期，请重新上传")

    data = cache[session_id]
    result = data['result']
    excel_bytes = generate_excel(result['persons'], result['weekdays'])

    record_name = data['record_name'].replace('.xlsx', '')
    m = re.search(r'(\d{4})(\d{2})\d{2}-(\d{4})(\d{2})\d{2}', record_name)
    filename = f"{int(m.group(2))}月考勤.xlsx" if m else "考勤表.xlsx"
    encoded_filename = quote(filename)

    return StreamingResponse(
        excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"}
    )
