from fastapi import FastAPI
app = FastAPI()

@app.get("/")
def root():
    return {"hello": "world"}

@app.get("/health")
def health():
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn, os
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
