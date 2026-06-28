import sys
print("Python started", flush=True)

from fastapi import FastAPI
print("FastAPI imported", flush=True)

app = FastAPI()

@app.get("/health")
async def health():
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn
    print("Starting uvicorn...", flush=True)
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
