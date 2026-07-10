from fastapi import FastAPI, UploadFile, File
import requests

app = FastAPI()

INGEST = "http://localhost:8001/ingest"
EXTRACT = "http://localhost:8002/extract"
SCORE = "http://localhost:8003/score"

DB = []   # in-memory store


@app.post("/process")
async def process(file: UploadFile = File(...)):

    files = {"file": (file.filename, await file.read())}

    ing = requests.post(INGEST, files=files).json()
    text = ing["text"]

    extracted = requests.post(EXTRACT, json={"text": text}).json()
    scored = requests.post(SCORE, json=extracted).json()

    record = {
        **extracted,
        **scored
    }

    DB.append(record)

    return record


@app.get("/stats")
def stats():
    return {
        "total_processed": len(DB),
        "avg_score": sum(x["score"] for x in DB) / len(DB) if DB else 0
    }


@app.get("/leaderboard")
def leaderboard(limit: int = 10):

    return sorted(DB, key=lambda x: x["score"], reverse=True)[:limit]


@app.get("/filter")
def filter(min_score: int = 0, skill: str = None):

    result = DB

    if skill:
        result = [r for r in result if skill in r.get("skills", [])]

    result = [r for r in result if r["score"] >= min_score]

    return result