from fastapi import FastAPI
import pandas as pd
import glob
import os
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# -----------------------------
# CORS (FIXED ORDER)
# -----------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# -----------------------------
# HELPERS
# -----------------------------
def get_latest_file(pattern: str):

    files = glob.glob(pattern)

    if not files:
        return None

    # safer than creation time on Windows
    latest_file = max(
        files,
        key=os.path.getmtime
    )

    return latest_file


def safe_columns(df, required_cols):

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    return df


# -----------------------------
# TOP 5 API
# -----------------------------
@app.get("/top5")
def top5():

    latest_file = get_latest_file("SDR_top_5_*.xlsx")

    if not latest_file:
        return {
            "status": "empty",
            "file": None,
            "total": 0,
            "candidates": []
        }

    try:
        df = pd.read_excel(latest_file)

    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to read Excel: {str(e)}",
            "candidates": []
        }

    # Ensure required columns exist (prevents crash)
    df = safe_columns(
        df,
        ["Name", "Score", "Score Breakdown"]
    )

    # Clean + normalize
    df = df.fillna("")

    # Sort defensively (in case file is not pre-sorted)
    if "Score" in df.columns:
        df = df.sort_values(
            by="Score",
            ascending=False
        )

    # Top 5
    df = df.head(5)

    return {
        "status": "success",
        "file": latest_file,
        "total_candidates_in_file": len(df),
        "candidates": df[
            ["Name", "Score", "Score Breakdown"]
        ].to_dict(orient="records")
    }