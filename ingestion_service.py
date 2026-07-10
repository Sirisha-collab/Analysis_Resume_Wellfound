from fastapi import FastAPI, UploadFile, File, HTTPException
import uuid
import os
import shutil
import pdfplumber
from docx import Document
from pdf2image import convert_from_path
import pytesseract
import csv

app = FastAPI()

TMP_DIR = "tmp"
os.makedirs(TMP_DIR, exist_ok=True)


# =====================================================
# OCR (only fallback)
# =====================================================

def ocr_pdf(path):
    text = ""
    images = convert_from_path(path)

    for img in images:
        text += pytesseract.image_to_string(img) + "\n"

    return text


# =====================================================
# CSV Reader
# =====================================================

def read_csv(path):
    try:
        rows = []

        with open(path, "r", encoding="utf-8", newline="") as csvfile:
            reader = csv.reader(csvfile)

            for row in reader:
                rows.append(", ".join(str(cell) for cell in row))

        return "\n".join(rows).strip()

    except UnicodeDecodeError:
        
        try:
            rows = []

            with open(path, "r", encoding="latin-1", newline="") as csvfile:
                reader = csv.reader(csvfile)

                for row in reader:
                    rows.append(", ".join(str(cell) for cell in row))

            return "\n".join(rows).strip()

        except Exception as e:
            print(f"CSV read error: {e}")
            return ""

    except Exception as e:
        print(f"CSV read error: {e}")
        return ""
    
# =====================================================
# PDF Reader (text + OCR fallback)
# =====================================================

def read_pdf(path):
    text = ""

    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"

    except Exception as e:
        print(f"PDF read error: {e}")

    # OCR fallback only if empty
    if not text.strip():
        try:
            print("OCR fallback triggered")
            text = ocr_pdf(path)
        except Exception as e:
            print(f"OCR failed: {e}")

    return text.strip()


# =====================================================
# DOCX Reader
# =====================================================

def read_docx(path):
    try:
        doc = Document(path)
        return "\n".join(p.text for p in doc.paragraphs).strip()
    except Exception as e:
        print(f"DOCX read error: {e}")
        return ""


def delete_file(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception as e:
        print(f"Cleanup error: {e}")


# =====================================================
# API ENDPOINT
# =====================================================

@app.post("/ingest")
async def ingest(file: UploadFile = File(...)):

    if not file.filename:
        raise HTTPException(status_code=400, detail="No file uploaded")

    ext = file.filename.lower().split(".")[-1]

    if ext not in ["pdf", "docx", "csv"]: 
        raise HTTPException(
            status_code=400,
            detail="Only PDF, DOCX, and CSV files are supported"
        )

    file_id = str(uuid.uuid4())
    file_path = os.path.join(TMP_DIR, f"{file_id}.{ext}")

    # Save file
    try:
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"File save failed: {str(e)}")

    # Extract text
    try:
        if ext == "pdf":
            text = read_pdf(file_path)
        elif ext == "docx":
            text = read_docx(file_path)
        elif ext == "csv":
            text = read_csv(file_path)

    except Exception as e:
        delete_file(file_path)
        raise HTTPException(status_code=500, detail=f"Text extraction failed: {str(e)}")

    # Cleanup temp file
    delete_file(file_path)

    # Handle empty extraction
    if not text or len(text.strip()) < 10:
        return {
            "file_id": file_id,
            "status": "failed",
            "message": "No readable text found in document",
            "text": ""
        }

    return {
        "file_id": file_id,
        "status": "success",
        "text_length": len(text),
        "text": text
    }