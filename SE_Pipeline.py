import os
from pydoc import text
import re
import pandas as pd
import requests
from datetime import datetime

from ingestion_service import read_pdf, read_docx, read_csv
from extract_service import (
    extract_name, extract_email, extract_phone,
    extract_location, compute_experience,
     extract_linkedin, extract_github,create_embedding
)      

from constants import SE_SKILLS, REQUIRED_BACKEND, REQUIRED_FRONTEND, HIGH, MEDIUM, PENALTIES, NEGATIVE

from rapidfuzz import fuzz
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

import spacy

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment


# -----------------------------
# EXCEL STYLING
# -----------------------------
def style_excel(file_path):
    wb = load_workbook(file_path)
    ws = wb.active

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    for column in ws.columns:
        max_length = 0
        col_letter = column[0].column_letter

        for cell in column:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_length + 5

    wb.save(file_path)


# -----------------------------
# CONFIG
# -----------------------------
RESUME_FOLDER = r"D:\Docs Latest\A Masters Required doc\Projects\Resumes Wellfound\Resume SE"
JOB_DESCRIPTION_FILE = r"D:\Docs Latest\A Masters Required doc\Projects\Resumes Wellfound\Resume SE\job description.txt"
CSV_FILE = r"D:\Docs Latest\A Masters Required doc\Projects\Resumes Wellfound\Resume SE\Applicants.csv"

with open(JOB_DESCRIPTION_FILE, "r", encoding="utf-8") as f:
    job_description = f.read()

job_embedding = create_embedding(job_description)

#exp into 3.5 instead of 3,
def safe_exp(val):
    try:
        return float(val)
    except:
        return 0.0

def extract_skills(text):

    text = text.lower()

    found = []

    for skill in SE_SKILLS:

        score = fuzz.partial_ratio(
            skill,
            text
        )

        if score >= 85:
            found.append(skill)

    return sorted(list(set(found)))

# -----------------------------
# METRICS DETECTION
# -----------------------------
def extract_metrics(row):
    text = " ".join([
        str(row.get("name", "")),
        str(row.get("location_city", "")),
        str(row.get("canonical_skills", ""))
    ]).lower()

    patterns = [
        r"\d+\+?\s*%",
        r"improved", r"built", r"developed",
        r"deployed", r"scaled", r"optimized",
        r"reduced", r"increased", r"architected"
    ]

    return any(re.search(p, text) for p in patterns)


# -----------------------------
# SCORING
# -----------------------------
def score_candidate(r):
    score = 0
    reasons = []
    negative_penalty = 0

    skills_raw = r.get("canonical_skills") or ""

    if isinstance(skills_raw, list):
        skills_text = " ".join(skills_raw).lower()
    else:
        skills_text = str(skills_raw).lower()

    exp = safe_exp(r.get("years_experience_in_role", 0))

    linkedin = r.get("linkedin_url")
    github = r.get("github_url")

    # rejection rules
    if exp < 2:
        return -1, "Rejected: less than 2 years experience",""

    if not linkedin or not github:
        return -1, "Rejected: missing LinkedIn or GitHub", ""

    if 2 <= exp <= 8:
        score += 15
    elif exp > 8:
        score += 5

    
    negative_skills =[]

    has_backend = any(x in skills_text for x in REQUIRED_BACKEND)
    has_frontend = any(x in skills_text for x in REQUIRED_FRONTEND)

    if not has_backend:
        return -1, "Rejected: no backend engineering skills", ""

    if not has_frontend:
        return -1, "Rejected: no frontend engineering skills", ""

    # Skill scoring
    matched = set()

    for skill in HIGH:
        if skill in skills_text and skill not in matched:
            score += 20
            reasons.append(f"{skill}")
            matched.add(skill)

    for skill in MEDIUM:
        if skill in skills_text and skill not in matched:
            score += 10
            reasons.append(f"{skill} ")
            matched.add(skill)

    # Full stack requirement
    if has_frontend and has_backend:
        score += 30
        reasons.append("Full-stack capability")

    # Cloud / deployment bonus
    if any(x in skills_text for x in ["aws", "azure", "gcp"]):
        score += 10
        reasons.append("Cloud experience")

    if "git" in skills_text:
        score += 10
        reasons.append("Git experience")

    # Evidence of engineering work
    if extract_metrics(r):
        score += 20
        reasons.append(" ")

    # Profile signals
    score += 10
    reasons.append(" ")

    # Apply skill penalties
    for skill, penalty in PENALTIES.items():
        if skill in skills_text:
            score += penalty
            reasons.append(f"{skill} ({penalty})")
            negative_skills.append(f"{skill}: {penalty}")

    # Apply non-engineering penalties
    for skill in NEGATIVE:
        if skill in skills_text:
            score -= 70
            reasons.append(f"{skill} (-70)")
            negative_skills.append(f"{skill}: -70")

    # Final quality gate
    if score < 70:
        negative_reason = ""
        if negative_skills:
            negative_reason = " | Negative skills: " + ", ".join(negative_skills)
    
    return score, "; ".join(reasons), ", ".join(negative_skills)

# -----------------------------
# CSV ONLY PIPELINE
# -----------------------------

all_candidates = []
ranked_candidates = []

print("Reading CSV:", CSV_FILE)


try:
    df = pd.read_csv(
        CSV_FILE,
        encoding="utf-8",
        sep=",",
        engine="python"
    )
except UnicodeDecodeError:
    print("UTF-8 failed, trying cp1252...")
    df = pd.read_csv(
        CSV_FILE,
        encoding="cp1252",
        sep=",",
        engine="python"
    )

df.columns = df.columns.str.strip()

print("CSV loaded successfully")
print("Total applicants:", len(df))

for _, row_data in df.iterrows():

    row = {
        "name": row_data.get("name", ""),
        "Email": row_data.get("Email", ""),
        "Phone": row_data.get("Phone", ""),
        "location_city": row_data.get("location_city", ""),
        "years_experience_in_role": safe_exp(
            row_data.get("years_experience_in_role", 0)
        ),
        "canonical_skills": row_data.get("canonical_skills", ""),
        "linkedin_url": row_data.get("linkedin_url", ""),
        "github_url": row_data.get("github_url", "")
    }

    all_candidates.append(row)

    score, reason, negative_penalty = score_candidate(row)

    if score < 0:
        continue

    row["Score"] = score
    row["Score Breakdown"] = reason
    row["Negative Penalty"] = negative_penalty

    ranked_candidates.append(row)


# -----------------------------
# EXPORT RESULTS
# -----------------------------

SE_DIR = "SE_results"
os.makedirs(SE_DIR, exist_ok=True)

timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')


# Complete CSV repository
df_all = pd.DataFrame(all_candidates)

repo_file = os.path.join(
    SE_DIR,
    f"se_candidate_repo_{timestamp}.xlsx"
)

df_all.to_excel(repo_file, index=False)
style_excel(repo_file)

print("Candidate repo saved:", repo_file)


# Top 50 candidates
top50 = sorted(
    ranked_candidates,
    key=lambda x: x["Score"],
    reverse=True
)[:50]


df_top50 = pd.DataFrame(top50)

top50_file = os.path.join(
    SE_DIR,
    f"se_top_50_{timestamp}.xlsx"
)

df_top50.to_excel(top50_file, index=False)
style_excel(top50_file)

print("Top 50 saved:", top50_file)

#Export candidate repo
SE_DIR = "SE_results"
os.makedirs(SE_DIR, exist_ok=True)

timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

df_all = pd.DataFrame(all_candidates)

repo_file =  os.path.join(
    SE_DIR,
    f"se_candidate_repo_{timestamp}.xlsx"
)
df_all.to_excel(repo_file, index=False)
style_excel(repo_file)

print(" SE Candidate repo saved:", repo_file)

#Export top 5 candidates
top5 = sorted(ranked_candidates, key=lambda x: x["Score"], reverse=True)[:5]

df_top5 = pd.DataFrame(top5)

top5_file =  os.path.join(
    SE_DIR,
    f"se_top_5_{timestamp}.xlsx"
)
df_top5.to_excel(top5_file, index=False)
style_excel(top5_file)

print(" SE Top 5 saved:", top5_file)