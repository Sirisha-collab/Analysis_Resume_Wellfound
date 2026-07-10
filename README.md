# Analysis_Resume_Wellfound

NLP-based resume screening system that extracts candidate information from PDF, DOCX, and CSV resumes, analyzes profiles, scores candidates, and generates ranked Excel reports.

## Features
1. Resume parsing from PDF, DOCX, CSV
2. Extracts: Candidate name, Skills, Experience, Education
3. NLP processing using: spaCy (en_core_web_lg) and Sentence Transformers (all-MiniLM-L6-v2)
4. Semantic resume-job matching
5. Candidate scoring and ranking
6. Generates Top 50 and Top 5 candidate Excel reports

## Run
````
pip install -r requirements.txt
python SE_Pipeline.py
````

Output
