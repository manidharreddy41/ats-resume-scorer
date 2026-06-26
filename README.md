# ATS Resume Scorer

A Claude-powered ATS (Applicant Tracking System) resume scorer that mirrors real Workday, Greenhouse, and Lever scoring logic. Upload any resume against any job description and get an instant score, keyword analysis, and suggested improvements.

## Features

- **100-point ATS score** — broken down by keyword coverage, technical skills, experience relevance, quantified impact, and formatting
- **Matched & missing keywords** — see exactly what the ATS sees
- **AI-generated suggestions** — grouped by section (Skills, Summary, Experience) with points gain estimate
- **Apply suggestions** — select bullets to add and download a tailored DOCX with correct placement under the right company
- **Works for any job role** — fully JD-driven, not hardcoded to any industry
- **PDF and DOCX support** — upload either format

## Setup

### Requirements

- Python 3.11+
- Anthropic API key (get one at [console.anthropic.com](https://console.anthropic.com))

### Install

```bash
git clone https://github.com/manidharreddy41/ats-resume-scorer.git
cd ats-resume-scorer
pip install -r requirements.txt
```

### Run

```bash
uvicorn app:app --host 0.0.0.0 --port 8765
```

Open [http://localhost:8765](http://localhost:8765) in your browser.

Enter your Anthropic API key in the key bar at the top, then upload a resume and paste a job description.

## Usage

1. Paste a job description (or enter a URL)
2. Upload your resume (PDF or DOCX)
3. Click **Analyze** — get your ATS score in ~15 seconds
4. Switch to the **Suggestions** tab
5. Select the bullets you want to add
6. Click **Apply & Re-score** — download your tailored resume

## Tech Stack

- **Backend** — FastAPI, pdfplumber, python-docx, pdf2docx
- **AI** — Anthropic Claude (claude-haiku-4-5)
- **Frontend** — Vanilla JS, no frameworks
