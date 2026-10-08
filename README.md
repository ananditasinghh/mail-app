# Mailer — Anandita

## Login
- Username: `admin`
- Password: `ananditaisbest`

## Run locally
1. Open Terminal and `cd` into this folder.
2. First time only: `pip3 install -r requirements.txt`
3. Run: `streamlit run mail_app.py`
4. Opens at http://localhost:8501 — log in with the credentials above.

## Sender account
- Gmail address is prefilled (anandita.singh.nsut21@gmail.com).
- Enter your 16-character App Password in the sidebar each session
  (generate at myaccount.google.com/apppasswords — not your normal password).

## What's inside
- **Recipients** — paste a messy list, extract + validate names, generate emails from one example pattern, or upload a CSV (max 100 per run).
- **Compose** — "First outreach" / "Follow-up" templates (your resume-based wording, with {name} and {company} placeholders), editable before sending. Optional PDF attachment.
- **Send** — preview, confirm, send immediately.
- **Tracker** — confirm company per email, then logs to `sent_tracker_admin.xlsx` (first mail vs follow-up 1/2/3, with dates). Stays in this folder.
