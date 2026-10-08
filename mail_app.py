import streamlit as st
import pandas as pd
import smtplib
import time
import re
import os
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

st.set_page_config(page_title="Mailer", page_icon="✉️", layout="centered")

# ---------------------------------------------------------------------------
# Login gate
# ---------------------------------------------------------------------------
try:
    USERS = dict(st.secrets["users"])
except Exception:
    # Local fallback — used only when running on your own machine without a secrets.toml
    USERS = {
        "admin": "ananditaisbest",
    }

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "username" not in st.session_state:
    st.session_state.username = None

if not st.session_state.authenticated:
    st.markdown(
        """
        <style>
        .stApp { background: #F7F8FB; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("## ✉️ Mailer")
    st.caption("Please log in to continue.")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in", type="primary")
        if submitted:
            uname = username.strip().lower()
            if uname in USERS and password == USERS[uname]:
                st.session_state.authenticated = True
                st.session_state.username = uname
                st.rerun()
            else:
                st.error("Incorrect username or password.")
    st.stop()

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Source+Serif+4:wght@600&display=swap');

    html, body, [class*="css"]  {
        font-family: 'Inter', -apple-system, sans-serif;
    }

    :root {
        --ink: #16233A;
        --muted: #5B6B84;
        --line: #E2E7EF;
        --paper: #F7F8FB;
        --accent: #0E6E5A;
        --accent-soft: #E4F2EE;
    }

    .stApp { background: var(--paper); }

    .app-title {
        font-family: 'Source Serif 4', serif;
        font-size: 2.1rem;
        font-weight: 600;
        color: var(--ink);
        margin-bottom: 0.1rem;
    }
    .app-subtitle {
        color: var(--muted);
        font-size: 0.95rem;
        margin-bottom: 1.6rem;
    }

    section[data-testid="stSidebar"] {
        background: #FFFFFF;
        border-right: 1px solid var(--line);
    }
    section[data-testid="stSidebar"] .block-container { padding-top: 2rem; }

    .status-pill {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 500;
        margin-bottom: 0.4rem;
    }
    .status-ok { background: var(--accent-soft); color: var(--accent); }
    .status-pending { background: #FBEFE0; color: #8A5A15; }

    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: #FFFFFF;
        border: 1px solid var(--line) !important;
        border-radius: 14px !important;
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 8px;
        font-weight: 500;
        border: 1px solid var(--line);
    }
    .stButton > button[kind="primary"] {
        background: var(--accent);
        border: none;
    }
    .stButton > button[kind="primary"]:hover {
        background: #0B5A49;
    }

    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        font-weight: 500;
        color: var(--muted);
    }
    .stTabs [aria-selected="true"] {
        color: var(--accent) !important;
    }

    h1, h2, h3 { color: var(--ink); }

    .field-caption { color: var(--muted); font-size: 0.85rem; margin-top: -0.4rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="app-title">✉️ Mailer</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-subtitle">Personalized bulk email, sent from your own Gmail. '
    'Runs entirely on your machine — nothing leaves your computer except the emails themselves.</div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
defaults = {
    "sent_log": [],
    "generated_csv": None,
    "extracted_names": None,
    "direct_df": None,
    "pending_tracker": None,
    "subject_template": "Quick update for {name}",
    "body_template": "Hi {name},\n\nWrite your email here.\n\nBest,\nYour Name",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

TRACKER_FILE = f"sent_tracker_{st.session_state.username}.xlsx"
TRACKER_COLUMNS = [
    "email", "company",
    "sent", "sent_date",
    "follow_up_1", "follow_up_1_date",
    "follow_up_2", "follow_up_2_date",
    "follow_up_3", "follow_up_3_date",
]


def guess_company(email):
    domain = email.split("@")[-1].lower().strip()
    parts = [p for p in domain.split(".") if p]
    while len(parts) > 1 and parts[0] in ("www", "mail", "smtp", "mx", "email"):
        parts = parts[1:]
    return parts[0] if parts else domain


def load_tracker():
    if os.path.exists(TRACKER_FILE):
        t = pd.read_excel(TRACKER_FILE)
        for col in TRACKER_COLUMNS:
            if col not in t.columns:
                t[col] = ""
        return t[TRACKER_COLUMNS].fillna("")
    return pd.DataFrame(columns=TRACKER_COLUMNS)


def update_tracker(sent_emails, company_map, date_str):
    tracker = load_tracker()
    tracker["_email_lower"] = tracker["email"].astype(str).str.lower()

    for email in sent_emails:
        email_l = email.lower()
        company = company_map.get(email, guess_company(email))
        match = tracker.index[tracker["_email_lower"] == email_l]

        if len(match) == 0:
            new_row = {c: "" for c in TRACKER_COLUMNS}
            new_row["email"] = email
            new_row["company"] = company
            new_row["sent"] = "Yes"
            new_row["sent_date"] = date_str
            tracker = pd.concat([tracker, pd.DataFrame([{**new_row, "_email_lower": email_l}])], ignore_index=True)
        else:
            idx = match[0]
            if not str(tracker.at[idx, "company"]).strip():
                tracker.at[idx, "company"] = company
            for n in (1, 2, 3):
                col, dcol = f"follow_up_{n}", f"follow_up_{n}_date"
                if not str(tracker.at[idx, col]).strip():
                    tracker.at[idx, col] = "Yes"
                    tracker.at[idx, dcol] = date_str
                    break

    tracker = tracker.drop(columns=["_email_lower"])
    tracker.to_excel(TRACKER_FILE, index=False)
    return tracker


def guess_pattern(example_name, example_email):
    if "@" not in example_email:
        return None, None
    local, domain = example_email.split("@", 1)
    tokens = [t for t in re.split(r"\s+", example_name.strip().lower()) if t]
    first = tokens[0] if tokens else ""
    last = tokens[-1] if len(tokens) > 1 else ""

    templates = {
        "first.last": f"{first}.{last}" if last else first,
        "last.first": f"{last}.{first}" if last else first,
        "firstlast": f"{first}{last}",
        "lastfirst": f"{last}{first}",
        "first_last": f"{first}_{last}" if last else first,
        "last_first": f"{last}_{first}" if last else first,
        "first-last": f"{first}-{last}" if last else first,
        "f.last": f"{first[0]}.{last}" if first and last else "",
        "first.l": f"{first}.{last[0]}" if first and last else "",
        "flast": f"{first[0]}{last}" if first and last else "",
        "firstl": f"{first}{last[0]}" if first and last else "",
        "first": first,
        "last": last,
    }
    local_lower = local.strip().lower()
    for name, val in templates.items():
        if val and val == local_lower:
            return name, domain
    return None, domain


def apply_template(template_name, full_name, domain):
    tokens = [t for t in re.split(r"\s+", full_name.strip().lower()) if t]
    first = tokens[0] if tokens else ""
    last = tokens[-1] if len(tokens) > 1 else ""
    mapping = {
        "first.last": f"{first}.{last}" if last else first,
        "last.first": f"{last}.{first}" if last else first,
        "firstlast": f"{first}{last}",
        "lastfirst": f"{last}{first}",
        "first_last": f"{first}_{last}" if last else first,
        "last_first": f"{last}_{first}" if last else first,
        "first-last": f"{first}-{last}" if last else first,
        "f.last": f"{first[0]}.{last}" if first and last else first,
        "first.l": f"{first}.{last[0]}" if first and last else first,
        "flast": f"{first[0]}{last}" if first and last else first,
        "firstl": f"{first}{last[0]}" if first and last else first,
        "first": first,
        "last": last,
    }
    local = mapping.get(template_name, first)
    return f"{local}@{domain}"


def extract_name_from_line(line):
    line = line.strip()
    if not line:
        return ""
    # strip emails from anywhere in the line
    line = re.sub(r"\S+@\S+", "", line)
    # strip leading bullets / numbering: "1.", "1)", "-", "*", "•"
    line = re.sub(r"^[\s\-\*\u2022]+", "", line)
    line = re.sub(r"^\d+[\.\)]\s*", "", line)
    # cut at the first separator that usually introduces title/company/extra info
    parts = re.split(r"\s+[-–|]\s+|,|\(", line)
    candidate = parts[0].strip() if parts else line
    if not candidate:
        return ""
    words = candidate.split()
    # prefer a run of Title-Case-looking words (typical for names)
    name_words = []
    for w in words:
        if re.match(r"^[A-Z][a-zA-Z'.-]*$", w):
            name_words.append(w)
        elif name_words:
            break
    if name_words:
        return " ".join(name_words[:4])
    # fallback: no clean Title Case run found — just take first couple of words as-is
    return " ".join(words[:3]).strip(" .,-")


EMAIL_TEMPLATES = {
    "First outreach": {
        "subject": "Application for Product Roles at {company}",
        "body": """Dear {name},

I am Anandita Singh, currently a Junior Data Analyst at Boston Consulting Group and a 2025 B.Tech graduate from NSUT Delhi, and I am writing to apply for the Product Related roles at {company}.

My experience sits at the intersection of technology, data, and product problem-solving. At BCG, I have owned the build of an AI platform orchestrating 6 specialized LLM agents, enabling consultants to query 3,000+ projects in plain English and reducing a multi-day review process to under 10 minutes. I have also built automation and decision-support tools used across global projects.

Earlier, at OYO, I worked on solving operational and pricing problems at scale, building automation that reduced reporting time by ~70% and correcting pricing and discount mismatches across 300+ properties.

Beyond my professional experience, I independently built and published ChatFit, an open-source Python library for LLM context management, taking it from an idea to a usable product on PyPI.

What draws me to product is the opportunity to identify a real user problem, understand it deeply, build the right solution, and measure whether it actually works. I would be excited to bring this builder-oriented and analytical approach to {company}.

I have attached my resume and would be grateful for the opportunity to connect.

Warm regards,
Anandita Singh
+91 8630723651""",
    },
    "Follow-up": {
        "subject": "Following up — Application at {company}",
        "body": """Dear {name},

I wanted to follow up on my earlier email regarding the opportunity at {company}, specifically in Product & Analytics.

I'm currently a Junior Data Analyst at BCG, where I work on building AI and data-driven solutions, including an internal platform that orchestrates 6 specialized LLM agents and has reduced a multi-day review process to under 10 minutes. Prior to BCG, I worked at OYO, where I built automation and worked on pricing and revenue problems across 300+ properties.

With my experience across travel-tech, data, AI, and building products independently, I believe the APM role at {company} would be a particularly strong fit for my background and interests. I would really appreciate it if you could let me know if there are any updates regarding my application or if there is someone I could connect with regarding the role.

Thank you for your time, and I look forward to hearing from you.

Warm regards,
Anandita Singh
+91 8630723651""",
    },
}


def render(template, row):
    return template.format(**row.to_dict())


# ---------------------------------------------------------------------------
# Sidebar — sender account (Gmail address prefilled; App Password typed each session)
# ---------------------------------------------------------------------------
if "gmail_address" not in st.session_state:
    st.session_state["gmail_address"] = "anandita.singh.nsut21@gmail.com"

with st.sidebar:
    st.markdown("### Sender account")
    gmail_address = st.text_input("Gmail address", placeholder="you@gmail.com", key="gmail_address")
    app_password = st.text_input(
        "16-character App Password",
        type="password",
        key="app_password",
        help="Generate at myaccount.google.com/apppasswords — not your normal Gmail password.",
    )

    account_ready = bool(gmail_address.strip()) and bool(app_password.strip())
    if account_ready:
        st.markdown('<span class="status-pill status-ok">Account ready</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="status-pill status-pending">Account not set</span>', unsafe_allow_html=True)

    st.caption(f"Logged in as **{st.session_state.username}**. Your address and password stay only in this local session.")

    st.divider()
    if os.path.exists(TRACKER_FILE):
        st.markdown("### Your tracker")
        with open(TRACKER_FILE, "rb") as f:
            st.download_button(f"⬇️ Download {TRACKER_FILE}", data=f.read(), file_name=TRACKER_FILE, use_container_width=True)

    st.divider()
    if st.button("Log out", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_recipients, tab_compose, tab_send, tab_tracker = st.tabs(
    ["📇 Recipients", "✍️ Compose", "🚀 Send", "📊 Tracker"]
)

# ---------- Recipients ----------
with tab_recipients:
    with st.container(border=True):
        st.markdown("#### Extract names from a pasted list")
        st.caption("Paste anything — a numbered list, names with titles/companies attached, a messy copy-paste. I'll pull out just the names.")

        raw_paste = st.text_area(
            "Paste your list",
            placeholder="1. Aleena Sharma - Product Manager, Google\nAnandita Singh, NSUT'25\nMeezu Singh (meezu.singh@bain.com)",
            height=130,
        )

        if st.button("Extract names"):
            if not raw_paste.strip():
                st.error("Paste something first.")
            else:
                lines = [l for l in raw_paste.split("\n") if l.strip()]
                if len(lines) > 100:
                    st.warning(f"Found {len(lines)} lines — trimming to first 100.")
                    lines = lines[:100]
                rows = [{"extracted_name": extract_name_from_line(l), "original_text": l.strip()} for l in lines]
                st.session_state.extracted_names = pd.DataFrame(rows)
                st.success(f"Extracted {len(rows)} name(s) — check them below and fix any that look wrong.")

        if st.session_state.extracted_names is not None:
            st.caption("Validate — edit any wrong names directly, or delete rows that aren't people.")
            edited_names = st.data_editor(
                st.session_state.extracted_names,
                num_rows="dynamic",
                use_container_width=True,
                key="names_editor",
                column_config={"original_text": st.column_config.TextColumn("Original text", disabled=True)},
            )
            st.session_state.extracted_names = edited_names

    if st.session_state.extracted_names is not None and len(st.session_state.extracted_names) > 0:
        with st.container(border=True):
            st.markdown("#### Generate emails from validated names")
            st.caption("Give one example of Name → Email, and the pattern is applied to everyone above.")

            c1, c2 = st.columns(2)
            with c1:
                example_name = st.text_input("Example name", placeholder="Anandita Singh")
            with c2:
                example_email = st.text_input("Example email", placeholder="anandita.singh@bcg.com")

            if st.button("Generate email list"):
                if not example_name or not example_email:
                    st.error("Please fill in the example name and example email first.")
                else:
                    template_name, domain = guess_pattern(example_name, example_email)
                    if not template_name:
                        st.error("Couldn't detect the pattern from that example. Try something like 'firstname.lastname@domain.com'.")
                    else:
                        names_list = [n.strip() for n in st.session_state.extracted_names["extracted_name"] if str(n).strip()]
                        rows = [{"name": n, "email": apply_template(template_name, n, domain)} for n in names_list]
                        st.session_state.generated_csv = pd.DataFrame(rows)
                        st.success(f"Detected pattern: **{template_name}@{domain}** — review below.")

            if st.session_state.generated_csv is not None:
                edited_df = st.data_editor(st.session_state.generated_csv, num_rows="dynamic", use_container_width=True)
                csv_bytes = edited_df.to_csv(index=False).encode("utf-8")
                bcol1, bcol2 = st.columns(2)
                with bcol1:
                    st.download_button("⬇️ Download CSV", data=csv_bytes, file_name="generated_recipients.csv", mime="text/csv", use_container_width=True)
                with bcol2:
                    if st.button("Use this list directly", use_container_width=True):
                        st.session_state.direct_df = edited_df.copy()
                        st.success("Loaded below.")

    with st.container(border=True):
        st.markdown("#### Upload recipient list")
        uploaded_file = st.file_uploader("CSV with an 'email' column (plus any others for personalization)", type=["csv"])

        df = None
        if uploaded_file:
            df = pd.read_csv(uploaded_file)
            df.columns = [c.strip().lower() for c in df.columns]
        elif st.session_state.direct_df is not None:
            df = st.session_state.direct_df.copy()
            df.columns = [c.strip().lower() for c in df.columns]
            st.info("Using the list generated above.")

        if df is not None:
            if "email" not in df.columns:
                st.error("Your CSV must have a column named 'email'.")
                df = None
            else:
                df = df[df["email"].notna() & (df["email"].astype(str).str.strip() != "")]
                df = df.drop_duplicates(subset="email").reset_index(drop=True)
                if len(df) > 100:
                    st.warning(f"Found {len(df)} recipients — trimming to first 100.")
                    df = df.iloc[:100].reset_index(drop=True)
                st.success(f"{len(df)} recipient(s) loaded · columns: {list(df.columns)}")
                st.dataframe(df, use_container_width=True, height=220)
        else:
            st.caption("No recipients loaded yet.")

# ---------- Compose ----------
with tab_compose:
    with st.container(border=True):
        st.markdown("#### Choose a starting template")
        template_choice = st.radio("Email type", ["First outreach", "Follow-up"], horizontal=True)
        if st.button("Load this template"):
            chosen = EMAIL_TEMPLATES[template_choice]
            st.session_state["subject_template"] = chosen["subject"]
            st.session_state["body_template"] = chosen["body"]
            st.success(f"Loaded the {template_choice.lower()} template below — edit freely, uses {{name}} and {{company}}.")

    with st.container(border=True):
        st.markdown("#### Write your email")
        subject_template = st.text_input("Subject", key="subject_template")
        body_template = st.text_area("Body (use {column_name} for personalization)", key="body_template", height=200)

        if df is not None:
            used_fields = set(re.findall(r"\{(\w+)\}", subject_template + body_template))
            missing = used_fields - set(df.columns)
            if missing:
                st.warning(f"These placeholders aren't columns in your recipient list: {missing}")

        pdf_attachment = st.file_uploader("Attach a PDF (optional, same file sent to everyone)", type=["pdf"])
        if pdf_attachment:
            st.caption(f"Attached: {pdf_attachment.name} ({pdf_attachment.size / 1024:.0f} KB)")

    with st.container(border=True):
        st.markdown("#### Preview")
        if df is not None and len(df) > 0:
            preview_idx = 0 if len(df) == 1 else st.slider("Preview recipient #", 1, len(df), 1) - 1
            row = df.iloc[preview_idx]
            try:
                st.text_input("To", value=row["email"], disabled=True)
                st.text_input("Subject", value=render(subject_template, row), disabled=True)
                st.text_area("Body", value=render(body_template, row), height=180, disabled=True)
            except KeyError as e:
                st.error(f"Missing column for placeholder: {e}")
        else:
            st.caption("Load recipients first to see a preview.")

# ---------- Send ----------
def send_batch(gmail_address, app_password, df, subject_template, body_template, pdf_attachment):
    progress = st.progress(0)
    status = st.empty()
    sent, failed = [], []

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(gmail_address.strip(), app_password.strip().replace(" ", ""))
    except smtplib.SMTPAuthenticationError:
        st.error("Login failed. Check your Gmail address and App Password in the sidebar.")
        st.stop()

    pdf_bytes = pdf_attachment.getvalue() if pdf_attachment else None
    pdf_name = pdf_attachment.name if pdf_attachment else None

    for i, row in df.iterrows():
        recipient = row["email"]
        try:
            subject = render(subject_template, row)
            body = render(body_template, row)

            msg = MIMEMultipart()
            msg["From"] = gmail_address.strip()
            msg["To"] = recipient
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))

            if pdf_bytes:
                part = MIMEApplication(pdf_bytes, _subtype="pdf")
                part.add_header("Content-Disposition", "attachment", filename=pdf_name)
                msg.attach(part)

            server.sendmail(gmail_address.strip(), recipient, msg.as_string())
            sent.append(recipient)
            status.write(f"✅ Sent to {recipient} ({i+1}/{len(df)})")
        except Exception as e:
            failed.append((recipient, str(e)))
            status.write(f"❌ Failed for {recipient}: {e}")

        progress.progress((i + 1) / len(df))
        time.sleep(2)  # gentle delay to avoid rate limits / spam flags

    server.quit()
    return sent, failed


with tab_send:
    with st.container(border=True):
        st.markdown("#### Ready to send")

        ready = df is not None and len(df) > 0 and account_ready

        if not account_ready:
            st.info("Enter your Gmail address and App Password in the sidebar first.")
        elif df is None or len(df) == 0:
            st.info("Load a recipient list in the Recipients tab first.")

        if ready:
            st.write(f"Sending as **{gmail_address}** to **{len(df)}** recipient(s).")

            confirm = st.checkbox("I've checked the preview and want to send now.")

            if st.button("🚀 Send all emails", disabled=not confirm, type="primary"):
                sent, failed = send_batch(gmail_address, app_password, df, subject_template, body_template, pdf_attachment)

                st.session_state.sent_log = sent
                st.session_state.pending_tracker = {"emails": sent, "date": date.today().isoformat()}

                st.success(f"Done. Sent: {len(sent)}. Failed: {len(failed)}.")
                if failed:
                    st.error("Failed recipients:")
                    for r, err in failed:
                        st.write(f"- {r}: {err}")
                st.info("Head to the Tracker tab to confirm companies and log this batch.")

# ---------- Tracker ----------
with tab_tracker:
    if st.session_state.pending_tracker is not None:
        pending = st.session_state.pending_tracker
        with st.container(border=True):
            st.markdown("#### Confirm & log this batch")
            st.caption("Check the guessed company for each email just sent, fix any, then save.")
            if pending["emails"]:
                preview_rows = [{"email": e, "company": guess_company(e)} for e in pending["emails"]]
                edited_companies = st.data_editor(pd.DataFrame(preview_rows), use_container_width=True, key="company_editor")

                if st.button("✅ Confirm & save to tracker", type="primary"):
                    company_map = dict(zip(edited_companies["email"], edited_companies["company"]))
                    tracker = update_tracker(pending["emails"], company_map, pending["date"])
                    st.session_state.pending_tracker = None
                    st.success(f"Tracker updated → {TRACKER_FILE}")
                    st.dataframe(tracker, use_container_width=True)
            else:
                st.info("No successful sends to record.")
                st.session_state.pending_tracker = None

    if os.path.exists(TRACKER_FILE) and st.session_state.pending_tracker is None:
        with st.container(border=True):
            st.markdown(f"#### Full tracker · {TRACKER_FILE}")
            tracker_df = load_tracker()
            st.dataframe(tracker_df, use_container_width=True, height=320)
            with open(TRACKER_FILE, "rb") as f:
                st.download_button("⬇️ Download tracker", data=f.read(), file_name=TRACKER_FILE, key="tracker_view_download")
    elif st.session_state.pending_tracker is None:
        st.caption("No tracker yet — it's created automatically after your first send.")
