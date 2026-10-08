"""
EduNexus AI  -  The Unified Student Success & Intelligence Hub
Built for the Bytexel 24-hour hackathon.

Run:  streamlit run app.py
Data: Student360_1000_Feature_Only_Single_Sheet_With_Gender.xlsx  (sheet: student_features)
      Put the file next to app.py (or upload it from the sidebar).  If it is missing the
      app falls back to realistic mock data with the SAME schema so it never crashes.
"""
import html
import io
import os
import re
import urllib.parse
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="EduNexus AI",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

FILE_NAME = "Student360_1000_Feature_Only_Single_Sheet_With_Gender.xlsx"
SHEET_NAME = "student_features"

RISK_ORDER = ["Critical", "High", "Moderate", "Safe"]
RISK_RANK = {r: i for i, r in enumerate(RISK_ORDER)}
RISK_COLORS = {"Critical": "#EF4444", "High": "#F97316", "Moderate": "#FACC15", "Safe": "#22C55E"}
BURN_COLORS = {"High": "#EF4444", "Moderate": "#FACC15", "Low": "#22C55E"}
STATUS_ORDER = ["Pending", "In Progress", "Completed"]
STATUS_COLORS = {"Pending": "#F87171", "In Progress": "#38BDF8", "Completed": "#22C55E"}
COMPETENCIES = ["Coding", "Aptitude", "Interview", "Communication"]
CTA_A, CTA_B = "#7C3AED", "#D946EF"  # call-to-action gradient (violet -> magenta)

# ─────────────────────────────────────────────────────────────────────────────
# STYLING  (dark theme, high contrast, vivid CTA buttons)
# ─────────────────────────────────────────────────────────────────────────────
CSS = f"""
<style>
.stApp {{background:#0E1117;}}
[data-testid="stSidebar"] {{background:#111827; border-right:1px solid #1F2A44;}}
.stApp, .stApp p, .stApp li, .stApp label, .stApp span.st-emotion-cache-0,
[data-testid="stMarkdownContainer"], [data-testid="stCaptionContainer"] {{color:#E5E7EB;}}
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5 {{color:#FFFFFF !important;}}
[data-testid="stCaptionContainer"] {{color:#CBD5E1 !important;}}
[data-testid="stHeader"] {{background:transparent;}}
.block-container {{padding-top:1.4rem; max-width:1500px;}}
/* hero */
.hero {{background:linear-gradient(120deg,#111C3A 0%,#1B1146 55%,#3B0F52 100%); border:1px solid #2B3A63;
  border-radius:18px; padding:22px 28px; margin-bottom:14px;}}
.hero-title {{font-size:2.5rem; font-weight:800; letter-spacing:.5px; color:#FFFFFF; line-height:1.1;}}
.hero-title span {{background:linear-gradient(90deg,#38BDF8,#D946EF); -webkit-background-clip:text;
  -webkit-text-fill-color:transparent;}}
.hero-tag {{color:#CBD5E1; font-size:1.05rem; margin-top:6px;}}
.hero-chip {{display:inline-block; margin-top:10px; padding:3px 12px; border-radius:999px; background:#0B1220;
  border:1px solid #334155; color:#CBD5E1; font-size:.78rem;}}
/* cards */
.card {{background:#162033; border:1px solid #26344D; border-radius:14px; padding:16px 18px; color:#F1F5F9; margin-bottom:10px;}}
.card-title {{font-size:.8rem; font-weight:700; letter-spacing:.9px; text-transform:uppercase; color:#93C5FD; margin-bottom:6px;}}
.card-text {{font-size:1rem; color:#F8FAFC; line-height:1.55;}}
.kpi {{background:#162033; border:1px solid #26344D; border-left:6px solid var(--accent,#38BDF8); border-radius:14px; padding:14px 18px;}}
.kpi-label {{font-size:.78rem; font-weight:700; letter-spacing:.8px; text-transform:uppercase; color:#CBD5E1;}}
.kpi-value {{font-size:2.15rem; font-weight:800; color:#FFFFFF; line-height:1.15;}}
.kpi-sub {{font-size:.8rem; color:#CBD5E1;}}
.sec-title {{font-size:1.35rem; font-weight:800; color:#FFFFFF; margin:6px 0 2px 0;}}
.sec-sub {{color:#CBD5E1; font-size:.95rem; margin-bottom:10px;}}
.step {{display:inline-block; background:{CTA_A}; color:#FFFFFF; border-radius:999px; padding:1px 10px; font-weight:800; font-size:.8rem; margin-right:8px;}}
.guide {{background:#0F2A3F; border:1px solid #1D5E8A; border-radius:12px; padding:10px 16px; color:#E0F2FE; margin:6px 0 12px 0;}}
.active-strip {{background:#14213A; border:1px solid #2F4A7D; border-radius:12px; padding:10px 16px; margin-bottom:12px; color:#F8FAFC;}}
.badge {{display:inline-block; padding:3px 12px; border-radius:999px; font-weight:800; font-size:.82rem; margin:2px 6px 2px 0;}}
.chip {{display:inline-block; padding:4px 12px; border-radius:8px; font-weight:700; font-size:.88rem; margin:3px 6px 3px 0;}}
.chip-ok {{background:#14532D; color:#DCFCE7; border:1px solid #22C55E;}}
.chip-bad {{background:#7F1D1D; color:#FEE2E2; border:1px solid #EF4444;}}
.chip-neutral {{background:#1E3A5F; color:#E0F2FE; border:1px solid #38BDF8;}}
.msgbox {{background:#0B1220; border:1px dashed #475569; border-radius:10px; padding:10px 14px; color:#E2E8F0; font-size:.9rem;}}
/* call-to-action buttons: vivid violet->magenta so they stand out from everything else */
button[data-testid^="stBaseButton-secondary"], button[data-testid^="stBaseButton-primary"],
a[data-testid^="stBaseLinkButton"], [data-testid="stFormSubmitButton"] button, .stDownloadButton button {{
  background:linear-gradient(90deg,{CTA_A},{CTA_B}) !important; color:#FFFFFF !important; border:0 !important;
  font-weight:800 !important; border-radius:10px !important; padding:.5rem 1.1rem !important;
  box-shadow:0 4px 18px rgba(217,70,239,.35);}}
button[data-testid^="stBaseButton-secondary"] p, button[data-testid^="stBaseButton-primary"] p,
a[data-testid^="stBaseLinkButton"] p, .stDownloadButton button p {{color:#FFFFFF !important; font-weight:800 !important;}}
button[data-testid^="stBaseButton-secondary"]:hover, button[data-testid^="stBaseButton-primary"]:hover,
a[data-testid^="stBaseLinkButton"]:hover, .stDownloadButton button:hover {{filter:brightness(1.15); transform:translateY(-1px);}}
/* tabs */
button[data-baseweb="tab"] {{font-weight:700 !important; font-size:1rem !important;}}
button[data-baseweb="tab"] p {{color:#CBD5E1 !important;}}
button[data-baseweb="tab"][aria-selected="true"] p {{color:#FFFFFF !important;}}
div[data-baseweb="tab-highlight"] {{background:{CTA_B} !important; height:3px !important;}}
/* inputs */
.stApp input, .stApp textarea {{color:#F8FAFC !important;}}
div[data-baseweb="input"], div[data-baseweb="textarea"], div[data-baseweb="select"] > div {{background:#1B2538 !important; border-color:#334155 !important;}}
div[data-baseweb="select"] * {{color:#F8FAFC !important;}}
div[data-baseweb="popover"] ul {{background:#1B2538 !important;}}
div[data-baseweb="popover"] li {{color:#F8FAFC !important;}}
[data-testid="stExpander"] {{border:1px solid #26344D; border-radius:12px; background:#111827;}}
hr {{border-color:#26344D !important;}}
</style>
"""


def _h(markup: str) -> str:
    """Strip indentation so Markdown never turns our HTML into a code block."""
    return "".join(line.strip() for line in markup.splitlines())


def esc(x) -> str:
    return html.escape(str(x))


# ─────────────────────────────────────────────────────────────────────────────
# SMALL UI HELPERS
# ─────────────────────────────────────────────────────────────────────────────
_chart_counter = {"n": 0}


def show_fig(fig):
    _chart_counter["n"] += 1
    key = f"chart_{_chart_counter['n']}"
    try:
        st.plotly_chart(fig, width="stretch", key=key)
    except Exception:
        st.plotly_chart(fig, use_container_width=True, key=key)


def show_df(data, **kw):
    try:
        return st.dataframe(data, width="stretch", **kw)
    except Exception:
        kw.pop("on_select", None)
        kw.pop("selection_mode", None)
        return st.dataframe(data, use_container_width=True, **kw)


def style_fig(fig, height=340):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#E5E7EB", size=13),
        margin=dict(l=10, r=10, t=45, b=10),
        height=height,
        title_font=dict(color="#FFFFFF", size=16),
        legend=dict(font=dict(color="#E5E7EB")),
    )
    fig.update_xaxes(gridcolor="#1F2A44", zerolinecolor="#1F2A44")
    fig.update_yaxes(gridcolor="#1F2A44", zerolinecolor="#1F2A44")
    return fig


def badge(label, kind="risk") -> str:
    palette = BURN_COLORS if kind == "burn" else RISK_COLORS
    bg = palette.get(str(label), "#64748B")
    fg = "#FFFFFF" if bg in ("#EF4444", "#64748B") else "#0B1220"
    return f'<span class="badge" style="background:{bg};color:{fg};">{esc(label)}</span>'


def kpi(label, value, sub="", accent="#38BDF8") -> str:
    return _h(
        f"""<div class="kpi" style="--accent:{accent}">
        <div class="kpi-label">{esc(label)}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div></div>"""
    )


def card(title, body_html) -> str:
    return _h(f'<div class="card"><div class="card-title">{esc(title)}</div><div class="card-text">{body_html}</div></div>')


def section(title, sub=""):
    st.markdown(_h(f'<div class="sec-title">{title}</div><div class="sec-sub">{sub}</div>'), unsafe_allow_html=True)


def guide(step, text):
    st.markdown(_h(f'<div class="guide"><span class="step">{step}</span>{text}</div>'), unsafe_allow_html=True)


def gauge(value, title, steps, rng=(0, 100), ref=None, suffix=""):
    gauge_cfg = {
        "axis": {"range": list(rng), "tickcolor": "#94A3B8", "tickfont": {"color": "#CBD5E1"}},
        "bar": {"color": "#38BDF8", "thickness": 0.28},
        "bgcolor": "#111827",
        "borderwidth": 0,
        "steps": [{"range": [a, b], "color": c} for a, b, c in steps],
    }
    if ref is not None:
        gauge_cfg["threshold"] = {"line": {"color": "#FFFFFF", "width": 3}, "thickness": 0.85, "value": ref}
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=float(value),
            number={"suffix": suffix, "font": {"color": "#FFFFFF", "size": 46}},
            title={"text": title, "font": {"color": "#E2E8F0", "size": 15}},
            gauge=gauge_cfg,
        )
    )
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#E5E7EB"), height=270, margin=dict(l=25, r=25, t=60, b=10)
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# DATA LAYER  (real file -> robust normalise -> mock fallback)
# ─────────────────────────────────────────────────────────────────────────────
SCHEMA = {  # column -> default if missing
    "Student_ID": "", "Student_Name": "", "Department": "Unknown", "Gender": "Unknown", "Semester": 1,
    "Success_Score": 50.0, "Academic_Risk": "Moderate", "Placement_Risk": "Moderate", "Overall_Risk": "Moderate",
    "Behavioral_Persona": "Balanced", "AI_Risk_Explanation": "No explanation available.",
    "Current_Skills": "", "Target_Role": "Not set", "Required_Skills": "", "Missing_Skills": "No major gap",
    "Skill_Gap_Level": "Moderate", "Skill_Gap_Score": 30.0, "Previous_Attendance_Pct": 75.0, "Attendance_Pct": 75.0,
    "Attendance_Drop_Pct": 0.0, "Backlogs": 0, "LMS_Weekly_Logins": 5.0, "Academic_Burnout_Stress_Risk_Index": 20.0,
    "Academic_Burnout_Stress_Risk": "Low", "Outreach_Message_Draft": "", "Parent_Nudge_Draft": "",
    "Peer_Buddy_ID": "", "Peer_Buddy_Name": "Not assigned", "Peer_Match_Score": 0.0,
    "What_If_Attendance_80_Score": 50.0, "What_If_Score_Plus_10": 60.0, "What_If_Improvement_Potential": 10,
    "Estimated_Recovery_Time_Weeks": 4, "Recommended_Action_Plan": "Continue regular mentoring",
    "Recommended_Faculty_Role": "Faculty Advisor", "Intervention_Priority": "P3 - Moderate",
    "Intervention_Status": "Pending", "Report_Eligible": "No",
}
NUMERIC = [
    "Semester", "Success_Score", "Skill_Gap_Score", "Previous_Attendance_Pct", "Attendance_Pct", "Attendance_Drop_Pct",
    "Backlogs", "LMS_Weekly_Logins", "Academic_Burnout_Stress_Risk_Index", "Peer_Match_Score",
    "What_If_Attendance_80_Score", "What_If_Score_Plus_10", "What_If_Improvement_Potential", "Estimated_Recovery_Time_Weeks",
]


def make_mock(n=1000, seed=7) -> pd.DataFrame:
    """Same schema as the real workbook - used only when the file cannot be read."""
    rng = np.random.default_rng(seed)
    first = ["Aarav", "Vihaan", "Aditya", "Ananya", "Diya", "Isha", "Rahul", "Priya", "Tanya", "Aman", "Neha", "Kabir"]
    last = ["Sharma", "Kumar", "Bansal", "Gupta", "Mehta", "Nair", "Verma", "Singh", "Reddy", "Iyer"]
    roles = {
        "Data Analyst": "SQL, Excel, Power BI, Python", "Software Developer": "DSA, Python/Java, Git, SQL",
        "Full Stack Developer": "HTML, CSS, JavaScript, React, Node.js", "Data Scientist": "Python, Statistics, SQL, ML",
        "ML Engineer": "Python, ML, SQL, Git", "Cloud Engineer": "Linux, Networking, AWS/Azure, Docker",
    }
    cur = ["Java, Spring", "Python, Statistics", "HTML, CSS, JavaScript", "C++, DSA", "Python, DSA", "SQL, Power BI", "Python, SQL"]
    ids = [f"STU{1001 + i}" for i in range(n)]
    dept = rng.choice(["CSE", "IT", "AIML", "ECE", "DS"], n, p=[.35, .18, .18, .17, .12])
    prev = rng.uniform(45, 100, n).round(1)
    drop = rng.uniform(0, 22, n).round(1)
    att = np.clip(prev - drop, 45, 98).round(1)
    bk = rng.choice([0, 1, 2, 3, 4, 5], n, p=[.32, .36, .2, .08, .03, .01])
    lms = rng.uniform(1, 12, n).round(1)
    gap = rng.normal(33, 8.6, n).clip(9, 60).round(1)
    score = (65.34 + 0.28 * att - 4 * bk - 0.64 * gap).clip(0, 100).round(1)
    overall = np.select([score < 45, score < 58, score < 66], ["Critical", "High", "Moderate"], "Safe")
    acad = np.select([(att < 55) & (bk >= 4), (att < 62) | (bk >= 3), (att < 75) | (bk >= 1)], ["Critical", "High", "Moderate"], "Safe")
    plac = np.select([gap > 52, gap > 40, gap > 25], ["Critical", "High", "Moderate"], "Safe")
    bidx = np.clip(1.1 * drop + 6 * bk + rng.normal(0, 4, n), 0, 100).round(1)
    blab = np.select([bidx >= 55, bidx >= 35], ["High", "Moderate"], "Low")
    tr = rng.choice(list(roles), n)
    miss = [", ".join(rng.choice(COMPETENCIES, rng.integers(0, 3), replace=False)) or "No major gap" for _ in range(n)]
    names = [f"{rng.choice(first)} {rng.choice(last)}" + (f" {1001 + i}" if rng.random() < .6 else "") for i in range(n)]
    buddy = []
    for i in range(n):
        pool = np.where((dept == dept[i]) & (np.arange(n) != i))[0]
        buddy.append(int(rng.choice(pool)))
    pri = np.select([overall == "Critical", overall == "High", overall == "Moderate"], ["P1 - Critical", "P2 - High", "P3 - Moderate"], "P4 - Routine")
    role_map = {"Critical": "Academic Mentor / HOD", "High": "Faculty Advisor", "Moderate": "Faculty Advisor", "Safe": "Faculty Advisor"}
    df = pd.DataFrame({
        "Student_ID": ids, "Student_Name": names, "Department": dept,
        "Gender": rng.permutation(["Male", "Female"] * (n // 2)), "Semester": rng.integers(1, 9, n),
        "Success_Score": score, "Academic_Risk": acad, "Placement_Risk": plac, "Overall_Risk": overall,
        "Behavioral_Persona": np.where(att < 60, "Low Attendance", np.where(gap > 40, "Placement Risk", "Balanced")),
        "AI_Risk_Explanation": [f"Primary factors: attendance {a}%, backlogs {b}. Overall risk is {o}." for a, b, o in zip(att, bk, overall)],
        "Current_Skills": rng.choice(cur, n), "Target_Role": tr, "Required_Skills": [roles[t] for t in tr],
        "Missing_Skills": miss, "Skill_Gap_Level": np.select([gap > 52, gap > 30], ["High", "Moderate"], "Low"),
        "Skill_Gap_Score": gap, "Previous_Attendance_Pct": prev, "Attendance_Pct": att, "Attendance_Drop_Pct": drop,
        "Backlogs": bk, "LMS_Weekly_Logins": lms, "Academic_Burnout_Stress_Risk_Index": bidx,
        "Academic_Burnout_Stress_Risk": blab,
        "Outreach_Message_Draft": np.where(
            overall == "Safe", "Progress message: continue current academic and placement preparation.",
            "Supportive outreach: discuss low attendance and create a short improvement plan."),
        "Parent_Nudge_Draft": np.where(
            overall == "Critical", "Suggested supportive parent/guardian progress update with advisor involvement.",
            np.where(overall == "High", "Optional parent/guardian progress update after advisor review.", "No parent nudge required.")),
        "Peer_Buddy_ID": [ids[b] for b in buddy], "Peer_Buddy_Name": [names[b] for b in buddy],
        "Peer_Match_Score": rng.uniform(20, 98, n).round(1),
        "What_If_Attendance_80_Score": (score + np.clip(80 - att, 0, None) * 0.2).round(1),
        "What_If_Score_Plus_10": (score + 10).round(1), "What_If_Improvement_Potential": 10,
        "Estimated_Recovery_Time_Weeks": rng.integers(2, 10, n),
        "Recommended_Action_Plan": np.where(overall == "Safe", "Continue regular mentoring", "Attendance improvement plan"),
        "Recommended_Faculty_Role": [role_map[o] for o in overall], "Intervention_Priority": pri,
        "Intervention_Status": np.where(overall == "Safe", "Completed", np.where(overall == "High", "In Progress", "Pending")),
        "Report_Eligible": np.where(np.isin(overall, ["High", "Critical"]), "Yes", "No"),
    })
    return df


def normalise(df: pd.DataFrame) -> pd.DataFrame:
    """Guarantee every expected column exists with sane types so no page can KeyError."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for col, default in SCHEMA.items():
        if col not in df.columns:
            df[col] = default
    for col in NUMERIC:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(SCHEMA[col])
    for col in df.columns:
        if col not in NUMERIC:
            df[col] = df[col].fillna("").astype(str).str.strip()
    df["Semester"] = df["Semester"].astype(int)
    df["Backlogs"] = df["Backlogs"].astype(int)
    df = df[df["Student_ID"] != ""].drop_duplicates("Student_ID").reset_index(drop=True)
    return df


def find_data_file():
    here = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    for p in [FILE_NAME, os.path.join(here, FILE_NAME), os.path.join(here, "data", FILE_NAME),
              os.path.join("/mnt/user-data/uploads", FILE_NAME)]:
        if os.path.exists(p):
            return p
    return None


@st.cache_data(show_spinner=False)
def read_excel_cached(src):
    """src is a file path (str) or raw bytes (uploaded file)."""
    handle = io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src
    try:
        return pd.read_excel(handle, sheet_name=SHEET_NAME)
    except ValueError:  # sheet name differs -> take the first sheet
        if isinstance(src, (bytes, bytearray)):
            handle = io.BytesIO(src)
        return pd.read_excel(handle, sheet_name=0)


def load_data(uploaded):
    try:
        if uploaded is not None:
            return normalise(read_excel_cached(uploaded.getvalue())), "upload", ""
        path = find_data_file()
        if path is None:
            raise FileNotFoundError(f"'{FILE_NAME}' not found next to app.py")
        return normalise(read_excel_cached(path)), "file", ""
    except Exception as exc:  # never crash the demo
        return normalise(make_mock()), "mock", str(exc)


@st.cache_data(show_spinner=False)
def fit_success_model(df: pd.DataFrame):
    """Success_Score is (almost) linear in attendance, backlogs and skill gap - fit it for the What-If tool."""
    X = np.c_[np.ones(len(df)), df["Attendance_Pct"], df["Backlogs"], df["Skill_Gap_Score"]]
    y = df["Success_Score"].values
    coef, *_ = np.linalg.lstsq(X, y, rcond=None)
    ss_res = ((y - X @ coef) ** 2).sum()
    ss_tot = ((y - y.mean()) ** 2).sum() or 1.0
    return coef, float(1 - ss_res / ss_tot)


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE: intervention tracker + active student
# ─────────────────────────────────────────────────────────────────────────────
def init_state(df: pd.DataFrame, source: str):
    sig = f"{source}-{len(df)}-{df['Student_ID'].iloc[0]}"
    if st.session_state.get("data_sig") != sig:
        t = df[["Student_ID", "Student_Name", "Department", "Overall_Risk", "Intervention_Priority",
                "Recommended_Faculty_Role", "Intervention_Status"]].copy()
        t = t.rename(columns={"Recommended_Faculty_Role": "Assigned_Faculty_Role"})
        t["Notes"] = ""
        t["Last_Updated"] = ""
        t = t.set_index("Student_ID", drop=False)
        t.index.name = None
        st.session_state.tracker = t
        st.session_state.tracker_ver = 0
        st.session_state.data_sig = sig
        crit = df[df["Overall_Risk"] == "Critical"]
        st.session_state.active_id = (crit if len(crit) else df.sort_values("Success_Score")).iloc[0]["Student_ID"]


def log_action(sid, note, status=None, role=None):
    t = st.session_state.tracker
    stamp = datetime.now().strftime("%d %b %H:%M")
    old = t.at[sid, "Notes"]
    if note:
        t.at[sid, "Notes"] = (old + " | " if old else "") + f"[{stamp}] {note}"
    if status:
        t.at[sid, "Intervention_Status"] = status
    if role:
        t.at[sid, "Assigned_Faculty_Role"] = role
    t.at[sid, "Last_Updated"] = stamp
    st.session_state.tracker_ver += 1


def on_directory_select():
    key = st.session_state.get("_dir_key")
    state = st.session_state.get(key)
    try:
        rows = state["selection"]["rows"]
    except Exception:
        rows = []
    ids = st.session_state.get("_dir_ids", [])
    if rows and rows[0] < len(ids):
        st.session_state["active_id"] = ids[rows[0]]


# ─────────────────────────────────────────────────────────────────────────────
# SKILL / MESSAGE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def split_list(s):
    return [x.strip() for x in str(s).split(",") if x.strip()]


def skill_match(required, current):
    cur = {c.lower() for c in current}
    return {r: any(alt.strip().lower() in cur for alt in r.split("/")) for r in required}


def topics_from_draft(draft):
    m = re.search(r"discuss (.+?) and create", str(draft))
    return m.group(1) if m else ""


def build_message(s, kind, channel):
    first = s["Student_Name"].split()[0]
    topics = topics_from_draft(s["Outreach_Message_Draft"])
    sign = f"{s['Recommended_Faculty_Role']}, {s['Department']} Department\nEduNexus AI Student Success Cell"
    if kind == "student":
        if channel == "SMS":
            body = (f"Hi {first}, your {s['Recommended_Faculty_Role']} would like a quick supportive catch-up about "
                    f"{topics or 'your progress'}. Please reply to book a slot. - EduNexus AI")
            return "", body
        subject = (f"Let's plan your next steps, {first}" if topics else f"Great progress, {first} - keep it going")
        if topics:
            core = (f"Your {s['Department']} faculty team has been reviewing your progress, and we would like to set up a short, "
                    f"supportive conversation to discuss {topics}. Together we will create a simple improvement plan "
                    f"(estimated recovery: ~{s['Estimated_Recovery_Time_Weeks']} weeks).")
        else:
            core = "Your recent progress looks steady. Please continue your current academic and placement preparation."
        body = f"Dear {s['Student_Name']},\n\n{core}\n\nSuggested next step: {s['Recommended_Action_Plan']}.\nYour peer buddy: {s['Peer_Buddy_Name']}.\n\nRegards,\n{sign}"
        return subject, body
    # parent
    if channel == "SMS":
        body = (f"Dear Parent/Guardian, this is a progress update about {first} from {s['Department']} Dept. "
                f"We are working with {first} on {topics or 'continued progress'}. Please feel free to call the advisor. - EduNexus AI")
        return "", body
    subject = f"Progress update for {s['Student_Name']} - {s['Department']} Department"
    focus = f"additional support in the areas of {topics}" if topics else "steady progress"
    body = (f"Dear Parent/Guardian of {s['Student_Name']},\n\nWe are writing to share a supportive update. "
            f"{first} is currently showing {focus}. Our faculty team has a plan in place ({s['Recommended_Action_Plan']}) "
            f"and we would appreciate your encouragement at home.\n\nPlease reach out to the advisor if you would like to discuss this.\n\nWarm regards,\n{sign}")
    return subject, body


# ═════════════════════════════════════════════════════════════════════════════
# APP START
# ═════════════════════════════════════════════════════════════════════════════
st.markdown(CSS, unsafe_allow_html=True)

with st.sidebar:
    st.markdown(_h('<div class="hero-title" style="font-size:1.6rem;">Edu<span>Nexus</span> AI</div>'
                   '<div class="hero-tag" style="font-size:.85rem;">Student Success &amp; Intelligence Hub</div>'), unsafe_allow_html=True)
    st.markdown("---")
    with st.expander("📂 Data source", expanded=False):
        uploaded = st.file_uploader("Upload the Student360 .xlsx (optional)", type=["xlsx"])

df, source, err = load_data(uploaded)
init_state(df, source)
by_id = df.set_index("Student_ID", drop=False)
by_id.index.name = None
tracker = st.session_state.tracker
id_list = df["Student_ID"].tolist()

with st.sidebar:
    st.markdown("### 🎯 Active student")
    st.caption("Every tab below shows THIS student. Type a name or ID to search, or click a row in the Student 360 directory.")
    st.selectbox(
        "Select student", id_list, key="active_id",
        format_func=lambda sid: f"{sid} — {by_id.at[sid, 'Student_Name']}",
    )
    st.markdown("---")
    st.markdown("### 🧭 How to use")
    st.markdown(_h(
        '<div style="color:#E5E7EB;font-size:.9rem;line-height:1.8;">'
        '<b>1.</b> Command Center → see the big picture<br><b>2.</b> Student 360 → find &amp; open a student<br>'
        '<b>3.</b> Skill Gap / Burnout → diagnose<br><b>4.</b> AI Action Copilot → act &amp; track</div>'), unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### 🎨 Risk legend")
    st.markdown("".join(badge(r) for r in RISK_ORDER), unsafe_allow_html=True)
    st.caption("Critical = red · High = orange · Moderate = amber · Safe = green")

active_id = st.session_state["active_id"]
S = by_id.loc[active_id]

# ── Hero ─────────────────────────────────────────────────────────────────────
src_label = {"file": "Live data: Excel workbook", "upload": "Live data: uploaded workbook", "mock": "⚠ Demo mock data"}[source]
st.markdown(_h(f"""<div class="hero"><div class="hero-title">🎓 Edu<span>Nexus</span> AI</div>
    <div class="hero-tag">The Unified Student Success &amp; Intelligence Hub</div>
    <div class="hero-chip">{src_label} · {len(df):,} students · {df['Department'].nunique()} departments</div></div>"""),
    unsafe_allow_html=True)
if source == "mock":
    st.warning(f"Could not read the workbook ({err}). Showing realistic mock data with the same schema so the demo keeps working. "
               "Place the .xlsx next to app.py or upload it from the sidebar.")


def active_strip():
    st.markdown(_h(f"""<div class="active-strip">👤 <b style="color:#FFFFFF;font-size:1.05rem;">{esc(S['Student_Name'])}</b>
        &nbsp;<span style="color:#CBD5E1;">{esc(S['Student_ID'])} · {esc(S['Department'])} · Semester {S['Semester']} · {esc(S['Gender'])}</span>
        &nbsp;&nbsp;{badge(S['Overall_Risk'])}
        <span style="color:#CBD5E1;font-size:.85rem;">&nbsp;Change student: sidebar ◀ or Student 360 directory</span></div>"""),
        unsafe_allow_html=True)


tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🏠 Executive Command Center", "🎓 Student 360 Directory", "💼 Placement Skill Gap",
    "🧠 Burnout & Wellness", "🤖 AI Action Copilot",
])

# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 - EXECUTIVE COMMAND CENTER
# ═════════════════════════════════════════════════════════════════════════════
with tab1:
    section("Executive Command Center", "Institution-wide view of student success, risk and intervention progress.")
    total = len(df)
    n_crit = int((df["Overall_Risk"] == "Critical").sum())
    n_high = int((df["Overall_Risk"] == "High").sum())
    done_pct = (tracker["Intervention_Status"] == "Completed").mean() * 100
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.markdown(kpi("Total Students", f"{total:,}", f"{df['Department'].nunique()} departments", "#38BDF8"), unsafe_allow_html=True)
    c2.markdown(kpi("Avg Success Score", f"{df['Success_Score'].mean():.1f}", "out of 100", "#A78BFA"), unsafe_allow_html=True)
    c3.markdown(kpi("Critical Risk", f"{n_crit}", "needs action today", "#EF4444"), unsafe_allow_html=True)
    c4.markdown(kpi("High Risk", f"{n_high}", "faculty follow-up", "#F97316"), unsafe_allow_html=True)
    c5.markdown(kpi("Interventions Done", f"{done_pct:.0f}%", "live from tracker", "#22C55E"), unsafe_allow_html=True)
    st.write("")

    # ── Export (one click) ──
    at_risk = df[df["Overall_Risk"].isin(["Critical", "High"])].copy()
    at_risk["_r"] = at_risk["Overall_Risk"].map(RISK_RANK)
    at_risk = at_risk.sort_values(["_r", "Success_Score"]).drop(columns="_r")
    status_now = tracker["Intervention_Status"].reindex(at_risk["Student_ID"]).values
    export_cols = ["Student_ID", "Student_Name", "Department", "Gender", "Semester", "Success_Score", "Overall_Risk",
                   "Academic_Risk", "Placement_Risk", "Attendance_Pct", "Backlogs", "Behavioral_Persona",
                   "AI_Risk_Explanation", "Recommended_Action_Plan", "Recommended_Faculty_Role", "Intervention_Priority"]
    export_df = at_risk[export_cols].copy()
    export_df["Intervention_Status"] = status_now
    csv_bytes = export_df.to_csv(index=False).encode("utf-8")
    xbuf = io.BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="At_Risk_Students", index=False)
    stamp = datetime.now().strftime("%Y%m%d")
    st.markdown(card("📤 Instant report for faculty meetings",
                     f"<b>{len(export_df)}</b> at-risk students (Critical + High), sorted worst-first, with AI explanation, "
                     "action plan and live intervention status. One click to download."), unsafe_allow_html=True)
    e1, e2, _ = st.columns([1, 1, 2])
    e1.download_button("⬇ Download At-Risk CSV", csv_bytes, f"EduNexus_at_risk_{stamp}.csv", "text/csv", key="dl_csv")
    e2.download_button("⬇ Download At-Risk Excel", xbuf.getvalue(), f"EduNexus_at_risk_{stamp}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")
    st.write("")

    # ── Charts row 1 ──
    r1, r2 = st.columns([1, 1.4])
    with r1:
        rc = df["Overall_Risk"].value_counts().reindex(RISK_ORDER).dropna().reset_index()
        rc.columns = ["Risk", "Students"]
        fig = px.pie(rc, names="Risk", values="Students", hole=0.55, color="Risk", color_discrete_map=RISK_COLORS,
                     category_orders={"Risk": RISK_ORDER}, title="Risk Distribution")
        fig.update_traces(textposition="outside", textinfo="label+percent", textfont=dict(color="#F8FAFC", size=13),
                          marker=dict(line=dict(color="#0E1117", width=2)), sort=False)
        fig.update_layout(showlegend=False)
        fig.add_annotation(text=f"<b>{total:,}</b><br>students", showarrow=False, font=dict(size=16, color="#FFFFFF"))
        show_fig(style_fig(fig, 360))
    with r2:
        dr = df.groupby(["Department", "Overall_Risk"]).size().reset_index(name="Students")
        fig = px.bar(dr, x="Department", y="Students", color="Overall_Risk", color_discrete_map=RISK_COLORS,
                     category_orders={"Overall_Risk": RISK_ORDER}, title="Department-wise Risk Breakdown", text="Students")
        fig.update_traces(textfont=dict(color="#0B1220"))
        fig.update_layout(barmode="stack", legend_title_text="Risk")
        show_fig(style_fig(fig, 360))

    r3, r4 = st.columns(2)
    with r3:
        gr = df.groupby(["Gender", "Overall_Risk"]).size().reset_index(name="Students")
        gr["Share"] = gr["Students"] / gr.groupby("Gender")["Students"].transform("sum") * 100
        fig = px.bar(gr, x="Overall_Risk", y="Share", color="Gender", barmode="group", title="Gender-wise Risk Breakdown (% of gender)",
                     category_orders={"Overall_Risk": RISK_ORDER}, text=gr["Share"].round(1).astype(str) + "%",
                     color_discrete_map={"Male": "#38BDF8", "Female": "#F472B6"})
        fig.update_traces(textposition="outside", textfont=dict(color="#F8FAFC"))
        fig.update_yaxes(title="% of students")
        fig.update_xaxes(title="")
        show_fig(style_fig(fig, 340))
    with r4:
        pr = df["Behavioral_Persona"].value_counts().reset_index()
        pr.columns = ["Persona", "Students"]
        fig = px.bar(pr.sort_values("Students"), x="Students", y="Persona", orientation="h", text="Students",
                     title="Behavioral Personas", color_discrete_sequence=["#A78BFA"])
        fig.update_traces(textposition="outside", textfont=dict(color="#F8FAFC"))
        fig.update_yaxes(title="")
        show_fig(style_fig(fig, 340))

    r5, r6 = st.columns([1, 1.4])
    with r5:
        sc = tracker["Intervention_Status"].value_counts().reindex(STATUS_ORDER).fillna(0).reset_index()
        sc.columns = ["Status", "Students"]
        fig = px.pie(sc, names="Status", values="Students", hole=0.55, color="Status", color_discrete_map=STATUS_COLORS,
                     title="Intervention Status (live)")
        fig.update_traces(textposition="outside", textinfo="label+value", textfont=dict(color="#F8FAFC", size=13),
                          marker=dict(line=dict(color="#0E1117", width=2)), sort=False)
        fig.update_layout(showlegend=False)
        show_fig(style_fig(fig, 320))
    with r6:
        fig = px.histogram(df, x="Success_Score", nbins=30, color="Overall_Risk", color_discrete_map=RISK_COLORS,
                           category_orders={"Overall_Risk": RISK_ORDER}, title="Success Score Distribution")
        fig.update_layout(barmode="stack", legend_title_text="Risk")
        show_fig(style_fig(fig, 320))

    section("🚨 Critical students — start here", "These students need action first. Open the Student 360 tab to see why.")
    crit_df = df[df["Overall_Risk"] == "Critical"].sort_values("Success_Score")
    show_df(crit_df[["Student_ID", "Student_Name", "Department", "Semester", "Success_Score", "Attendance_Pct", "Backlogs",
                     "AI_Risk_Explanation", "Recommended_Faculty_Role"]], hide_index=True, height=min(60 + 36 * len(crit_df), 420))

# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 - STUDENT 360 DIRECTORY & PROFILE
# ═════════════════════════════════════════════════════════════════════════════
with tab2:
    section("Student 360 Directory & Deep Dive", "Find a student, click their row, and the full profile opens below.")
    guide("STEP 1", "Search or filter the directory.  <b>STEP 2</b> Click a row's left checkbox.  <b>STEP 3</b> Scroll down to the 360 profile.")

    f1, f2, f3, f4, f5 = st.columns([2, 1, 1, 1, 1])
    q = f1.text_input("🔍 Search name, ID, role or skill", placeholder="e.g. Aarav, STU1042, Data Analyst, Python")
    f_dept = f2.multiselect("Department", sorted(df["Department"].unique()))
    f_gender = f3.multiselect("Gender", sorted(df["Gender"].unique()))
    f_risk = f4.multiselect("Overall risk", [r for r in RISK_ORDER if r in set(df["Overall_Risk"])])
    f_sem = f5.multiselect("Semester", sorted(df["Semester"].unique()))
    g1, g2 = st.columns([1, 1])
    f_persona = g1.multiselect("Behavioral persona", sorted(df["Behavioral_Persona"].unique()))
    f_status = g2.multiselect("Intervention status", STATUS_ORDER)

    view = df.copy()
    view["Intervention_Status"] = tracker["Intervention_Status"].reindex(view["Student_ID"]).values
    if q.strip():
        needle = q.strip().lower()
        hay = (view["Student_ID"] + " " + view["Student_Name"] + " " + view["Target_Role"] + " " + view["Current_Skills"]).str.lower()
        view = view[hay.str.contains(re.escape(needle), na=False)]
    for col, vals in [("Department", f_dept), ("Gender", f_gender), ("Overall_Risk", f_risk), ("Semester", f_sem),
                      ("Behavioral_Persona", f_persona), ("Intervention_Status", f_status)]:
        if vals:
            view = view[view[col].isin(vals)]
    view = view.assign(_r=view["Overall_Risk"].map(RISK_RANK)).sort_values(["_r", "Success_Score"]).drop(columns="_r")

    dir_cols = ["Student_ID", "Student_Name", "Department", "Gender", "Semester", "Success_Score", "Overall_Risk",
                "Attendance_Pct", "Backlogs", "Behavioral_Persona", "Target_Role", "Intervention_Status"]
    table = view[dir_cols].reset_index(drop=True)
    st.caption(f"Showing **{len(table):,}** of **{len(df):,}** students · sorted by risk (worst first) · click any header to re-sort")

    sig = f"{q}|{f_dept}|{f_gender}|{f_risk}|{f_sem}|{f_persona}|{f_status}"
    dir_key = f"dir_{abs(hash(sig))}"
    st.session_state["_dir_key"] = dir_key
    st.session_state["_dir_ids"] = table["Student_ID"].tolist()
    show_df(
        table, hide_index=True, height=340, key=dir_key, on_select=on_directory_select, selection_mode="single-row",
        column_config={
            "Success_Score": st.column_config.ProgressColumn("Success Score", min_value=0, max_value=100, format="%.1f"),
            "Attendance_Pct": st.column_config.NumberColumn("Attendance %", format="%.1f"),
            "Overall_Risk": st.column_config.TextColumn("Overall Risk"),
        },
    )
    st.download_button("⬇ Export this filtered list (CSV)", table.to_csv(index=False).encode("utf-8"),
                       "EduNexus_directory_filtered.csv", "text/csv", key="dl_dir")

    st.markdown("---")
    section(f"👤 360 Profile — {esc(S['Student_Name'])}", f"{esc(S['Student_ID'])} · {esc(S['Department'])} · Semester {S['Semester']} · {esc(S['Gender'])}")

    p1, p2 = st.columns([1, 1.5])
    with p1:
        avg = df["Success_Score"].mean()
        fig = gauge(S["Success_Score"], "Success Score (0-100)",
                    [(0, 45, "rgba(239,68,68,.55)"), (45, 60, "rgba(249,115,22,.55)"),
                     (60, 75, "rgba(250,204,21,.55)"), (75, 100, "rgba(34,197,94,.55)")], ref=avg)
        show_fig(fig)
        st.caption(f"White line = cohort average ({avg:.1f}).  Student is **{S['Success_Score'] - avg:+.1f}** vs average.")
    with p2:
        st.markdown(card("Risk flags",
                         f"Overall {badge(S['Overall_Risk'])} Academic {badge(S['Academic_Risk'])} "
                         f"Placement {badge(S['Placement_Risk'])} Burnout {badge(S['Academic_Burnout_Stress_Risk'], 'burn')}"),
                    unsafe_allow_html=True)
        acad_lbl = {"Safe": "High Academic", "Moderate": "Moderate Academic"}.get(S["Academic_Risk"], "Low Academic")
        plac_lbl = {"Safe": "High Placement", "Moderate": "Moderate Placement"}.get(S["Placement_Risk"], "Low Placement")
        st.markdown(card("Behavioral persona",
                         f'<span class="chip chip-neutral">{esc(S["Behavioral_Persona"])}</span>'
                         f'<span class="chip chip-neutral">{acad_lbl} / {plac_lbl}</span>'),
                    unsafe_allow_html=True)
        st.markdown(card("🧠 Why is this student at risk? (AI explanation)", esc(S["AI_Risk_Explanation"])), unsafe_allow_html=True)

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.markdown(kpi("Attendance", f"{S['Attendance_Pct']:.1f}%", f"drop {S['Attendance_Drop_Pct']:.1f} pts", "#38BDF8"), unsafe_allow_html=True)
    m2.markdown(kpi("Backlogs", f"{S['Backlogs']}", "active", "#F97316" if S["Backlogs"] >= 2 else "#22C55E"), unsafe_allow_html=True)
    m3.markdown(kpi("LMS Logins / wk", f"{S['LMS_Weekly_Logins']:.1f}", f"cohort {df['LMS_Weekly_Logins'].mean():.1f}", "#A78BFA"), unsafe_allow_html=True)
    m4.markdown(kpi("Skill Gap", f"{S['Skill_Gap_Score']:.1f}", esc(S["Skill_Gap_Level"]) + " gap", "#FACC15"), unsafe_allow_html=True)
    m5.markdown(kpi("Priority", esc(S["Intervention_Priority"]), esc(tracker.at[active_id, "Intervention_Status"]), "#D946EF"), unsafe_allow_html=True)
    st.write("")
    a1, a2 = st.columns(2)
    a1.markdown(card("Recommended action plan", f"{esc(S['Recommended_Action_Plan'])}<br><span style='color:#CBD5E1'>Owner: {esc(S['Recommended_Faculty_Role'])} · "
                                                  f"Est. recovery: {S['Estimated_Recovery_Time_Weeks']} weeks</span>"), unsafe_allow_html=True)
    a2.markdown(card("Target role", f"{esc(S['Target_Role'])}<br><span style='color:#CBD5E1'>Current skills: {esc(S['Current_Skills'])}</span>"), unsafe_allow_html=True)
    guide("NEXT ➜", "Open <b>Placement Skill Gap</b> or <b>Burnout &amp; Wellness</b> to diagnose, then <b>AI Action Copilot</b> to send outreach and track the intervention.")

# ═════════════════════════════════════════════════════════════════════════════
# TAB 3 - PLACEMENT SKILL GAP
# ═════════════════════════════════════════════════════════════════════════════
with tab3:
    section("Placement Skill Gap Analysis", "Current skills vs what the target role requires.")
    active_strip()
    cur_skills = split_list(S["Current_Skills"])
    req_skills = split_list(S["Required_Skills"])
    match = skill_match(req_skills, cur_skills)
    n_have = sum(match.values())
    pct = n_have / len(req_skills) * 100 if req_skills else 0
    miss_items = [m for m in split_list(S["Missing_Skills"]) if m.lower() != "no major gap"]

    t1, t2, t3 = st.columns([1.2, 1.2, 1])
    with t1:
        st.markdown(card("🎯 Target role", f"<span style='font-size:1.3rem;font-weight:800;color:#FFFFFF'>{esc(S['Target_Role'])}</span>"),
                    unsafe_allow_html=True)
        st.markdown(card("✅ Current skills", "".join(f'<span class="chip chip-neutral">{esc(c)}</span>' for c in cur_skills) or "None recorded"),
                    unsafe_allow_html=True)
    with t2:
        chips = "".join(f'<span class="chip {"chip-ok" if ok else "chip-bad"}">{"✔" if ok else "✖"} {esc(r)}</span>' for r, ok in match.items())
        st.markdown(card("📋 Required skills (green = have, red = missing)", chips or "None"), unsafe_allow_html=True)
        gap_tech = [r for r, ok in match.items() if not ok]
        st.markdown(card("⚠ Technical skills to learn", "".join(f'<span class="chip chip-bad">{esc(g)}</span>' for g in gap_tech)
                         or '<span class="chip chip-ok">All required skills covered</span>'), unsafe_allow_html=True)
    with t3:
        fig = gauge(pct, "Technical skill match", [(0, 40, "rgba(239,68,68,.55)"), (40, 70, "rgba(250,204,21,.55)"), (70, 100, "rgba(34,197,94,.55)")], suffix="%")
        show_fig(fig)

    section("Readiness competencies", "Flagged by the placement model in <b>Missing_Skills</b>. "
                                      "Coding · Aptitude · Interview · Communication.")
    cc = st.columns(4)
    for col, comp in zip(cc, COMPETENCIES):
        flagged = comp in miss_items
        label = {"Coding": "Coding Score", "Aptitude": "Aptitude Score", "Interview": "Mock Interview", "Communication": "Communication"}[comp]
        col.markdown(kpi(label, "Gap flagged" if flagged else "On track", "needs training" if flagged else "meets benchmark",
                         "#EF4444" if flagged else "#22C55E"), unsafe_allow_html=True)
    st.caption("ℹ The workbook stores competency gaps as flags (no raw Coding / Aptitude / Mock-Interview numbers), so the cards show the flag status.")
    st.write("")

    x1, x2 = st.columns(2)
    with x1:
        role_avg = df.groupby("Target_Role")["Skill_Gap_Score"].mean().reset_index().sort_values("Skill_Gap_Score")
        role_avg["Mine"] = np.where(role_avg["Target_Role"] == S["Target_Role"], "Selected student's role", "Other roles")
        fig = px.bar(role_avg, x="Skill_Gap_Score", y="Target_Role", orientation="h", color="Mine", title="Average skill-gap score by target role",
                     color_discrete_map={"Selected student's role": "#D946EF", "Other roles": "#475569"}, text=role_avg["Skill_Gap_Score"].round(1))
        fig.add_vline(x=S["Skill_Gap_Score"], line_color="#38BDF8", line_dash="dash",
                      annotation_text=f"Student {S['Skill_Gap_Score']:.1f}", annotation_font_color="#38BDF8")
        fig.update_traces(textposition="outside", textfont=dict(color="#F8FAFC"))
        fig.update_layout(legend_title_text="")
        fig.update_yaxes(title="")
        show_fig(style_fig(fig, 340))
    with x2:
        dep = df[df["Department"] == S["Department"]]
        counts = {c: dep["Missing_Skills"].str.contains(c, na=False).mean() * 100 for c in COMPETENCIES}
        cdf = pd.DataFrame({"Competency": list(counts), "Students flagged %": list(counts.values())})
        cdf["Student"] = np.where(cdf["Competency"].isin(miss_items), "Flagged for this student", "Not flagged")
        fig = px.bar(cdf, x="Competency", y="Students flagged %", color="Student", title=f"Competency gaps across {S['Department']} department",
                     color_discrete_map={"Flagged for this student": "#EF4444", "Not flagged": "#38BDF8"},
                     text=cdf["Students flagged %"].round(0).astype(int).astype(str) + "%")
        fig.update_traces(textposition="outside", textfont=dict(color="#F8FAFC"))
        fig.update_layout(legend_title_text="")
        show_fig(style_fig(fig, 340))

    st.markdown(card("Placement summary",
                     f"Placement risk {badge(S['Placement_Risk'])} Skill gap level <b>{esc(S['Skill_Gap_Level'])}</b> "
                     f"(score {S['Skill_Gap_Score']:.1f}). Focus: "
                     f"<b>{esc(', '.join(miss_items)) if miss_items else 'maintain current preparation'}</b>.<br>"
                     f"Recommended action: {esc(S['Recommended_Action_Plan'])}"), unsafe_allow_html=True)
    guide("NEXT ➜", "Go to <b>AI Action Copilot → Outreach</b> to send the student a placement-prep plan.")

# ═════════════════════════════════════════════════════════════════════════════
# TAB 4 - BURNOUT & WELLNESS
# ═════════════════════════════════════════════════════════════════════════════
with tab4:
    section("Burnout & Wellness Intelligence",
            "Academic Burnout / Stress Risk Index — driven by sudden attendance drops and multiple backlogs.")
    active_strip()
    idx_col, lab_col = "Academic_Burnout_Stress_Risk_Index", "Academic_Burnout_Stress_Risk"
    mins = df.groupby(lab_col)[idx_col].min()
    mod_start = float(mins.get("Moderate", 30))
    high_start = float(mins.get("High", 50))
    if not mod_start < high_start:
        mod_start, high_start = 30.0, 50.0

    b1, b2 = st.columns([1, 1.2])
    with b1:
        fig = gauge(S[idx_col], "Burnout / Stress Risk Index",
                    [(0, mod_start, "rgba(34,197,94,.55)"), (mod_start, high_start, "rgba(250,204,21,.55)"), (high_start, 100, "rgba(239,68,68,.55)")],
                    ref=float(df[idx_col].mean()))
        show_fig(fig)
        st.markdown(f"Burnout level: {badge(S[lab_col], 'burn')} &nbsp; <span style='color:#CBD5E1'>White line = cohort average ({df[idx_col].mean():.1f})</span>", unsafe_allow_html=True)
    with b2:
        fig = go.Figure()
        fig.add_bar(x=["Previous attendance", "Current attendance"], y=[S["Previous_Attendance_Pct"], S["Attendance_Pct"]],
                    marker_color=["#64748B", "#38BDF8" if S["Attendance_Drop_Pct"] < 10 else "#EF4444"],
                    text=[f"{S['Previous_Attendance_Pct']:.1f}%", f"{S['Attendance_Pct']:.1f}%"], textposition="outside", textfont=dict(color="#F8FAFC"))
        fig.update_layout(title=f"Attendance trend  (drop: {S['Attendance_Drop_Pct']:.1f} pts)", showlegend=False)
        fig.update_yaxes(range=[0, 115], title="%")
        show_fig(style_fig(fig, 300))

    lms_mean, lms_q25 = df["LMS_Weekly_Logins"].mean(), df["LMS_Weekly_Logins"].quantile(0.25)
    dept_lms = df.loc[df["Department"] == S["Department"], "LMS_Weekly_Logins"].mean()
    pctile = (df["LMS_Weekly_Logins"] < S["LMS_Weekly_Logins"]).mean() * 100
    e1, e2, e3, e4 = st.columns(4)
    e1.markdown(kpi("LMS logins / week", f"{S['LMS_Weekly_Logins']:.1f}", "student", "#A78BFA"), unsafe_allow_html=True)
    e2.markdown(kpi("Dept average", f"{dept_lms:.1f}", S["Department"], "#38BDF8"), unsafe_allow_html=True)
    e3.markdown(kpi("Cohort average", f"{lms_mean:.1f}", "all students", "#38BDF8"), unsafe_allow_html=True)
    e4.markdown(kpi("Engagement percentile", f"{pctile:.0f}th", "LMS activity vs cohort", "#22C55E" if pctile >= 25 else "#EF4444"), unsafe_allow_html=True)
    st.write("")

    signals = []
    if S["Attendance_Drop_Pct"] >= 10:
        signals.append(f"📉 Attendance dropped <b>{S['Attendance_Drop_Pct']:.1f} points</b> ({S['Previous_Attendance_Pct']:.1f}% → {S['Attendance_Pct']:.1f}%)")
    if S["Backlogs"] >= 2:
        signals.append(f"📚 <b>{S['Backlogs']} active backlogs</b> — compounding academic pressure")
    if S["LMS_Weekly_Logins"] <= lms_q25:
        signals.append(f"💻 LMS logins ({S['LMS_Weekly_Logins']:.1f}/wk) are in the <b>bottom quartile</b> — disengagement signal")
    if S["Attendance_Pct"] < 65:
        signals.append(f"🏫 Attendance <b>{S['Attendance_Pct']:.1f}%</b> is below the 65% comfort line")
    sig_html = "<br>".join(signals) if signals else "✅ No acute stress signals detected for this student."
    w1, w2 = st.columns(2)
    w1.markdown(card("🔎 Wellness signals", sig_html), unsafe_allow_html=True)
    w2.markdown(card("💬 Suggested support", f"{esc(S['Recommended_Action_Plan'])}<br><span style='color:#CBD5E1'>Route to: {esc(S['Recommended_Faculty_Role'])}</span>"),
                unsafe_allow_html=True)

    section("Cohort stress map", "Each dot is a student. Top-right = big attendance drop AND many backlogs. The ⭐ is the selected student.")
    jit = np.random.default_rng(3)
    pdf = df.assign(Backlogs_jitter=df["Backlogs"] + jit.uniform(-0.22, 0.22, len(df)))
    fig = px.scatter(pdf, x="Attendance_Drop_Pct", y="Backlogs_jitter", color=lab_col, color_discrete_map=BURN_COLORS,
                     category_orders={lab_col: ["Low", "Moderate", "High"]}, size=idx_col, size_max=14, opacity=0.65,
                     hover_name="Student_Name", hover_data={"Backlogs_jitter": False, "Backlogs": True, "Department": True, idx_col: ":.1f"},
                     labels={"Attendance_Drop_Pct": "Attendance drop (pts)", "Backlogs_jitter": "Backlogs", lab_col: "Burnout risk"})
    fig.add_trace(go.Scatter(x=[S["Attendance_Drop_Pct"]], y=[S["Backlogs"]], mode="markers", name="Selected student",
                             marker=dict(symbol="star", size=22, color="#FFFFFF", line=dict(color="#D946EF", width=3))))
    show_fig(style_fig(fig, 420))

    section("🔥 Burnout watchlist", "Highest stress index first — filter by department for a faculty huddle.")
    wl_dept = st.selectbox("Department", ["All"] + sorted(df["Department"].unique()), key="wl_dept")
    wl = df if wl_dept == "All" else df[df["Department"] == wl_dept]
    wl = wl.sort_values(idx_col, ascending=False).head(15)
    show_df(wl[["Student_ID", "Student_Name", "Department", idx_col, lab_col, "Attendance_Drop_Pct", "Backlogs", "LMS_Weekly_Logins"]],
            hide_index=True, height=420,
            column_config={idx_col: st.column_config.ProgressColumn("Burnout index", min_value=0, max_value=100, format="%.1f")})
    guide("NEXT ➜", "Use <b>AI Action Copilot</b> to send a supportive message and assign a counselor.")

# ═════════════════════════════════════════════════════════════════════════════
# TAB 5 - AI ACTION COPILOT & INTERVENTION TRACKER
# ═════════════════════════════════════════════════════════════════════════════
with tab5:
    section("AI Action Copilot & Intervention Tracker", "Reach out, match a mentor, simulate improvement, and track every action.")
    active_strip()
    sub1, sub2, sub3, sub4 = st.tabs(["✉️ Outreach & Parent Nudge", "🤝 Peer Buddy", "🎛 What-If Simulator", "📋 Intervention Tracker"])

    # ── Outreach ────────────────────────────────────────────────────────────
    with sub1:
        guide("HOW", "Pick Email or SMS → review / edit the AI draft → click the bright button. "
                     "Then press <b>Mark as sent &amp; log</b> so it appears in the tracker.")
        oc1, oc2 = st.columns(2)

        def message_panel(col, kind, title):
            with col:
                st.markdown(f'<div class="sec-title" style="font-size:1.1rem;">{title}</div>', unsafe_allow_html=True)
                if kind == "parent":
                    need = S["Parent_Nudge_Draft"]
                    if "no parent nudge" in need.lower():
                        st.markdown('<span class="badge" style="background:#22C55E;color:#0B1220;">Not required</span> optional update only', unsafe_allow_html=True)
                    elif "optional" in need.lower():
                        st.markdown('<span class="badge" style="background:#FACC15;color:#0B1220;">Optional</span> after advisor review', unsafe_allow_html=True)
                    else:
                        st.markdown('<span class="badge" style="background:#EF4444;color:#FFFFFF;">Recommended</span> with advisor involvement', unsafe_allow_html=True)
                channel = st.radio("Channel", ["Email", "SMS"], horizontal=True, key=f"ch_{kind}_{active_id}")
                subject, body = build_message(S, kind, channel)
                if channel == "Email":
                    subject = st.text_input("Subject", subject, key=f"sub_{kind}_{active_id}")
                body = st.text_area("Message (editable)", body, height=250 if channel == "Email" else 120, key=f"body_{kind}_{active_id}_{channel}")
                if channel == "Email":
                    url = "mailto:?subject=" + urllib.parse.quote(subject) + "&body=" + urllib.parse.quote(body)
                    st.link_button("📧 Open in Email app", url)
                else:
                    st.link_button("💬 Open in SMS app", "sms:?&body=" + urllib.parse.quote(body))
                if st.button("✅ Mark as sent & log", key=f"log_{kind}_{active_id}"):
                    cur_status = tracker.at[active_id, "Intervention_Status"]
                    log_action(active_id, f"{title} sent via {channel}", status="In Progress" if cur_status == "Pending" else None)
                    st.success("Logged! Status updated in the Intervention Tracker.")
                src = S["Outreach_Message_Draft"] if kind == "student" else S["Parent_Nudge_Draft"]
                st.caption(f"AI source draft from dataset: {src}")

        message_panel(oc1, "student", "Student outreach")
        message_panel(oc2, "parent", "Parent nudge")

    # ── Peer buddy ──────────────────────────────────────────────────────────
    with sub2:
        guide("HOW", "The recommended mentor comes from the dataset's peer-matching model (same department). Alternatives are shown for flexibility.")
        buddy_id = S["Peer_Buddy_ID"]
        pb1, pb2 = st.columns([1.2, 1])
        with pb1:
            if buddy_id in by_id.index:
                B = by_id.loc[buddy_id]
                same = B["Department"] == S["Department"]
                st.markdown(card("⭐ Recommended peer mentor",
                                 f"<span style='font-size:1.4rem;font-weight:800;color:#FFFFFF'>{esc(B['Student_Name'])}</span><br>"
                                 f"{esc(B['Student_ID'])} · {esc(B['Department'])} · Semester {B['Semester']}<br><br>"
                                 f"Risk {badge(B['Overall_Risk'])} Success score <b>{B['Success_Score']:.1f}</b> "
                                 f"({B['Success_Score'] - S['Success_Score']:+.1f} vs {esc(S['Student_Name'].split()[0])})<br>"
                                 f"Attendance <b>{B['Attendance_Pct']:.1f}%</b> · Backlogs <b>{B['Backlogs']}</b><br><br>"
                                 f"{'✅ Same department' if same else '⚠ Different department'}"), unsafe_allow_html=True)
            else:
                st.markdown(card("⭐ Recommended peer mentor", esc(S["Peer_Buddy_Name"])), unsafe_allow_html=True)
        with pb2:
            fig = gauge(S["Peer_Match_Score"], "Peer match score", [(0, 40, "rgba(239,68,68,.55)"), (40, 70, "rgba(250,204,21,.55)"), (70, 100, "rgba(34,197,94,.55)")])
            show_fig(fig)
        if st.button("🤝 Assign this buddy & log", key=f"buddy_{active_id}"):
            log_action(active_id, f"Peer buddy assigned: {S['Peer_Buddy_Name']} ({buddy_id})")
            st.success("Peer buddy assignment logged in the tracker.")

        section("Other strong same-department mentors", "Higher success score, low risk, nearest semester.")
        pool = df[(df["Department"] == S["Department"]) & (df["Student_ID"] != active_id) & (df["Student_ID"] != buddy_id)
                  & (df["Overall_Risk"].isin(["Safe", "Moderate"])) & (df["Success_Score"] > S["Success_Score"])].copy()
        pool["Semester_gap"] = (pool["Semester"] - S["Semester"]).abs()
        pool = pool.sort_values(["Semester_gap", "Success_Score"], ascending=[True, False]).head(3)
        if len(pool):
            show_df(pool[["Student_ID", "Student_Name", "Semester", "Success_Score", "Overall_Risk", "Attendance_Pct", "Backlogs"]], hide_index=True)
        else:
            st.info("No stronger low-risk peers found in this department.")

    # ── What-if ─────────────────────────────────────────────────────────────
    with sub3:
        guide("HOW", "Move the sliders to see exactly how much the Success Score would improve.")
        coef, r2 = fit_success_model(df)
        c_att, c_bk, c_gap = coef[1], coef[2], coef[3]
        w1, w2 = st.columns([1, 1.2])
        with w1:
            cur_att = float(S["Attendance_Pct"])
            if cur_att < 100:
                new_att = st.slider("Raise attendance to (%)", cur_att, 100.0, float(min(cur_att + 10, 100)), 0.5, key=f"wi_att_{active_id}")
            else:
                new_att = cur_att
                st.caption("Attendance already at 100%.")
            if S["Backlogs"] > 0:
                cleared = st.slider("Backlogs cleared", 0, int(S["Backlogs"]), min(1, int(S["Backlogs"])), key=f"wi_bk_{active_id}")
            else:
                cleared = 0
                st.caption("No backlogs to clear.")
            max_gap = int(S["Skill_Gap_Score"])
            red = st.slider("Skill-gap reduction (points) — better coding / aptitude / interview prep", 0, max(max_gap, 1), min(5, max(max_gap, 1)),
                            key=f"wi_gap_{active_id}") if max_gap > 0 else 0
        d_att = c_att * (new_att - cur_att)
        d_bk = -c_bk * cleared
        d_gap = -c_gap * red
        base = float(S["Success_Score"])
        projected = float(np.clip(base + d_att + d_bk + d_gap, 0, 100))
        with w2:
            k1, k2, k3 = st.columns(3)
            k1.markdown(kpi("Current score", f"{base:.1f}", "today", "#64748B"), unsafe_allow_html=True)
            k2.markdown(kpi("Projected score", f"{projected:.1f}", "after plan", "#22C55E"), unsafe_allow_html=True)
            k3.markdown(kpi("Improvement", f"{projected - base:+.1f}", "points", "#D946EF"), unsafe_allow_html=True)
            fig = go.Figure(go.Waterfall(
                x=["Current", "Attendance", "Backlogs cleared", "Skill gap", "Projected"],
                measure=["absolute", "relative", "relative", "relative", "total"],
                y=[base, d_att, d_bk, d_gap, 0], connector={"line": {"color": "#475569"}},
                increasing={"marker": {"color": "#22C55E"}}, decreasing={"marker": {"color": "#EF4444"}},
                totals={"marker": {"color": "#38BDF8"}},
                text=[f"{base:.1f}", f"{d_att:+.1f}", f"{d_bk:+.1f}", f"{d_gap:+.1f}", f"{projected:.1f}"], textposition="outside",
                textfont=dict(color="#F8FAFC")))
            fig.update_layout(title="Where the improvement comes from", showlegend=False)
            fig.update_yaxes(range=[max(0, min(base, projected) - 12), min(100, max(base, projected) + 10)])
            show_fig(style_fig(fig, 320))
        st.caption(f"Model: Success Score fitted on all {len(df):,} students (R² = {r2:.2f}) → "
                   f"+1 attendance pt = {c_att:+.2f}, +1 backlog = {c_bk:+.2f}, +1 skill-gap pt = {c_gap:+.2f}. "
                   "The workbook has no CGPA/exam column, so 'grades' are represented by backlogs and skill gap.")
        pre1, pre2 = st.columns(2)
        pre1.markdown(card("Dataset scenario: attendance → 80%", f"Score becomes <b>{S['What_If_Attendance_80_Score']:.1f}</b> "
                                                                  f"({S['What_If_Attendance_80_Score'] - base:+.1f})"), unsafe_allow_html=True)
        pre2.markdown(card("Dataset scenario: +10 points", f"Score becomes <b>{S['What_If_Score_Plus_10']:.1f}</b> · "
                                                            f"est. recovery {S['Estimated_Recovery_Time_Weeks']} weeks"), unsafe_allow_html=True)
        if st.button("💾 Save this plan to tracker notes", key=f"wi_save_{active_id}"):
            log_action(active_id, f"What-if plan: attendance {cur_att:.0f}%→{new_att:.0f}%, clear {cleared} backlog(s), "
                                  f"cut skill gap {red} pts → score {base:.1f}→{projected:.1f}")
            st.success("Plan saved to the Intervention Tracker notes.")

    # ── Tracker ─────────────────────────────────────────────────────────────
    with sub4:
        guide("HOW", "Update the selected student below (Step 1), or edit any row in the table (Step 2). Changes save automatically for this session — download to keep them.")
        t = st.session_state.tracker
        sc = t["Intervention_Status"].value_counts()
        s1, s2, s3 = st.columns(3)
        s1.markdown(kpi("Pending", f"{int(sc.get('Pending', 0))}", "not started", "#F87171"), unsafe_allow_html=True)
        s2.markdown(kpi("In Progress", f"{int(sc.get('In Progress', 0))}", "being worked", "#38BDF8"), unsafe_allow_html=True)
        s3.markdown(kpi("Completed", f"{int(sc.get('Completed', 0))}", "closed", "#22C55E"), unsafe_allow_html=True)
        st.write("")

        st.markdown(_h('<div class="sec-title" style="font-size:1.1rem;"><span class="step">STEP 1</span>Update the selected student</div>'), unsafe_allow_html=True)
        roles = sorted(set(df["Recommended_Faculty_Role"]) | set(t["Assigned_Faculty_Role"]) | {"Faculty Advisor", "Student Counselor", "Placement Cell"})
        u1, u2 = st.columns([1.2, 1])
        with u1:
            new_status = st.select_slider("Status  (Pending → In Progress → Completed)", STATUS_ORDER,
                                          value=t.at[active_id, "Intervention_Status"] if t.at[active_id, "Intervention_Status"] in STATUS_ORDER else "Pending",
                                          key=f"st_{active_id}_{st.session_state.tracker_ver}")
            note = st.text_area("Add a note", placeholder="e.g. Met student, agreed on weekly attendance check-ins", key=f"note_{active_id}_{st.session_state.tracker_ver}")
        with u2:
            cur_role = t.at[active_id, "Assigned_Faculty_Role"]
            new_role = st.selectbox("Assigned faculty role", roles, index=roles.index(cur_role) if cur_role in roles else 0,
                                    key=f"role_{active_id}_{st.session_state.tracker_ver}")
            st.markdown(card("Existing notes", esc(t.at[active_id, "Notes"]) or "<span style='color:#CBD5E1'>No notes yet</span>"), unsafe_allow_html=True)
        if st.button("💾 Save update for this student", key=f"save_{active_id}"):
            changes = []
            if new_status != t.at[active_id, "Intervention_Status"]:
                changes.append(f"status → {new_status}")
            if new_role != t.at[active_id, "Assigned_Faculty_Role"]:
                changes.append(f"assigned to {new_role}")
            text = "; ".join(changes + ([note.strip()] if note.strip() else []))
            log_action(active_id, text or "Reviewed", status=new_status, role=new_role)
            st.success("Saved.")
            st.rerun()

        st.markdown("---")
        st.markdown(_h('<div class="sec-title" style="font-size:1.1rem;"><span class="step">STEP 2</span>Manage all interventions</div>'), unsafe_allow_html=True)
        h1, h2, h3, h4 = st.columns(4)
        tf_pri = h1.multiselect("Priority", sorted(t["Intervention_Priority"].unique()), default=[p for p in sorted(t["Intervention_Priority"].unique()) if p.startswith(("P1", "P2"))])
        tf_status = h2.multiselect("Status", STATUS_ORDER, key="tf_status")
        tf_dept = h3.multiselect("Department", sorted(t["Department"].unique()), key="tf_dept")
        tf_search = h4.text_input("Search name / ID", key="tf_search")
        tv = t.copy()
        if tf_pri:
            tv = tv[tv["Intervention_Priority"].isin(tf_pri)]
        if tf_status:
            tv = tv[tv["Intervention_Status"].isin(tf_status)]
        if tf_dept:
            tv = tv[tv["Department"].isin(tf_dept)]
        if tf_search.strip():
            tv = tv[(tv["Student_ID"] + " " + tv["Student_Name"]).str.lower().str.contains(re.escape(tf_search.strip().lower()))]
        tv = tv.sort_values(["Intervention_Priority", "Student_ID"]).head(300)
        st.caption(f"Showing {len(tv)} rows (max 300). Double-click Status, Faculty role or Notes to edit.")
        edit_sig = abs(hash((str(tf_pri), str(tf_status), str(tf_dept), tf_search, st.session_state.tracker_ver)))
        edited = st.data_editor(
            tv[["Student_ID", "Student_Name", "Department", "Overall_Risk", "Intervention_Priority", "Assigned_Faculty_Role",
                "Intervention_Status", "Notes", "Last_Updated"]],
            key=f"editor_{edit_sig}", hide_index=True, height=420,
            disabled=["Student_ID", "Student_Name", "Department", "Overall_Risk", "Intervention_Priority", "Last_Updated"],
            column_config={
                "Intervention_Status": st.column_config.SelectboxColumn("Status", options=STATUS_ORDER, required=True),
                "Assigned_Faculty_Role": st.column_config.SelectboxColumn("Faculty role", options=roles, required=True),
                "Notes": st.column_config.TextColumn("Notes", width="large"),
            },
        )
        # auto-apply edits to the master tracker
        try:
            for col in ["Intervention_Status", "Assigned_Faculty_Role", "Notes"]:
                st.session_state.tracker.loc[edited.index, col] = edited[col].values
        except Exception:
            pass
        st.download_button("⬇ Download tracker (CSV)", st.session_state.tracker.to_csv(index=False).encode("utf-8"),
                           f"EduNexus_intervention_tracker_{stamp}.csv", "text/csv", key="dl_tracker")

st.markdown("---")
st.caption("EduNexus AI · Bytexel Hackathon · Data: Student360 feature workbook · Built with Streamlit & Plotly")