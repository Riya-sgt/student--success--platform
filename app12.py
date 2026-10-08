"""
EduNexus AI - Smart Campus Analytics & Student Success Platform
Bytexel Hackathon final submission.

Run:   streamlit run app.py
Data:  Student360_1000_Enhanced_ML_Dataset-1.xlsx (1000 students x 45 columns)
       - the full 1000-student dataset is BUILT IN, so the app always shows all 1000 students
       - upload another .xlsx from the sidebar to override it
       - or place the .xlsx next to app.py and it will be picked up automatically.
"""
import base64
import html
import io
import os
import re
import zlib
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG & CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="EduNexus AI", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

LOCAL_FILES = [
    "Student360_1000_Enhanced_ML_Dataset-1.xlsx",
    "Student360_1000_Enhanced_ML_Dataset.xlsx",
    "Student360_1000_Feature_Only_Single_Sheet_With_Gender.xlsx",
]
ROLES = ["Administrator", "Faculty / Mentor", "Placement Officer", "Student (View Own Data)"]
RISK_ORDER = ["Critical", "High", "Moderate", "Safe"]
RISK_COLORS = {"Critical": "#EF4444", "High": "#F97316", "Moderate": "#FACC15", "Safe": "#22C55E"}
BURN_COLORS = {"Critical": "#EF4444", "High": "#EF4444", "Moderate": "#FACC15", "Low": "#22C55E"}
STATUS_ORDER = ["Pending", "In Progress", "Completed"]
STATUS_COLORS = {"Pending": "#F87171", "In Progress": "#38BDF8", "Completed": "#22C55E"}
CTA_A, CTA_B = "#7C3AED", "#D946EF"
RADAR_AXES = [
    ("Aptitude", "Aptitude_Score"), ("Coding", "Coding_Score"), ("Interview", "Mock_Interview_Score"),
    ("Technical", "Technical_Skill_Score"), ("Soft Skills", "Soft_Skill_Score"),
]
# static fallback baseline (full 1000-student cohort means) used when the loaded data is tiny
RADAR_BASELINE_FALLBACK = {"Aptitude_Score": 68.5, "Coding_Score": 69.4, "Mock_Interview_Score": 67.4,
                           "Technical_Skill_Score": 62.6, "Soft_Skill_Score": 63.7}

# ─────────────────────────────────────────────────────────────────────────────
# CSS  (hidden header, dark gradient hero, KPI boxes, cards, vivid CTA buttons)
# ─────────────────────────────────────────────────────────────────────────────
CSS = f"""
<style>
[data-testid="stHeader"] {{display: none !important;}}
[data-testid="stToolbar"], #MainMenu, footer {{display: none !important;}}
[data-testid="stSidebarCollapseButton"] {{display: none !important;}}
.stApp {{background:#0E1117;}}
[data-testid="stSidebar"] {{background:#111827; border-right:1px solid #1F2A44;}}
.block-container {{padding-top:1.2rem; max-width:1500px;}}
.stApp p, .stApp li, .stApp label, [data-testid="stMarkdownContainer"] {{color:#E5E7EB;}}
.stApp h1, .stApp h2, .stApp h3, .stApp h4 {{color:#FFFFFF !important;}}
[data-testid="stCaptionContainer"] {{color:#CBD5E1 !important;}}
.hero {{background:linear-gradient(120deg,#0F1B3D 0%,#1E1250 55%,#4A0F5E 100%); border:1px solid #2B3A63;
  border-radius:18px; padding:22px 28px; margin-bottom:14px;}}
.hero-title {{font-size:2.5rem; font-weight:800; color:#FFFFFF; line-height:1.1;}}
.hero-title span {{background:linear-gradient(90deg,#38BDF8,#D946EF); -webkit-background-clip:text; -webkit-text-fill-color:transparent;}}
.hero-tag {{color:#CBD5E1; font-size:1.05rem; margin-top:6px;}}
.hero-chip {{display:inline-block; margin:10px 8px 0 0; padding:3px 12px; border-radius:999px; background:#0B1220; border:1px solid #334155; color:#E2E8F0; font-size:.78rem;}}
.card {{background:#162033; border:1px solid #26344D; border-radius:14px; padding:16px 18px; color:#F1F5F9; margin-bottom:10px;}}
.card-title {{font-size:.8rem; font-weight:700; letter-spacing:.9px; text-transform:uppercase; color:#93C5FD; margin-bottom:6px;}}
.card-text {{font-size:1rem; color:#F8FAFC; line-height:1.55;}}
.kpi {{background:#162033; border:1px solid #26344D; border-left:6px solid var(--accent,#38BDF8); border-radius:14px; padding:14px 18px; margin-bottom:6px;}}
.kpi-label {{font-size:.76rem; font-weight:700; letter-spacing:.8px; text-transform:uppercase; color:#CBD5E1;}}
.kpi-value {{font-size:2rem; font-weight:800; color:#FFFFFF; line-height:1.15;}}
.kpi-sub {{font-size:.8rem; color:#CBD5E1;}}
.sec-title {{font-size:1.35rem; font-weight:800; color:#FFFFFF; margin:6px 0 2px 0;}}
.sec-sub {{color:#CBD5E1; font-size:.95rem; margin-bottom:10px;}}
.active-strip {{background:#14213A; border:1px solid #2F4A7D; border-radius:12px; padding:10px 16px; margin-bottom:12px; color:#F8FAFC;}}
.badge {{display:inline-block; padding:3px 12px; border-radius:999px; font-weight:800; font-size:.82rem; margin:2px 6px 2px 0;}}
.chip {{display:inline-block; padding:4px 12px; border-radius:8px; font-weight:700; font-size:.88rem; margin:3px 6px 3px 0;}}
.chip-ok {{background:#14532D; color:#DCFCE7; border:1px solid #22C55E;}}
.chip-bad {{background:#7F1D1D; color:#FEE2E2; border:1px solid #EF4444;}}
.chip-warn {{background:#713F12; color:#FEF3C7; border:1px solid #FACC15;}}
.chip-neutral {{background:#1E3A5F; color:#E0F2FE; border:1px solid #38BDF8;}}
.inbox-new {{background:linear-gradient(90deg,#4A0F5E,#7C1D6F); border:2px solid #F0ABFC; border-radius:14px; padding:16px 20px; margin-bottom:14px; color:#FFFFFF;}}
.inbox-empty {{background:#162033; border:1px dashed #475569; border-radius:14px; padding:12px 20px; margin-bottom:14px; color:#E2E8F0;}}
.mail {{background:#0B1220; border:1px solid #475569; border-radius:10px; padding:12px 16px; color:#F8FAFC; white-space:pre-wrap; font-size:.95rem; line-height:1.5; margin-top:8px;}}
.seg {{background:#162033; border:1px solid #26344D; border-top:5px solid var(--accent,#38BDF8); border-radius:14px; padding:16px 18px; margin-bottom:8px; min-height:170px;}}
.seg-count {{font-size:2.4rem; font-weight:800; color:#FFFFFF; line-height:1.1;}}
.seg-title {{font-size:1.05rem; font-weight:800; color:#FFFFFF;}}
.seg-desc {{color:#CBD5E1; font-size:.88rem; margin-top:4px;}}
button[data-testid^="stBaseButton-secondary"], button[data-testid^="stBaseButton-primary"],
[data-testid="stFormSubmitButton"] button, .stDownloadButton button {{
  background:linear-gradient(90deg,{CTA_A},{CTA_B}) !important; color:#FFFFFF !important; border:0 !important;
  font-weight:800 !important; border-radius:10px !important; padding:.5rem 1.1rem !important; box-shadow:0 4px 18px rgba(217,70,239,.35);}}
button[data-testid^="stBaseButton-secondary"] p, button[data-testid^="stBaseButton-primary"] p, .stDownloadButton button p {{color:#FFFFFF !important; font-weight:800 !important;}}
button[data-testid^="stBaseButton-secondary"]:hover, button[data-testid^="stBaseButton-primary"]:hover {{filter:brightness(1.15);}}
button[data-baseweb="tab"] {{font-weight:700 !important; font-size:1rem !important;}}
button[data-baseweb="tab"] p {{color:#CBD5E1 !important;}}
button[data-baseweb="tab"][aria-selected="true"] p {{color:#FFFFFF !important;}}
div[data-baseweb="tab-highlight"] {{background:{CTA_B} !important; height:3px !important;}}
.stApp input, .stApp textarea {{color:#F8FAFC !important; -webkit-text-fill-color:#F8FAFC !important;}}
div[data-baseweb="input"], div[data-baseweb="textarea"], div[data-baseweb="select"] > div {{background:#1B2538 !important; border-color:#334155 !important;}}
div[data-baseweb="select"] * {{color:#F8FAFC !important; -webkit-text-fill-color:#F8FAFC !important;}}
div[data-baseweb="popover"] ul {{background:#1B2538 !important;}}
div[data-baseweb="popover"] li {{color:#F8FAFC !important;}}
[data-testid="stExpander"] {{border:1px solid #26344D; border-radius:12px; background:#111827;}}
hr {{border-color:#26344D !important;}}
</style>
"""


# ─────────────────────────────────────────────────────────────────────────────
# SMALL HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _h(markup: str) -> str:
    """Remove indentation/newlines so Markdown never turns our HTML into a code block."""
    return "".join(line.strip() for line in markup.splitlines())


def esc(x) -> str:
    return html.escape(str(x))


_chart_n = {"n": 0}


def show_fig(fig):
    _chart_n["n"] += 1
    ss_ = st.session_state
    key = f"chart_{ss_.get('role', 'r')}_{ss_.get('active_id', 'x')}_{ss_.get('fac_dept', '')}_{_chart_n['n']}"
    try:
        st.plotly_chart(fig, width="stretch", key=key)
    except Exception:
        st.plotly_chart(fig, use_container_width=True, key=key)


def show_df(data, **kw):
    try:
        return st.dataframe(data, width="stretch", **kw)
    except Exception:
        return st.dataframe(data, use_container_width=True, **kw)


def style_fig(fig, height=340):
    fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(color="#E5E7EB", size=13), margin=dict(l=10, r=10, t=50, b=10), height=height,
                      title_font=dict(color="#FFFFFF", size=16), legend=dict(font=dict(color="#E5E7EB")))
    fig.update_xaxes(gridcolor="#1F2A44", zerolinecolor="#1F2A44")
    fig.update_yaxes(gridcolor="#1F2A44", zerolinecolor="#1F2A44")
    return fig


def badge(label, kind="risk") -> str:
    palette = BURN_COLORS if kind == "burn" else RISK_COLORS
    bg = palette.get(str(label), "#64748B")
    fg = "#FFFFFF" if bg in ("#EF4444", "#64748B") else "#0B1220"
    return f'<span class="badge" style="background:{bg};color:{fg};">{esc(label)}</span>'


def kpi(label, value, sub="", accent="#38BDF8") -> str:
    return _h(f"""<div class="kpi" style="--accent:{accent}"><div class="kpi-label">{esc(label)}</div>
        <div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>""")


def card(title, body_html) -> str:
    return _h(f'<div class="card"><div class="card-title">{esc(title)}</div><div class="card-text">{body_html}</div></div>')


def section(title, sub=""):
    st.markdown(_h(f'<div class="sec-title">{title}</div><div class="sec-sub">{sub}</div>'), unsafe_allow_html=True)


def gauge(value, title, steps, rng=(0, 100), ref=None, suffix="", num_size=44):
    cfg = {"axis": {"range": list(rng), "tickcolor": "#94A3B8", "tickfont": {"color": "#CBD5E1"}},
           "bar": {"color": "#38BDF8", "thickness": 0.28}, "bgcolor": "#111827", "borderwidth": 0,
           "steps": [{"range": [a, b], "color": c} for a, b, c in steps]}
    if ref is not None:
        cfg["threshold"] = {"line": {"color": "#FFFFFF", "width": 3}, "thickness": 0.85, "value": float(ref)}
    fig = go.Figure(go.Indicator(mode="gauge+number", value=float(value),
                                 number={"suffix": suffix, "font": {"color": "#FFFFFF", "size": num_size}},
                                 title={"text": title, "font": {"color": "#E2E8F0", "size": 15}}, gauge=cfg))
    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#E5E7EB"), height=265, margin=dict(l=25, r=25, t=60, b=10))
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# DATA LAYER
# ─────────────────────────────────────────────────────────────────────────────
SCHEMA = {  # exact 45-column schema -> default used if a column is missing
    "Student_ID": "", "Student_Name": "", "Department": "Unknown", "Gender": "Unknown", "Semester": 1, "CGPA": 7.0,
    "Success_Score": 50.0, "Academic_Risk": "Moderate", "Placement_Risk": "Moderate", "Aptitude_Score": 65.0,
    "Coding_Score": 65.0, "Mock_Interview_Score": 65.0, "Overall_Risk": "Moderate", "Behavioral_Persona": "Balanced",
    "AI_Risk_Explanation": "No explanation available.", "Feedback_Score": 65.0, "Current_Skills": "", "Target_Role": "Not set",
    "Required_Skills": "", "Missing_Skills": "No major gap", "Skill_Gap_Level": "Moderate", "Skill_Gap_Score": 30.0,
    "Technical_Skill_Score": 60.0, "Soft_Skill_Score": 60.0, "Previous_Attendance_Pct": 75.0, "Attendance_Pct": 75.0,
    "Attendance_Drop_Pct": 0.0, "Backlogs": 0, "LMS_Weekly_Logins": 5.0, "Academic_Burnout_Stress_Risk_Index": 20.0,
    "Academic_Burnout_Stress_Risk": "Low", "Outreach_Message_Draft": "Progress message: continue current academic and placement preparation.",
    "Parent_Nudge_Draft": "No parent nudge required.", "Peer_Buddy_ID": "", "Peer_Buddy_Name": "Not assigned",
    "Peer_Match_Score": 0.0, "What_If_Attendance_80_Score": 50.0, "What_If_Score_Plus_10": 60.0,
    "What_If_Improvement_Potential": 10, "Estimated_Recovery_Time_Weeks": 4,
    "Recommended_Action_Plan": "Continue regular mentoring", "Recommended_Faculty_Role": "Faculty Advisor",
    "Intervention_Priority": "P3 - Moderate", "Intervention_Status": "Pending", "Report_Eligible": "No",
}
NUMERIC = ["Semester", "CGPA", "Success_Score", "Aptitude_Score", "Coding_Score", "Mock_Interview_Score", "Feedback_Score",
           "Skill_Gap_Score", "Technical_Skill_Score", "Soft_Skill_Score", "Previous_Attendance_Pct", "Attendance_Pct",
           "Attendance_Drop_Pct", "Backlogs", "LMS_Weekly_Logins", "Academic_Burnout_Stress_Risk_Index", "Peer_Match_Score",
           "What_If_Attendance_80_Score", "What_If_Score_Plus_10", "What_If_Improvement_Potential", "Estimated_Recovery_Time_Weeks"]


# ---- Built-in copy of Student360_1000_Enhanced_ML_Dataset-1.xlsx (1000 students x 45 columns, zlib+base64 CSV).
# ---- Do not edit this block. It guarantees the app always shows the full 1000-student dataset.
_EMBEDDED_DATASET_B64 = (
    "eNrcvVuX3EZ2Jvp+fkUtvk5ONeIeYT+ppbatsdSWmz3jNU9aZbKaLDdF8hSLknV+/cH+dgAZCAQycYlEFjlrFq0E2WQWsLFjX77Ly6fPr+/fP/"
    "38/XeHl/E//3z3y/3hu/uPd49Pv7SfD/98//71/ePh5f0v95+e2v/49p9/+ubw8vOrV/efPv388tWHx/vDN6/uXt//8vDq5788fPr74ad3d6/u"
    "6X/KH7/5+PRAf3X8s99+eP3w/k388OOHV3//+fv37V/768P9b/Hiv/16/3j37h3/r/94//bu14cP7YWff7p//PTh/d3hm+/xWz//6b8/vrt7f/"
    "f08OH94Z/u71//5137l8V/5PPjI32Bl39/ePfu0+Gvd49v7tuv8+Hd/eEv9//v54fH+9fd7/348OkTvhB/xP/5+Z/vPv78w/2v9++Sz/w3//X+"
    "1dv3D6/ar8O/w1dffvjb0+DCT4/37bf+/Onnb56e2vt39/7V/c8/vXo6TH/87vHDR1z7Y/tjvPvw5tPhhx9f/vwf9/d/f/f7zz98ePPw/tPxPv"
    "/x8+P7D5/bf/LpkZ4Cbsf37WP675N/5PBvn9sPd6/e/vxje+nuDf2jd397Ovx0h5v158+vj5fu7x/bv+L1698pNJJPiA58/vHuqf2b+Mf9j7d3"
    "bQz9Lf1pfZP9Hj79/NO79qaIpr/6/S8fHz/8yuHy04f2f/30cPfu8KdPTw+/3D21T+kv96/a3378/ee/Pvxyj9vx6UDXfvmFwvL1z9+8ouff/r"
    "V37wfX/+nu1ed3T7/zM+cIo7+b/uhjG04PT78Pr758unv6TH/1xw+PTz//6d3Dm4f/fHf//7z86/8WTSMO39w93v168/Lt3eMvd4dvX/7p8ONd"
    "+/fKg7ttDlbc6sPLu79ROLdvSvu1D07fhoMXt+0fsO2f6K//8OG3mz+9f9PeerxaL9rv8svd4+83f7t79fTh8dM/3Lxr/8BdfGFuPuH+4dp9/z"
    "+6vYnvx81j+0hvHj7ddH/77Qv610L7t/7+9LZ9J25e/vsPLxCav7UP+OY7iucPH9s3+MV3L7853PCf+sP/uvv17nDzzw9P8c+/6N7Xw8237d38"
    "TMFOd+jF8adQ/tYdrGl/MNPQz9nc+oM3t+rQtJfEQdKvqr0n7U/bvggf3lAA3vzCIfcPN68+tHf8/ef7m1f8kt7cxZC9uXv/+uZjlzxuPj5SDs"
    "K/fXv484ebj4jSm/cUpTeP8S2+PdATEi607/j73+9u/vnzx6e7G1xov5viZ+PolzbmDL/NN0+Pdw/v23f+5m8fHm8K9yfGzs03r399+PTh8fCT"
    "uvmf/V1ug/89ZbD2K8XokIf/8/D27u79zb9+bh/m4buXHBziYG+to2+gODr+5eHN24NR7Ucjbi3dQHW8p33SvMGbOo6M3+7v/t7eO/qnu8DApY"
    "cue7b35K793fY2nwwR69t/uw+RH9sn/k+f2z/Zxv+rv6cx8i9//fGHNgRevjzcUIi8fPX48LGNkb+0+aP9P39u/8Lb//rUhgvn8yxYDjd9Uk/i"
    "RtOtsBQehp5ICO038YgZ+q/mQGHV/hmT3JU9osfJ4ft9gyvacfgoCh9F4aNK4VO6ef94Q+da8mQ+PrY37eFV+pS/vX/37lxgqcM3r9tEdXfzx7"
    "v3n9qseIwsfysEh3WWdygftbfUUJzJ4/U/3r2jpPz6cD6sTgTPwYf2X+xjp80i7cP9gRJaG3YUM8eoQjJBWnkRT/wkfdD3c4q+pG8zR/uD0pd2"
    "7cMXkpLIQYj2v6Vpf4DdMkg5Bto76fg2+5hCZCkG0nuwMHfowzeP//X5/c3/uaeD5fu/8gNu/1EKPt/egez5hvau4Q1u860oPN9ZeWN8yJzMGF"
    "7TU8fz/OnDb/ePN3/8/sXhu7s2z37z/u7d75+e4m/+6b9ftQVT/2e6EybJEd3Rkp4lmn5SOi5twC1GMJj2xxS6/aVNFW2IyNA+h91iIfjsxbvB"
    "Je3b74dn4uiXqWAY3JiF0WAO/3r3nw+PN//rw6e3D22d275HfZ0hPd0mlQVEm001vUaSwtQvKDRmhcWy2kOnyQG1BO7Fy1cPVGNRmPS/2RZb7Y"
    "WHV59inqBj6FSUaPq72wfQnuq6DX6OkkA1hsXJgYyBU2S3KDE2HL75pT31/3zXPjF8VPQF8JCcPnFkZHdlYYzYw18+vG3/2W/ffmiPlWEtKtpY"
    "oFel/1/3/9HetoYLNUuJA6XI9KEwJ0HQX3FLoZdUE8gAy3JDX22mxwO+ZfsDNVQliPapOxwUqv3msi0vBVWgSbC//PyRqvaHX+9vPsQO5x9uXj"
    "98evW5jYHxz4Jn/qr9Y0/tf958etv+T28ejm0IhUP79P/tI0VB+/JzDPzhzee7x9cP7Y3/2EXX54+v8Vf8rT3mb+746d1Q33f/24mytP2hJD0l"
    "JDxHQaK7lu/m1bv7NvraZ3LziX+kf7xZnGFkGz14vN+/v+lehMP/vf8Uw8cd/vUz/VQv27/t7TF62hx8qwK9Z5aD45hh6Ku2j7k9rV37C36Tss"
    "ux1ZtqY45/4OaX9js+fHx3f/OfXXO7oHhFpL04GNneuEJ+mX8InS1QlaHGpn2PqHwylFId/bOCU0v7RrffwA7y7IvzsbfuNswK0hc1otQoffi+"
    "PW/aW/njw6e3j22c0iVDpYZBsHLKPfhkXDH6Lv94syGI4+ip7SE+v//U1tDnw9gf/u/dp7c3bey+/fCUpkFFR2Wgp2STrkvL9qOmY1I3ad1Us+"
    "tadlhSef7i2//xP2IZvaqcKvZYil5Xemc1TR5C2ykcgqXsSSdSmzwDVdk7npMiO7FucEXj7KaH5ARHlz5XTNXqqpJju++oLEWNoS8kS4cnpjzU"
    "tViZ9uwnqu78zS/VVqvad0eVTxI425r346ynmA4lpUNKAZbSoaXhUlveUBR5VOY04miu3aRZepWs5jmP5FA6mag2d/HLKrb2C/3l7u3ndzc/3r"
    "99uhtU9fZWi4MOadBxugo8W0MbHw/ks6lqW4jFQ7Z9A8ThBc8EV5TwE7PDifCi+ke69gdsy4uGZxeIMdGg0KNjByFIAfZiYZH3LE7WcYugqaQ1"
    "VEKYwMFqqp+qx6eVJ8bTh6po3673KFf/9e7jhw/Ijv90/wuFqo4NKB0uWQNKcxtL3YU1aedxZuK0LO1R/S+y2u/bdx8+v05mTz+0yea/29x2//"
    "Tbh8e/cyP5Hy//8M3/95mC4bv2Vb5/fHHo43DQbzSUOgRNzw2FoeNOiafZVPQ+gyyH4hv33ym8IuXecnhXqqUw2X4jCoxv3tw9/nb3jgZVMTLa"
    "hvfWq4Ox3awK8WUxp6Lyua28zMwEtmnCHfOXo3xZ7Xg8WWj1zYExNI+j6VCglE1hQ8kr0ATLzkheC37u/XLXxBDM0GQOD9vaLn8tS1HbBuZnMp"
    "g6fPfQfuOf2p8tDVKajWhNU5kYpPglUCnjaajkZw7L33+4+eXuv+hOUdS9fmyf5+OptBVomdGdph8f21swa2D+5+6feXP3EfGD9tNR7e4N/UJ5"
    "Cgs3FDwHVNBU3++4cKNR5a/trW5z1NsbfA40r8Ut9rqrw77t/unH+zef39093tA/9IFuxNIMlPWoaPj6w6lNQShPm/Hh1HBVaJtFa9gqPZ6l1L"
    "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"
)


@st.cache_data(show_spinner=False)
def embedded_dataset() -> pd.DataFrame:
    return pd.read_csv(io.BytesIO(zlib.decompress(base64.b64decode(_EMBEDDED_DATASET_B64))))


def normalise(df: pd.DataFrame) -> pd.DataFrame:
    """Guarantee every one of the 45 columns exists with the right dtype so nothing can KeyError."""
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
    if df.empty:
        raise ValueError("no rows")
    return df


@st.cache_data(show_spinner=False)
def read_excel_cached(src):
    handle = io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src
    return pd.read_excel(handle, sheet_name=0)  # first sheet works for 'Sheet1' and 'student_features'


def find_local_file():
    here = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    for folder in [".", here, os.path.join(here, "data"), os.getcwd()]:
        for name in LOCAL_FILES:
            p = os.path.join(folder, name)
            if os.path.exists(p):
                return p
    return None


def load_data(uploaded):
    """Uploaded file -> local file next to app.py -> built-in 1000-student dataset. Never raises."""
    note = ""
    if uploaded is not None:
        try:
            return normalise(read_excel_cached(uploaded.getvalue())), "upload", note
        except Exception:
            note = "The uploaded file could not be read, so the default dataset is shown."
    try:
        path = find_local_file()
        if path:
            return normalise(read_excel_cached(path)), "local", note
    except Exception:
        pass
    return normalise(embedded_dataset()), "builtin", note


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE  (no external database)
# ─────────────────────────────────────────────────────────────────────────────
def init_state(df: pd.DataFrame, source: str):
    sig = f"{source}-{len(df)}-{df['Student_ID'].iloc[0]}-{df['Student_ID'].iloc[-1]}"
    if st.session_state.get("data_sig") != sig:
        st.session_state.data_sig = sig
        st.session_state.status = df.set_index("Student_ID")["Intervention_Status"].to_dict()   # Intervention_Status
        st.session_state.drafts = {}      # Outreach_Message_Draft (teacher-edited text per student)
        st.session_state.inbox = {}       # messages delivered to each student's inbox
        st.session_state.activity = {}    # last workflow action per student
        crit = df[df["Overall_Risk"] == "Critical"]
        st.session_state.active_id = (crit if len(crit) else df).iloc[0]["Student_ID"]
        st.session_state.my_id = st.session_state.active_id


def default_email(s, sender_role) -> str:
    draft = str(s["Outreach_Message_Draft"]).strip()
    m1 = re.match(r"Supportive outreach:\s*(.*)", draft, flags=re.I)
    m2 = re.match(r"Progress message:\s*(.*)", draft, flags=re.I)
    if m1:
        core = f"Your faculty team would like to {m1.group(1)}"
    elif m2:
        core = f"Great progress so far - {m2.group(1)}"
    else:
        core = draft or "Please continue your current academic and placement preparation."
    return (f"Dear {s['Student_Name']},\n\n{core}\n\n"
            f"Recommended next step: {s['Recommended_Action_Plan']}.\n"
            f"Your peer buddy: {s['Peer_Buddy_Name']} (match score {s['Peer_Match_Score']:.0f}%).\n"
            f"Expected recovery window: about {s['Estimated_Recovery_Time_Weeks']} weeks.\n\n"
            f"Regards,\n{s['Recommended_Faculty_Role']}, {s['Department']} Department\nEduNexus AI Student Success Cell")


# Callbacks run BEFORE the script reruns, so every tab (incl. the Student 360 inbox) sees the new state instantly.
def cb_send_email(aid, sender_role, key):
    ss = st.session_state
    body = ss.get(key, "")
    now = datetime.now().strftime("%d %b %Y, %H:%M")
    ss.drafts[aid] = body
    ss.inbox.setdefault(aid, []).append({
        "time": now, "from": sender_role, "subject": "Your personalised support plan - EduNexus AI",
        "body": body, "read": False})
    ss.status[aid] = "In Progress"
    ss.activity[aid] = f"{now} · Email sent by {sender_role}"
    ss.flash = ("ok", "Email delivered to the student's inbox and status set to In Progress. Open the Student 360 tab to see it.")
    st.toast("📧 Email sent - status: In Progress")


def cb_complete(aid, sender_role):
    ss = st.session_state
    now = datetime.now().strftime("%d %b %Y, %H:%M")
    ss.status[aid] = "Completed"
    ss.activity[aid] = f"{now} · Marked Completed by {sender_role}"
    ss.flash = ("ok", "Intervention marked as Completed.")
    st.toast("✅ Intervention completed")


def cb_regen(key, aid, sender_role):
    s = st.session_state.get("_cur_student")
    if s is not None:
        st.session_state[key] = default_email(s, sender_role)
        st.session_state.drafts.pop(aid, None)


def cb_mark_read(aid):
    for m in st.session_state.inbox.get(aid, []):
        m["read"] = True


# ─────────────────────────────────────────────────────────────────────────────
# SKILL HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def split_list(s):
    return [x.strip() for x in str(s).split(",") if x.strip()]


def skill_match(required, current):
    cur = {c.lower() for c in current}
    return {r: any(alt.strip().lower() in cur for alt in r.split("/")) for r in required}


# ═════════════════════════════════════════════════════════════════════════════
# APP START: SIDEBAR  (uploader -> role -> student dropdown)
# ═════════════════════════════════════════════════════════════════════════════
st.markdown(CSS, unsafe_allow_html=True)

with st.sidebar:
    st.markdown(_h('<div class="hero-title" style="font-size:1.6rem;">Edu<span>Nexus</span> AI</div>'
                   '<div class="hero-tag" style="font-size:.85rem;">Smart Campus Analytics &amp; Student Success</div>'),
                unsafe_allow_html=True)
    st.markdown("---")
    uploaded = st.file_uploader("📂 Upload Student360 dataset (.xlsx)", type=["xlsx"])

df, source, load_note = load_data(uploaded)
init_state(df, source)
by_id = df.set_index("Student_ID", drop=False)
by_id.index.name = None
ss = st.session_state
name_of = by_id["Student_Name"].to_dict()

def demo_pin(sid) -> str:
    """Demo credential for the student sign-in: the last 4 digits of the Student ID (STU1002 -> 1002)."""
    return re.sub(r"\D", "", str(sid))[-4:]


def cb_sign_out():
    st.session_state.student_auth = None


with st.sidebar:
    if load_note:
        st.caption(load_note)
    st.markdown("### 🔐 Access level")
    role = st.radio("Signed in as", ROLES, key="role", label_visibility="collapsed")
    is_student = role.startswith("Student")
    is_staff = not is_student
    all_ids = df["Student_ID"].tolist()
    depts_all = sorted(df["Department"].unique())

    # ---- who is allowed to see which students -------------------------------------------------
    if not is_student:
        ss.student_auth = None                      # leaving student mode always signs the student out
    elif ss.get("student_auth") not in by_id.index:
        ss.student_auth = None
    signed_in = (not is_student) or (ss.student_auth is not None)

    options = all_ids
    if is_student:
        if ss.student_auth is None:
            st.markdown("**Student sign-in**")
            with st.form("student_login", clear_on_submit=False):
                sid_in = st.text_input("Student ID", placeholder="e.g. STU1002")
                pin_in = st.text_input("PIN", type="password", placeholder="4-digit PIN")
                go_login = st.form_submit_button("🔓 Sign in")
            if go_login:
                sid_try = sid_in.strip().upper()
                if sid_try in by_id.index and pin_in.strip() == demo_pin(sid_try):
                    ss.student_auth = sid_try
                    st.rerun()
                else:
                    st.error("Invalid Student ID or PIN.")
            with st.expander("ℹ️ Demo credentials (for judges)"):
                st.caption("PIN = the last 4 digits of the Student ID. Example: STU1002 → PIN 1002. "
                           "Once signed in, a student can only see their own data and cannot switch to another student.")
        else:
            st.markdown(_h(f'<div class="card" style="padding:10px 14px;"><div class="card-title">Signed in as</div>'
                           f'<div class="card-text"><b>{esc(by_id.at[ss.student_auth, "Student_Name"])}</b><br>'
                           f'<span style="color:#CBD5E1">{esc(ss.student_auth)}</span></div></div>'), unsafe_allow_html=True)
            st.button("🚪 Sign out", key="sign_out", on_click=cb_sign_out)
            options = [ss.student_auth]
            ss.active_id = ss.student_auth
    elif role == "Faculty / Mentor":
        if ss.get("fac_dept") not in depts_all:
            ss.fac_dept = by_id.at[ss.active_id, "Department"] if ss.get("active_id") in by_id.index else depts_all[0]
        st.selectbox("Your department (faculty sign-in)", depts_all, key="fac_dept")
        options = df[df["Department"] == ss.fac_dept]["Student_ID"].tolist()
    elif role == "Placement Officer":
        opts = df[df["Placement_Risk"].isin(["High", "Critical"])]["Student_ID"].tolist()
        options = opts if opts else all_ids

    if signed_in:
        if ss.get("active_id") not in options:
            ss.active_id = options[0]
        st.markdown("### 🎯 Active student")
        if is_student:
            st.selectbox("Your profile", options, key="active_id", disabled=True,
                         format_func=lambda sid: f"{sid} - {name_of.get(sid, sid)}")
            st.caption("🔒 Locked: students can only view their own data.")
        else:
            st.selectbox("Select student (type to search)", options, key="active_id",
                         format_func=lambda sid: f"{sid} - {name_of.get(sid, sid)}")
            if role == "Faculty / Mentor":
                st.caption(f"{len(options)} students in {ss.fac_dept} department.")
            elif role == "Placement Officer":
                st.caption(f"{len(options)} students with High/Critical placement risk.")
            else:
                st.caption(f"{len(options):,} students - full institution access.")
    st.markdown("---")
    st.markdown("### 📂 Dataset")
    st.caption({"upload": "Using: your uploaded file", "local": "Using: Excel file found next to app.py",
                "builtin": "Using: built-in 1000-student dataset"}[source])
    st.caption("Upload option is at the top of this sidebar - drop a .xlsx there to replace the data.")
    st.markdown("---")
    st.markdown("### 🎨 Risk legend")
    st.markdown("".join(badge(r) for r in RISK_ORDER), unsafe_allow_html=True)

# ---- data scope for the signed-in person (drives Command Center, tracker, etc.) ----
if is_student:
    scope_df = df[df["Student_ID"] == ss.get("student_auth")]
    scope_text = "Your personal data only"
elif role == "Faculty / Mentor":
    scope_df = df[df["Department"] == ss.fac_dept]
    scope_text = f"{ss.fac_dept} department - {len(scope_df):,} students"
elif role == "Placement Officer":
    scope_df = df[df["Placement_Risk"].isin(["High", "Critical"])]
    if scope_df.empty:
        scope_df = df
    scope_text = f"Placement caseload - {len(scope_df):,} students with High/Critical placement risk"
else:
    scope_df = df
    scope_text = f"Entire institution - {len(df):,} students"

src_label = {"upload": "Uploaded dataset", "local": "Excel file (local)", "builtin": "Built-in dataset"}[source]

if not signed_in:
    st.markdown(_h(f"""<div class="hero"><div class="hero-title">🎓 Edu<span>Nexus</span> AI</div>
        <div class="hero-tag">Smart Campus Analytics &amp; Student Success Platform</div>
        <span class="hero-chip">🔐 Viewing as: {esc(role)}</span></div>"""), unsafe_allow_html=True)
    st.markdown(card("🔒 Student sign-in required",
                     "Please sign in from the sidebar with your <b>Student ID</b> and <b>PIN</b>. "
                     "After signing in you will only see your own data - Command Center, Student 360, Skills, Wellness and Support plan."),
                unsafe_allow_html=True)
    st.stop()

active_id = ss["active_id"]
S = by_id.loc[active_id]
ss["_cur_student"] = S
cur_status = ss.status.get(active_id, S["Intervention_Status"])

# ── Hero ─────────────────────────────────────────────────────────────────────
st.markdown(_h(f"""<div class="hero"><div class="hero-title">🎓 Edu<span>Nexus</span> AI</div>
    <div class="hero-tag">Smart Campus Analytics &amp; Student Success Platform - The Unified Student Success &amp; Intelligence Hub</div>
    <span class="hero-chip">📊 {src_label} · {len(df):,} students</span>
    <span class="hero-chip">⬆ Upload your own .xlsx from the sidebar</span>
    <span class="hero-chip">🔐 Viewing as: {esc(role)}</span>
    <span class="hero-chip">🔎 Scope: {esc(scope_text)}</span>
    <span class="hero-chip">🎯 Active: {esc(S['Student_Name'])}</span></div>"""), unsafe_allow_html=True)


def active_strip():
    st.markdown(_h(f"""<div class="active-strip">👤 <b style="color:#FFFFFF;font-size:1.05rem;">{esc(S['Student_Name'])}</b>
        &nbsp;<span style="color:#CBD5E1;">{esc(S['Student_ID'])} · {esc(S['Department'])} · Semester {S['Semester']} · CGPA {S['CGPA']:.2f}</span>
        &nbsp;&nbsp;{badge(S['Overall_Risk'])}
        <span style="color:#CBD5E1;font-size:.85rem;">&nbsp;Change student in the sidebar ◀</span></div>"""), unsafe_allow_html=True)


tab_labels = ["🏠 Command Center", "🎓 Student 360", "💼 Skill Gap", "🧠 Mental Health & Wellness", "🤖 AI Copilot", "🎛 What-If Simulator"]
if role == "Placement Officer":
    tab_labels[3] = "🔒 Wellness (restricted)"
if is_student:
    tab_labels[0] = "🏠 My Command Center"
    tab_labels[4] = "🤝 My Support & Buddy"
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(tab_labels)

# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 - COMMAND CENTER  (role-specific: every person only sees data relevant to them)
# ═════════════════════════════════════════════════════════════════════════════
COMPETENCY_LIST = ["Coding", "Aptitude", "Interview", "Communication"]
CONTEXT_METRICS = [("Success", "Success_Score"), ("Attendance", "Attendance_Pct"), ("Aptitude", "Aptitude_Score"),
                   ("Coding", "Coding_Score"), ("Interview", "Mock_Interview_Score"),
                   ("Technical", "Technical_Skill_Score"), ("Soft Skills", "Soft_Skill_Score")]


def render_matrix(sd, allow_lists, key_prefix):
    """AI Cohort Matrix computed on the viewer's own scope `sd`."""
    persona = sd["Behavioral_Persona"]
    total_n = max(len(sd), 1)
    segments = [
        ("⚠️ Low Attendance", persona == "Low Attendance", "#F97316", "Attendance below target - risk of academic slide."),
        ("💼 High Academic / Low Placement", persona == "Placement Risk", "#D946EF", "Solid academics but weak placement readiness."),
        ("✅ Consistent All-Rounders", persona == "Balanced", "#22C55E", "Balanced performance across all signals."),
        ("😴 Low Engagement", persona == "Low Engagement", "#38BDF8", "Low LMS activity and participation."),
        ("📚 Academic Strugglers", persona == "Academic Struggler", "#EF4444", "Backlogs and weak grades need mentoring."),
        ("🏆 High Performers", persona == "High Performer", "#FACC15", "Top performers - ideal peer mentors."),
    ]
    for row_start in (0, 3):
        cols = st.columns(3)
        for col, (title, mask, color, desc) in zip(cols, segments[row_start:row_start + 3]):
            sub = sd[mask]
            avg_s = f"{sub['Success_Score'].mean():.1f}" if len(sub) else "-"
            avg_c = f"{sub['CGPA'].mean():.2f}" if len(sub) else "-"
            col.markdown(_h(f"""<div class="seg" style="--accent:{color}"><div class="seg-title">{title}</div>
                <div class="seg-count">{len(sub):,}</div>
                <div class="seg-desc">{len(sub) / total_n * 100:.1f}% of this view · Avg score {avg_s} · Avg CGPA {avg_c}</div>
                <div class="seg-desc">{desc}</div></div>"""), unsafe_allow_html=True)
            if allow_lists and len(sub):
                with col.expander(f"View {min(len(sub), 50)} students"):
                    show_df(sub.sort_values("Success_Score")[["Student_ID", "Student_Name", "Department", "Success_Score", "CGPA", "Overall_Risk"]].head(50),
                            hide_index=True, height=260)


def context_chart(sd, scope_name, title):
    """Selected student vs the average of the viewer's scope - changes with every student selection."""
    labels = [a for a, _ in CONTEXT_METRICS]
    stu = [float(S[c]) for _, c in CONTEXT_METRICS]
    avg_ = [float(sd[c].mean()) for _, c in CONTEXT_METRICS]
    fig = go.Figure()
    fig.add_bar(x=labels, y=stu, name=S["Student_Name"], marker_color="#D946EF", text=[f"{v:.0f}" for v in stu],
                textposition="outside", textfont=dict(color="#F8FAFC"))
    fig.add_bar(x=labels, y=avg_, name=f"{scope_name} average", marker_color="#38BDF8", text=[f"{v:.0f}" for v in avg_],
                textposition="outside", textfont=dict(color="#F8FAFC"))
    fig.update_layout(barmode="group", title=title, legend=dict(orientation="h", y=-0.15))
    fig.update_yaxes(range=[0, 115], title="Score / %")
    show_fig(style_fig(fig, 380))


with tab1:
    # ───────────────────────── STUDENT: personal command center ─────────────────────────
    if is_student:
        section("My Command Center", f"Your personal dashboard - {esc(S['Student_Name'])} · {esc(S['Department'])} · Semester {S['Semester']}")
        dd = df[df["Department"] == S["Department"]]
        rank = int((dd["Success_Score"] > S["Success_Score"]).sum()) + 1
        pctl = float((df["Success_Score"] < S["Success_Score"]).mean() * 100)
        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(kpi("My Success Score", f"{S['Success_Score']:.1f}", f"dept average {dd['Success_Score'].mean():.1f}", "#A78BFA"), unsafe_allow_html=True)
        k2.markdown(kpi("Rank in Department", f"#{rank}", f"out of {len(dd):,} in {esc(S['Department'])}", "#38BDF8"), unsafe_allow_html=True)
        k3.markdown(kpi("Campus Percentile", f"{pctl:.0f}th", "higher than this % of students", "#22C55E"), unsafe_allow_html=True)
        k4.markdown(kpi("My Overall Risk", esc(S["Overall_Risk"]), f"CGPA {S['CGPA']:.2f} · attendance {S['Attendance_Pct']:.0f}%", RISK_COLORS.get(S["Overall_Risk"], "#64748B")), unsafe_allow_html=True)
        st.write("")
        c1, c2 = st.columns(2)
        with c1:
            context_chart(dd, f"{S['Department']} dept", "Me vs my department average")
        with c2:
            fig = px.histogram(dd, x="Success_Score", nbins=20, title=f"Where I stand in {S['Department']} (Success Score)",
                               color_discrete_sequence=["#38BDF8"])
            fig.add_vline(x=float(S["Success_Score"]), line_color="#D946EF", line_width=4,
                          annotation_text=f"You: {S['Success_Score']:.1f}", annotation_font_color="#F0ABFC")
            fig.update_layout(showlegend=False)
            fig.update_xaxes(title="Success Score")
            fig.update_yaxes(title="Students (anonymous)")
            show_fig(style_fig(fig, 380))
        c3, c4 = st.columns([1, 1.2])
        with c3:
            rc = dd["Overall_Risk"].value_counts().reindex(RISK_ORDER).dropna().reset_index()
            rc.columns = ["Risk", "Students"]
            fig = px.pie(rc, names="Risk", values="Students", hole=0.55, color="Risk", color_discrete_map=RISK_COLORS,
                         category_orders={"Risk": RISK_ORDER}, title=f"{S['Department']} risk mix - you are: {S['Overall_Risk']}")
            fig.update_traces(textposition="outside", textinfo="label+percent", textfont=dict(color="#F8FAFC", size=13),
                              marker=dict(line=dict(color="#0E1117", width=2)), sort=False)
            fig.update_layout(showlegend=False)
            show_fig(style_fig(fig, 340))
        with c4:
            same = int((dd["Behavioral_Persona"] == S["Behavioral_Persona"]).sum())
            st.markdown(card("🧬 My behavioural segment",
                             f"<span style='font-size:1.3rem;font-weight:800;color:#FFFFFF'>{esc(S['Behavioral_Persona'])}</span><br>"
                             f"{same:,} students in {esc(S['Department'])} share this persona.<br><br>"
                             f"Next best action: <b>{esc(S['Recommended_Action_Plan'])}</b><br>"
                             f"<span style='color:#CBD5E1'>Owner: {esc(S['Recommended_Faculty_Role'])}</span>"), unsafe_allow_html=True)
            st.markdown(card("🔒 Privacy", "You only see your own record plus anonymous department-level numbers. Other students' details are never shown."),
                        unsafe_allow_html=True)

    # ───────────────────────── PLACEMENT OFFICER: placement command center ─────────────────────────
    elif role == "Placement Officer":
        cd = scope_df
        section("Placement Command Center", f"Placement readiness across all {len(df):,} students · your caseload = {len(cd):,} students with High/Critical placement risk")
        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(kpi("Students Tracked", f"{len(df):,}", f"{df['Department'].nunique()} departments", "#38BDF8"), unsafe_allow_html=True)
        k2.markdown(kpi("Placement Caseload", f"{len(cd):,}", f"{len(cd) / len(df) * 100:.0f}% need placement support", "#F97316"), unsafe_allow_html=True)
        k3.markdown(kpi("Avg Skill Gap Score", f"{df['Skill_Gap_Score'].mean():.1f}", f"caseload {cd['Skill_Gap_Score'].mean():.1f}", "#FACC15"), unsafe_allow_html=True)
        k4.markdown(kpi("Avg Mock Interview", f"{df['Mock_Interview_Score'].mean():.1f}", f"caseload {cd['Mock_Interview_Score'].mean():.1f}", "#A78BFA"), unsafe_allow_html=True)
        st.write("")
        c1, c2 = st.columns([1, 1.4])
        with c1:
            pr = df["Placement_Risk"].value_counts().reindex(RISK_ORDER).dropna().reset_index()
            pr.columns = ["Risk", "Students"]
            fig = px.pie(pr, names="Risk", values="Students", hole=0.55, color="Risk", color_discrete_map=RISK_COLORS,
                         category_orders={"Risk": RISK_ORDER}, title="Placement Risk Distribution")
            fig.update_traces(textposition="outside", textinfo="label+percent", textfont=dict(color="#F8FAFC", size=13),
                              marker=dict(line=dict(color="#0E1117", width=2)), sort=False)
            fig.update_layout(showlegend=False)
            show_fig(style_fig(fig, 360))
        with c2:
            dr = df.groupby(["Department", "Placement_Risk"]).size().reset_index(name="Students")
            fig = px.bar(dr, x="Department", y="Students", color="Placement_Risk", color_discrete_map=RISK_COLORS,
                         category_orders={"Placement_Risk": RISK_ORDER}, title="Department-wise Placement Risk", text="Students")
            fig.update_traces(textfont=dict(color="#0B1220"))
            fig.update_layout(barmode="stack", legend_title_text="Placement risk")
            show_fig(style_fig(fig, 360))
        c3, c4 = st.columns(2)
        with c3:
            ra = cd.groupby("Target_Role")["Skill_Gap_Score"].mean().reset_index().sort_values("Skill_Gap_Score")
            fig = px.bar(ra, x="Skill_Gap_Score", y="Target_Role", orientation="h", title="Caseload: avg skill-gap score by target role",
                         text=ra["Skill_Gap_Score"].round(1), color_discrete_sequence=["#D946EF"])
            fig.update_traces(textposition="outside", textfont=dict(color="#F8FAFC"))
            fig.update_yaxes(title="")
            show_fig(style_fig(fig, 340))
        with c4:
            counts = {c: cd["Missing_Skills"].str.contains(c, na=False).sum() for c in COMPETENCY_LIST}
            cdf_ = pd.DataFrame({"Competency": list(counts), "Students": list(counts.values())})
            fig = px.bar(cdf_, x="Competency", y="Students", title="Caseload: competency gaps flagged", text="Students",
                         color_discrete_sequence=["#38BDF8"])
            fig.update_traces(textposition="outside", textfont=dict(color="#F8FAFC"))
            show_fig(style_fig(fig, 340))
        sc = cd.groupby("Department")[["Aptitude_Score", "Coding_Score", "Mock_Interview_Score"]].mean().reset_index().melt("Department", var_name="Score", value_name="Average")
        sc["Score"] = sc["Score"].str.replace("_Score", "").str.replace("_", " ")
        fig = px.bar(sc, x="Department", y="Average", color="Score", barmode="group", title="Caseload: placement scores by department",
                     color_discrete_sequence=["#38BDF8", "#D946EF", "#FACC15"])
        fig.update_yaxes(range=[0, 100])
        show_fig(style_fig(fig, 340))
        section("🧬 AI Cohort Matrix - placement caseload", "Behavioural segments inside your caseload.")
        render_matrix(cd, True, "po")
        section("🔍 Selected student in context", "Changes whenever you pick another student in the sidebar.")
        context_chart(cd, "Caseload", "Selected student vs placement-caseload average")

    # ───────────────────────── ADMINISTRATOR / FACULTY ─────────────────────────
    else:
        sd = scope_df
        is_fac = role == "Faculty / Mentor"
        section("Command Center" if not is_fac else f"Command Center - {esc(ss.fac_dept)} Department",
                f"Aggregated view of {len(sd):,} students - {esc(scope_text)}.")
        total = len(sd)
        n_crit = int((sd["Overall_Risk"] == "Critical").sum())
        n_highrisk = int(sd["Overall_Risk"].isin(["High", "Critical"]).sum())
        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(kpi("Total Students", f"{total:,}", f"{sd['Department'].nunique()} department(s)", "#38BDF8"), unsafe_allow_html=True)
        k2.markdown(kpi("Avg Success Score", f"{sd['Success_Score'].mean():.1f}", f"Avg CGPA {sd['CGPA'].mean():.2f}", "#A78BFA"), unsafe_allow_html=True)
        k3.markdown(kpi("High Risk Count", f"{n_highrisk:,}", f"{n_highrisk / max(total, 1) * 100:.0f}% of this view (High + Critical)", "#F97316"), unsafe_allow_html=True)
        k4.markdown(kpi("Critical Cases", f"{n_crit}", "need action today", "#EF4444"), unsafe_allow_html=True)
        st.write("")
        c1, c2 = st.columns([1, 1.4])
        with c1:
            rc = sd["Overall_Risk"].value_counts().reindex(RISK_ORDER).dropna().reset_index()
            rc.columns = ["Risk", "Students"]
            fig = px.pie(rc, names="Risk", values="Students", hole=0.55, color="Risk", color_discrete_map=RISK_COLORS,
                         category_orders={"Risk": RISK_ORDER}, title="Overall Risk Distribution")
            fig.update_traces(textposition="outside", textinfo="label+percent", textfont=dict(color="#F8FAFC", size=13),
                              marker=dict(line=dict(color="#0E1117", width=2)), sort=False)
            fig.update_layout(showlegend=False)
            fig.add_annotation(text=f"<b>{total:,}</b><br>students", showarrow=False, font=dict(size=16, color="#FFFFFF"))
            show_fig(style_fig(fig, 360))
        with c2:
            if is_fac:
                dr = sd.groupby(["Semester", "Overall_Risk"]).size().reset_index(name="Students")
                fig = px.bar(dr, x="Semester", y="Students", color="Overall_Risk", color_discrete_map=RISK_COLORS,
                             category_orders={"Overall_Risk": RISK_ORDER}, title=f"{ss.fac_dept}: Risk by Semester", text="Students")
                fig.update_xaxes(type="category")
            else:
                dr = sd.groupby(["Department", "Overall_Risk"]).size().reset_index(name="Students")
                fig = px.bar(dr, x="Department", y="Students", color="Overall_Risk", color_discrete_map=RISK_COLORS,
                             category_orders={"Overall_Risk": RISK_ORDER}, title="Department Risk Breakdown", text="Students")
            fig.update_traces(textfont=dict(color="#0B1220"))
            fig.update_layout(barmode="stack", legend_title_text="Risk")
            show_fig(style_fig(fig, 360))
        section("🧬 AI Cohort Matrix", "Students grouped by Behavioral_Persona so you can target the right support at scale.")
        render_matrix(sd, True, "adm")
        section("🔍 Selected student in context", "Changes whenever you pick another student in the sidebar.")
        context_chart(sd, "Department" if is_fac else "Institution", "Selected student vs " + ("department" if is_fac else "institution") + " average")

# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 - STUDENT 360
# ═════════════════════════════════════════════════════════════════════════════
with tab2:
    section(f"Student 360 - {esc(S['Student_Name'])}", f"{esc(S['Student_ID'])} · {esc(S['Department'])} · Semester {S['Semester']} · {esc(S['Gender'])}")

    # ---- Email inbox (messages sent from the AI Copilot tab) ----
    msgs = ss.inbox.get(active_id, [])
    unread = [m for m in msgs if not m["read"]]
    if msgs:
        latest = msgs[-1]
        head = f"📬 {len(unread)} NEW message{'s' if len(unread) != 1 else ''} from your faculty" if unread else "📭 Inbox - all messages read"
        cls = "inbox-new" if unread else "inbox-empty"
        st.markdown(_h(f"""<div class="{cls}"><div style="font-size:1.15rem;font-weight:800;color:#FFFFFF;">{head}</div>
            <div style="color:#E2E8F0;font-size:.85rem;margin-top:4px;">From: {esc(latest['from'])} · {esc(latest['time'])} · Subject: {esc(latest['subject'])}</div>
            <div class="mail">{esc(latest['body'])}</div></div>"""), unsafe_allow_html=True)
        if unread:
            st.button("✔ Mark messages as read", key="mark_read", on_click=cb_mark_read, args=(active_id,))
        if len(msgs) > 1:
            with st.expander(f"Earlier messages ({len(msgs) - 1})"):
                for m in reversed(msgs[:-1]):
                    st.markdown(_h(f'<div class="card-title">{esc(m["time"])} · {esc(m["from"])}</div><div class="mail">{esc(m["body"])}</div>'), unsafe_allow_html=True)
    else:
        st.markdown(_h('<div class="inbox-empty">📭 <b>Inbox</b> - no messages yet. Messages sent by faculty from the AI Copilot tab will appear here instantly.</div>'),
                    unsafe_allow_html=True)

    # ---- Dynamic KPIs ----
    a1, a2, a3, a4, a5 = st.columns(5)
    cg_avg = df["CGPA"].mean()
    a1.markdown(kpi("Success Score", f"{S['Success_Score']:.1f}", f"cohort avg {df['Success_Score'].mean():.1f}", "#A78BFA"), unsafe_allow_html=True)
    a2.markdown(kpi("CGPA", f"{S['CGPA']:.2f}", f"cohort avg {cg_avg:.2f} · {S['CGPA'] - cg_avg:+.2f}", "#38BDF8"), unsafe_allow_html=True)
    a3.markdown(kpi("Academic Risk", esc(S["Academic_Risk"]), f"Backlogs: {S['Backlogs']}", RISK_COLORS.get(S["Academic_Risk"], "#64748B")), unsafe_allow_html=True)
    a4.markdown(kpi("Placement Risk", esc(S["Placement_Risk"]), f"Target: {esc(S['Target_Role'])}", RISK_COLORS.get(S["Placement_Risk"], "#64748B")), unsafe_allow_html=True)
    a5.markdown(kpi("Attendance", f"{S['Attendance_Pct']:.1f}%", f"drop {S['Attendance_Drop_Pct']:.1f} pts", "#22C55E" if S["Attendance_Pct"] >= 75 else "#F97316"), unsafe_allow_html=True)
    st.write("")

    g1, g2 = st.columns([1, 1.6])
    with g1:
        avg = df["Success_Score"].mean()
        show_fig(gauge(S["Success_Score"], "Success Score (0-100)",
                       [(0, 45, "rgba(239,68,68,.55)"), (45, 60, "rgba(249,115,22,.55)"), (60, 75, "rgba(250,204,21,.55)"), (75, 100, "rgba(34,197,94,.55)")], ref=avg))
        st.caption(f"White line = cohort average ({avg:.1f}). This student is {S['Success_Score'] - avg:+.1f} vs average.")
    with g2:
        st.markdown(card("Risk flags", f"Overall {badge(S['Overall_Risk'])} Academic {badge(S['Academic_Risk'])} "
                                       f"Placement {badge(S['Placement_Risk'])} Burnout {badge(S['Academic_Burnout_Stress_Risk'], 'burn')}"),
                    unsafe_allow_html=True)
        st.markdown(card("⚠️ WHY AT RISK?", esc(S["AI_Risk_Explanation"]) +
                         f"<br><span style='color:#CBD5E1;font-size:.9rem'>Persona: <b>{esc(S['Behavioral_Persona'])}</b> · Priority: <b>{esc(S['Intervention_Priority'])}</b></span>"),
                    unsafe_allow_html=True)
        st.markdown(card("🎯 NEXT BEST ACTION", f"<b>{esc(S['Recommended_Action_Plan'])}</b><br>"
                                                 f"<span style='color:#CBD5E1'>Owner: <b style='color:#FFFFFF'>{esc(S['Recommended_Faculty_Role'])}</b> · "
                                                 f"Estimated recovery: {S['Estimated_Recovery_Time_Weeks']} weeks · Status: <b style='color:#FFFFFF'>{esc(cur_status)}</b></span>"),
                    unsafe_allow_html=True)

    st.markdown("---")
    section("💼 Placement snapshot", f"Placement readiness for target role: <b>{esc(S['Target_Role'])}</b>")
    q1, q2, q3, q4 = st.columns(4)
    q1.markdown(kpi("Placement Risk", esc(S["Placement_Risk"]), "owner: Placement Cell", RISK_COLORS.get(S["Placement_Risk"], "#64748B")), unsafe_allow_html=True)
    q2.markdown(kpi("Coding Score", f"{S['Coding_Score']:.1f}", f"cohort avg {df['Coding_Score'].mean():.1f}", "#22C55E" if S["Coding_Score"] >= df["Coding_Score"].mean() else "#F97316"), unsafe_allow_html=True)
    q3.markdown(kpi("Aptitude Score", f"{S['Aptitude_Score']:.1f}", f"cohort avg {df['Aptitude_Score'].mean():.1f}", "#22C55E" if S["Aptitude_Score"] >= df["Aptitude_Score"].mean() else "#F97316"), unsafe_allow_html=True)
    q4.markdown(kpi("Mock Interview", f"{S['Mock_Interview_Score']:.1f}", f"cohort avg {df['Mock_Interview_Score'].mean():.1f}", "#22C55E" if S["Mock_Interview_Score"] >= df["Mock_Interview_Score"].mean() else "#F97316"), unsafe_allow_html=True)
    _gaps = [m for m in split_list(S["Missing_Skills"]) if m.lower() != "no major gap"]
    _focus = (f"Focus areas: <b>{esc(', '.join(_gaps))}</b>. Placement Cell should schedule targeted practice."
              if _gaps else "No major placement gap - ready for campus drives.")
    st.markdown(card("Placement readiness", f"Skill gap level {badge(S['Skill_Gap_Level'])} <span style='color:#CBD5E1'>score {S['Skill_Gap_Score']:.1f}</span> · "
                                           f"Technical skill {S['Technical_Skill_Score']:.0f} · Soft skill {S['Soft_Skill_Score']:.0f}<br>{_focus}"),
                unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# TAB 3 - SKILL GAP
# ═════════════════════════════════════════════════════════════════════════════
with tab3:
    section("Skill Gap Analysis", "Current skills vs the skills required for the target role.")
    active_strip()
    cur_sk, req_sk = split_list(S["Current_Skills"]), split_list(S["Required_Skills"])
    match = skill_match(req_sk, cur_sk)
    pct = sum(match.values()) / len(req_sk) * 100 if req_sk else 0
    miss_comp = [m for m in split_list(S["Missing_Skills"]) if m.lower() != "no major gap"]

    l, r = st.columns([1.05, 1])
    with l:
        st.markdown(card("🎯 Target role", f"<span style='font-size:1.3rem;font-weight:800;color:#FFFFFF'>{esc(S['Target_Role'])}</span> "
                                          f"&nbsp;Skill gap: {badge(S['Skill_Gap_Level'])} <span style='color:#CBD5E1'>score {S['Skill_Gap_Score']:.1f}</span>"),
                    unsafe_allow_html=True)
        st.markdown(card("✅ Current skills", "".join(f'<span class="chip chip-neutral">{esc(c)}</span>' for c in cur_sk) or "None recorded"), unsafe_allow_html=True)
        st.markdown(card("📋 Required skills (green = have · red = missing)",
                         "".join(f'<span class="chip {"chip-ok" if ok else "chip-bad"}">{"✔" if ok else "✖"} {esc(k)}</span>' for k, ok in match.items()) or "None"),
                    unsafe_allow_html=True)
        WEAK_RULES = {"Coding": ("Coding_Score", 60), "Aptitude": ("Aptitude_Score", 60),
                      "Interview": ("Mock_Interview_Score", 60), "Communication": ("Soft_Skill_Score", 55)}
        tech_missing = [k for k, ok in match.items() if not ok]
        parts = []
        if tech_missing:
            parts.append("<b>Technical skills to learn</b><br>" + "".join(f'<span class="chip chip-bad">✖ {esc(k)}</span>' for k in tech_missing))
        flagged_html = ""
        for comp in miss_comp:
            col_ = WEAK_RULES.get(comp, (None, 0))[0]
            flagged_html += f'<span class="chip chip-bad">⛔ {esc(comp)}{f" · {S[col_]:.0f}/100" if col_ else ""}</span>'
        if flagged_html:
            parts.append("<b>Readiness gaps flagged for this student</b><br>" + flagged_html)
        watch_html = "".join(f'<span class="chip chip-warn">👀 {c} · {S[col_]:.0f}/100</span>'
                             for c, (col_, th) in WEAK_RULES.items() if c not in miss_comp and S[col_] < th)
        if watch_html:
            parts.append("<b>Watch list (score below benchmark)</b><br>" + watch_html)
        if parts:
            weakest = min(WEAK_RULES.items(), key=lambda kv: S[kv[1][0]])
            parts.append(f"<b>Learn first:</b> {weakest[0]} (score {S[weakest[1][0]]:.0f})")
        st.markdown(card("⛔ Missing skills (to learn)", "<br>".join(parts) or '<span class="chip chip-ok">No major gap - all required skills covered</span>'),
                    unsafe_allow_html=True)
        n_have, n_req = sum(match.values()), len(req_sk)
        tech_s, gap_s = float(S["Technical_Skill_Score"]), float(S["Skill_Gap_Score"])
        match_score = float(np.clip(0.5 * pct + 0.3 * tech_s + 0.2 * (100 - gap_s), 0, 100))
        show_fig(gauge(match_score, "Required-skill match", [(0, 40, "rgba(239,68,68,.55)"), (40, 70, "rgba(250,204,21,.55)"), (70, 100, "rgba(34,197,94,.55)")], suffix="%"))
        st.caption(f"{n_have} of {n_req} required skills covered ({pct:.0f}%) · blended with this student's Technical score ({tech_s:.0f}) and skill-gap score ({gap_s:.1f}).")
    with r:
        if len(df) >= 50:
            baseline = {col: float(df[col].mean()) for _, col in RADAR_AXES}
        else:
            baseline = dict(RADAR_BASELINE_FALLBACK)
        labels = [a for a, _ in RADAR_AXES]
        stu_vals = [float(S[c]) for _, c in RADAR_AXES]
        base_vals = [baseline[c] for _, c in RADAR_AXES]
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(r=base_vals + base_vals[:1], theta=labels + labels[:1], name="Cohort average",
                                      line=dict(color="#38BDF8", dash="dash", width=2), fill="toself", fillcolor="rgba(56,189,248,0.12)"))
        fig.add_trace(go.Scatterpolar(r=stu_vals + stu_vals[:1], theta=labels + labels[:1], name=S["Student_Name"],
                                      line=dict(color="#D946EF", width=3), fill="toself", fillcolor="rgba(217,70,239,0.30)"))
        fig.update_layout(title="Skill Radar - student vs cohort baseline",
                          polar=dict(bgcolor="rgba(0,0,0,0)", radialaxis=dict(range=[0, 100], gridcolor="#334155", tickfont=dict(color="#CBD5E1"), linecolor="#334155"),
                                     angularaxis=dict(tickfont=dict(color="#FFFFFF", size=14), gridcolor="#334155")),
                          legend=dict(orientation="h", y=-0.08))
        show_fig(style_fig(fig, 470))
    cmp_rows = []
    for (lab, col), b in zip(RADAR_AXES, base_vals):
        v = float(S[col])
        cmp_rows.append({"Competency": lab, "Student": round(v, 1), "Cohort average": round(b, 1), "Difference": round(v - b, 1),
                         "Status": "Above average" if v >= b else "Below average"})
    show_df(pd.DataFrame(cmp_rows), hide_index=True)

# ═════════════════════════════════════════════════════════════════════════════
# TAB 4 - MENTAL HEALTH & WELLNESS
# ═════════════════════════════════════════════════════════════════════════════
with tab4:
    if role == "Placement Officer":
        section("Mental Health & Wellness", "Restricted")
        st.markdown(card("🔒 Restricted for Placement Officers",
                         "Wellness and stress data is confidential. It is visible only to Administrators, Faculty / Mentors and the student themselves."),
                    unsafe_allow_html=True)
    else:
        section("Mental Health & Wellness", "Academic Burnout / Stress Risk Index, LMS engagement and attendance drops.")
        active_strip()
        idx_col, lab_col = "Academic_Burnout_Stress_Risk_Index", "Academic_Burnout_Stress_Risk"
        mins = df.groupby(lab_col)[idx_col].min()
        mod_start, high_start = float(mins.get("Moderate", 30)), float(mins.get("High", 50))
        if not mod_start < high_start:
            mod_start, high_start = 30.0, 50.0
        level = S[lab_col]
        lms_avg, lms_q25 = float(df["LMS_Weekly_Logins"].mean()), float(df["LMS_Weekly_Logins"].quantile(0.25))

        w1, w2, w3, w4 = st.columns(4)
        w1.markdown(kpi("Burnout / Stress Index", f"{S[idx_col]:.1f}", f"level: {esc(level)}", BURN_COLORS.get(level, "#64748B")), unsafe_allow_html=True)
        w2.markdown(kpi("LMS Logins / week", f"{S['LMS_Weekly_Logins']:.1f}", f"cohort avg {lms_avg:.1f}", "#A78BFA"), unsafe_allow_html=True)
        w3.markdown(kpi("Attendance Drop", f"{S['Attendance_Drop_Pct']:.1f} pts", f"{S['Previous_Attendance_Pct']:.1f}% → {S['Attendance_Pct']:.1f}%", "#F97316" if S["Attendance_Drop_Pct"] >= 10 else "#22C55E"), unsafe_allow_html=True)
        w4.markdown(kpi("Active Backlogs", f"{S['Backlogs']}", "academic pressure", "#EF4444" if S["Backlogs"] >= 3 else "#38BDF8"), unsafe_allow_html=True)
        st.write("")

        p1, p2, p3 = st.columns([1, 1, 1])
        with p1:
            show_fig(gauge(S[idx_col], "Burnout / Stress Risk Index",
                           [(0, mod_start, "rgba(34,197,94,.55)"), (mod_start, high_start, "rgba(250,204,21,.55)"), (high_start, 100, "rgba(239,68,68,.55)")], ref=df[idx_col].mean(), num_size=24))
        with p2:
            fig = go.Figure(go.Bar(x=["Previous", "Current"], y=[S["Previous_Attendance_Pct"], S["Attendance_Pct"]],
                                   marker_color=["#64748B", "#EF4444" if S["Attendance_Drop_Pct"] >= 10 else "#38BDF8"],
                                   text=[f"{S['Previous_Attendance_Pct']:.1f}%", f"{S['Attendance_Pct']:.1f}%"], textposition="outside", textfont=dict(color="#F8FAFC")))
            fig.update_layout(title=f"Attendance (drop {S['Attendance_Drop_Pct']:.1f} pts)", showlegend=False)
            fig.update_yaxes(range=[0, 115], title="%")
            show_fig(style_fig(fig, 265))
        with p3:
            fig = go.Figure(go.Bar(x=["This student", "Cohort average", "Cohort bottom quartile"], y=[S["LMS_Weekly_Logins"], lms_avg, lms_q25],
                                   marker_color=["#D946EF", "#38BDF8", "#64748B"],
                                   text=[f"{S['LMS_Weekly_Logins']:.1f}", f"{lms_avg:.1f}", f"{lms_q25:.1f}"], textposition="outside", textfont=dict(color="#F8FAFC")))
            fig.update_layout(title="LMS weekly logins", showlegend=False)
            fig.update_yaxes(range=[0, max(14, S["LMS_Weekly_Logins"] + 2)])
            show_fig(style_fig(fig, 265))

        signals = []
        if S["Attendance_Drop_Pct"] >= 10:
            signals.append(f"📉 Attendance fell <b>{S['Attendance_Drop_Pct']:.1f} points</b> ({S['Previous_Attendance_Pct']:.1f}% → {S['Attendance_Pct']:.1f}%)")
        if S["Backlogs"] >= 2:
            signals.append(f"📚 <b>{S['Backlogs']} active backlogs</b> are adding academic pressure")
        if S["LMS_Weekly_Logins"] <= lms_q25:
            signals.append(f"💻 LMS logins ({S['LMS_Weekly_Logins']:.1f}/week) are in the cohort's bottom quartile")
        if S["Attendance_Pct"] < 65:
            signals.append(f"🏫 Attendance of <b>{S['Attendance_Pct']:.1f}%</b> is below the 65% comfort line")
        sig_html = "<br>".join(signals) if signals else "✅ No acute stress signals detected."

        if str(level) in ("High", "Critical"):
            color, title = "#EF4444", "🚨 WARNING - High stress / burnout risk"
            steps = ["Refer to the Student Counselor within 48 hours", "Hold a one-to-one check-in with the faculty mentor this week",
                     "Reduce academic load: agree a backlog-clearance timetable", "Activate the peer buddy and notify parent/guardian if advised"]
        elif str(level) == "Moderate":
            color, title = "#FACC15", "⚠️ WATCH - Moderate stress level"
            steps = ["Schedule a mentor check-in within two weeks", "Set weekly attendance and LMS targets", "Encourage peer-buddy study sessions"]
        else:
            color, title = "#22C55E", "✅ STABLE - Low stress level"
            steps = ["No intervention needed - continue routine mentoring", "Keep monitoring attendance and LMS activity monthly"]
        plan_html = "".join(f"<div>• {esc(x)}</div>" for x in steps)
        st.markdown(_h(f"""<div class="card" style="border-left:8px solid {color};"><div class="card-title" style="color:{color};">Dynamic wellness action plan</div>
            <div class="card-text"><div style="font-size:1.2rem;font-weight:800;color:#FFFFFF;margin-bottom:6px;">{title}</div>{plan_html}</div></div>"""),
            unsafe_allow_html=True)
        st.markdown(card("🔎 Detected wellness signals", sig_html), unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# TAB 5 - AI COPILOT
# ═════════════════════════════════════════════════════════════════════════════
with tab5:
    section("AI Copilot - Teacher Email System", "Review the AI draft, send it to the student, and track the workflow live.")
    active_strip()
    flash = ss.pop("flash", None)
    if flash:
        st.success(flash[1])

    left, right = st.columns([1.1, 1])
    with left:
        if is_student:
            st.markdown(card("🔒 Staff tool", "The email composer is available to faculty, placement officers and administrators. "
                                             "Any message sent to you appears in your <b>Student 360</b> inbox."), unsafe_allow_html=True)
            st.markdown(card("Your intervention status", f"<b>{esc(cur_status)}</b> · {esc(ss.activity.get(active_id, 'No activity yet'))}"), unsafe_allow_html=True)
        else:
            draft_key = f"draft_text_{active_id}"
            if draft_key not in ss:
                ss[draft_key] = ss.drafts.get(active_id) or default_email(S, role)
            st.markdown(f'<div class="sec-title" style="font-size:1.1rem;">✉️ AI-generated outreach to {esc(S["Student_Name"])}</div>', unsafe_allow_html=True)
            st.text_area("Email body (editable)", key=draft_key, height=300)
            st.caption(f"Source column Outreach_Message_Draft: {S['Outreach_Message_Draft']}")
            b1, b2, b3 = st.columns(3)
            b1.button("📧 Send Email to Student (Mark In Progress)", key=f"send_{active_id}", on_click=cb_send_email, args=(active_id, role, draft_key))
            b2.button("✅ Mark Completed", key=f"done_{active_id}", on_click=cb_complete, args=(active_id, role))
            b3.button("🔄 Regenerate AI draft", key=f"regen_{active_id}", on_click=cb_regen, args=(draft_key, active_id, role))
            st.markdown(f"Current status: **{cur_status}**")

        st.markdown("---")
        st.markdown('<div class="sec-title" style="font-size:1.1rem;">🤝 Peer Buddy Match</div>', unsafe_allow_html=True)
        pb1, pb2 = st.columns([1.2, 1])
        with pb1:
            B = by_id.loc[S["Peer_Buddy_ID"]] if S["Peer_Buddy_ID"] in by_id.index else None
            extra = ""
            if B is not None:
                extra = (f"<br><span style='color:#CBD5E1'>{esc(B['Student_ID'])} · {esc(B['Department'])} · Semester {B['Semester']}<br>"
                         f"Success score {B['Success_Score']:.1f} · CGPA {B['CGPA']:.2f}</span>")
            st.markdown(card("Recommended mentor", f"<span style='font-size:1.3rem;font-weight:800;color:#FFFFFF'>{esc(S['Peer_Buddy_Name'])}</span>{extra}"), unsafe_allow_html=True)
        with pb2:
            show_fig(gauge(S["Peer_Match_Score"], "Match score", [(0, 40, "rgba(239,68,68,.55)"), (40, 70, "rgba(250,204,21,.55)"), (70, 100, "rgba(34,197,94,.55)")], num_size=24))

    with right:
        st.markdown('<div class="sec-title" style="font-size:1.1rem;">📋 Live Intervention Tracker</div>', unsafe_allow_html=True)
        st.caption(f"Scope: {scope_text}")
        stat_counts = pd.Series({sid_: ss.status[sid_] for sid_ in scope_df["Student_ID"]}).value_counts()
        m1, m2, m3 = st.columns(3)
        m1.markdown(kpi("Pending", int(stat_counts.get("Pending", 0)), "", STATUS_COLORS["Pending"]), unsafe_allow_html=True)
        m2.markdown(kpi("In Progress", int(stat_counts.get("In Progress", 0)), "", STATUS_COLORS["In Progress"]), unsafe_allow_html=True)
        m3.markdown(kpi("Completed", int(stat_counts.get("Completed", 0)), "", STATUS_COLORS["Completed"]), unsafe_allow_html=True)
        if is_student:
            view_opts = ["My workflow"]
        else:
            view_opts = ["Updated this session", "Active P1-P2 workflows", "All students"]
        mode = st.radio("Show", view_opts, horizontal=True, key="trk_mode", index=min(1, len(view_opts) - 1) if not is_student else 0)

        t = scope_df[["Student_ID", "Student_Name", "Department", "Overall_Risk", "Intervention_Priority", "Success_Score"]].copy()
        t["Status"] = t["Student_ID"].map(ss.status)
        t["Last Action"] = t["Student_ID"].map(ss.activity).fillna("-")
        t["Touched"] = t["Student_ID"].isin(ss.activity.keys())
        if mode == "My workflow":
            t = t[t["Student_ID"] == active_id]
        elif mode == "Updated this session":
            t = t[t["Touched"]]
        elif mode == "Active P1-P2 workflows":
            t = t[(t["Intervention_Priority"].str.startswith(("P1", "P2")) & (t["Status"] != "Completed")) | t["Touched"]]
        t = t.sort_values(["Touched", "Intervention_Priority", "Success_Score"], ascending=[False, True, True]).head(200)
        shown = t.drop(columns=["Touched", "Success_Score"]).rename(columns={"Student_ID": "ID", "Student_Name": "Student", "Overall_Risk": "Risk", "Intervention_Priority": "Priority"})
        if shown.empty:
            st.info("Nothing here yet - send an email on the left and the workflow appears instantly.")
        else:
            def _color(v):
                c = {"Pending": "#7F1D1D", "In Progress": "#0C4A6E", "Completed": "#14532D"}.get(v)
                return f"background-color:{c};color:#FFFFFF;font-weight:700" if c else ""
            try:
                styled = shown.style.map(_color, subset=["Status"])
            except Exception:
                try:
                    styled = shown.style.applymap(_color, subset=["Status"])
                except Exception:
                    styled = shown
            show_df(styled, hide_index=True, height=470)

# ═════════════════════════════════════════════════════════════════════════════
# TAB 6 - WHAT-IF SIMULATOR
# ═════════════════════════════════════════════════════════════════════════════
with tab6:
    section("What-If Simulator", "See how the Success Score changes if attendance reaches 80% or skills improve by 10 points.")
    active_strip()
    base = float(S["Success_Score"])
    att80 = float(S["What_If_Attendance_80_Score"])
    plus10 = float(S["What_If_Score_Plus_10"])
    d_att, d_skill = att80 - base, plus10 - base
    x1, x2, x3 = st.columns(3)
    x1.markdown(kpi("Current Success Score", f"{base:.1f}", f"attendance {S['Attendance_Pct']:.1f}%", "#64748B"), unsafe_allow_html=True)
    x2.markdown(kpi("If attendance reaches 80%", f"{att80:.1f}", f"{d_att:+.1f} points", "#38BDF8"), unsafe_allow_html=True)
    x3.markdown(kpi("If skills improve +10 pts", f"{plus10:.1f}", f"{d_skill:+.1f} points", "#22C55E"), unsafe_allow_html=True)
    st.write("")
    y1, y2 = st.columns([1.3, 1])
    with y1:
        fig = go.Figure(go.Bar(x=["Current", "Attendance → 80%", "Skills +10 pts"], y=[base, att80, plus10],
                               marker_color=["#64748B", "#38BDF8", "#22C55E"],
                               text=[f"{base:.1f}", f"{att80:.1f} ({d_att:+.1f})", f"{plus10:.1f} ({d_skill:+.1f})"],
                               textposition="outside", textfont=dict(color="#F8FAFC", size=14)))
        fig.update_layout(title=f"Projected Success Score - {S['Student_Name']}", showlegend=False)
        fig.update_yaxes(range=[max(0, base - 15), min(105, max(plus10, att80) + 10)], title="Success Score")
        show_fig(style_fig(fig, 380))
    with y2:
        if S["Attendance_Pct"] >= 80:
            att_msg = f"Attendance is already <b>{S['Attendance_Pct']:.1f}%</b> (above 80%), so raising it gives no extra uplift."
        else:
            att_msg = f"Raising attendance from <b>{S['Attendance_Pct']:.1f}%</b> to <b>80%</b> adds <b>{d_att:+.1f}</b> points."
        best = "improving skills (+10 pts)" if d_skill >= d_att else "raising attendance to 80%"
        st.markdown(card("📈 Insight", f"{att_msg}<br>A +10 skill improvement adds <b>{d_skill:+.1f}</b> points.<br>"
                                     f"Biggest lever for this student: <b>{best}</b>."), unsafe_allow_html=True)
        st.markdown(card("⏱ Recovery outlook", f"Estimated recovery time: <b>{S['Estimated_Recovery_Time_Weeks']} weeks</b><br>"
                                              f"Improvement potential: <b>{S['What_If_Improvement_Potential']} points</b><br>"
                                              f"Recommended plan: {esc(S['Recommended_Action_Plan'])}"), unsafe_allow_html=True)
        show_fig(gauge(plus10, "Best-case score (+10 skills)", [(0, 45, "rgba(239,68,68,.55)"), (45, 60, "rgba(249,115,22,.55)"), (60, 75, "rgba(250,204,21,.55)"), (75, 100, "rgba(34,197,94,.55)")], ref=base))

st.markdown("---")
st.caption("EduNexus AI · Bytexel Hackathon · Built with Streamlit & Plotly")
