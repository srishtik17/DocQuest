import os
import time
import uuid
import base64
from io import BytesIO
from pathlib import Path

import streamlit as st
import numpy as np
import faiss
import pandas as pd
import io
import json
import html
import warnings
import plotly.express as px

from dotenv import load_dotenv
from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation
from google import genai
from google.genai import types
from langchain_text_splitters import RecursiveCharacterTextSplitter


# CONFIGURATION

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

API_KEY = os.getenv("GEMINI_API", "")
if not API_KEY:
    try:
        API_KEY = st.secrets["GEMINI_API"]
    except Exception:
        API_KEY = 

EMBED_MODEL = "gemini-embedding-001"

CHAT_MODELS = [
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash"
]


# DashQuest dashboard configuration
AGGS = ["sum", "mean", "count", "min", "max", "nunique"]
KINDS = ["kpi", "bar", "line", "pie", "scatter", "hist", "table"]
NEEDS = {
    "bar": ["x"],
    "line": ["x"],
    "pie": ["x"],
    "scatter": ["x", "y"],
    "hist": ["x"],
}
PALETTE = ["#8a5fcc", "#ee9fdc", "#6fb7e9", "#f2b36d", "#7fd1ae", "#c3a5ec", "#e57fa0"]

CHUNK_SIZE = 1000
CHUNK_OVERLAP = 200
TOP_K = 5

SUPPORTED_FILES = ["pdf", "docx", "pptx", "txt", "csv", "xlsx"]

st.set_page_config(
    page_title="DocQuest",
    page_icon="🧸",
    layout="wide"
)


# THEME: KUCHU PUCHU BEARS (UI ONLY)
# Optional: drop your own bear artwork in an "assets" folder next to this file:
#   assets/sidebar_bear.png   (top of sidebar)
#   assets/sleeping_bear.png  (empty state in the middle)
#   assets/books_bear.png     (bottom-right corner, bear on books)
#   assets/footer_bear.png    (bottom of sidebar)
# If a file is missing, a built-in cartoon bear is drawn instead.

ASSET_DIR = Path(__file__).resolve().parent / "assets"

THEME_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&family=Caveat:wght@600&display=swap');

:root {
    --lav-50: #fdfbff;
    --lav-100: #f3e9ff;
    --lav-200: #e6d6fb;
    --lav-300: #d3b8f2;
    --purple: #9a72d6;
    --purple-dark: #5f3f99;
    --ink: #3a2a5c;
    --pink: #f7c6e8;
    --pink-deep: #f1b4e3;
    --mint-bg: #eaf8f1;
    --mint-border: #bfe8d3;
    --mint-ink: #2f7a58;
}

html, body, .stApp, .stMarkdown, p, label, button, input, textarea, li, h1, h2, h3, h4 {
    font-family: 'Nunito', sans-serif;
}
.stApp { font-size: 15px; }

.stApp {
    color: var(--ink);
    background:
        radial-gradient(900px 500px at 15% 0%, rgba(214,190,248,.55), transparent 60%),
        radial-gradient(800px 500px at 100% 100%, rgba(247,198,232,.5), transparent 60%),
        linear-gradient(rgba(167,139,250,.08) 1px, transparent 1px),
        linear-gradient(90deg, rgba(167,139,250,.08) 1px, transparent 1px),
        linear-gradient(160deg, #f0e4fd 0%, #f4e6fb 50%, #fbe6f4 100%);
    background-size: auto, auto, 38px 38px, 38px 38px, auto;
}

header[data-testid="stHeader"] { background: transparent; }
[data-testid="stDecoration"], [data-testid="stAppDeployButton"], footer { display: none !important; }

/* ---------- Sidebar ---------- */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ebdcfa 0%, #f6e1f3 55%, #ebdcfa 100%);
    border-right: 1.5px solid rgba(167,139,250,.18);
    border-radius: 0 34px 34px 0;
    box-shadow: 6px 0 30px rgba(124,77,189,.05);
}
[data-testid="stSidebar"] [data-testid="stSidebarContent"] { padding-top: 0; }

.dg-side-top { text-align: center; padding-top: .5rem; }
.dg-side-top svg, .dg-side-top img { width: 130px; max-width: 70%; height: auto; }
.dg-brand { font-size: 1.45rem; font-weight: 800; color: var(--ink); margin: .2rem 0 0; line-height: 1.1; }
.dg-tag { color: var(--purple-dark); font-weight: 600; font-size: .85rem; margin: .1rem 0 .8rem; }

.dg-nav { display: flex; flex-direction: column; gap: .35rem; margin-bottom: 1rem; }
.dg-nav a, .dg-nav summary {
    display: flex; align-items: center; gap: .8rem;
    padding: .5rem 1rem; font-size: .92rem; border-radius: 999px;
    color: var(--ink) !important; text-decoration: none !important;
    font-weight: 700; cursor: pointer; list-style: none;
    transition: background .15s ease;
}
.dg-nav summary::-webkit-details-marker { display: none; }
.dg-nav a:hover, .dg-nav summary:hover { background: rgba(243,182,228,.45); }
.dg-nav a.active { background: var(--pink); color: var(--purple-dark) !important; }
.dg-about p { margin: .3rem 1rem .6rem; font-size: .9rem; color: var(--purple-dark); }

.dg-side-foot { text-align: center; margin-top: 1.2rem; }
.dg-hand { font-family: 'Caveat', cursive; font-size: 1.2rem; color: var(--purple-dark); transform: rotate(-6deg); display: inline-block; }
.dg-side-foot svg, .dg-side-foot img { width: 105px; max-width: 60%; height: auto; margin-top: -.3rem; }

/* ---------- Main card ---------- */
.block-container, [data-testid="stMainBlockContainer"] {
    max-width: 960px;
    margin: 1rem auto 6rem;
    padding: 1.4rem 2rem 1.6rem !important;
    background: linear-gradient(160deg, rgba(248,240,255,.92), rgba(253,238,249,.92));
    backdrop-filter: blur(12px);
    border: 1.5px solid rgba(167,139,250,.3);
    border-radius: 28px;
    box-shadow: 0 14px 40px rgba(124,77,189,.12);
}

.dg-header { display: flex; align-items: flex-start; gap: .7rem; margin-bottom: .8rem; }
.dg-spark { font-size: 1.8rem; line-height: 1; color: var(--purple); }
.dg-header h1 { font-size: 1.6rem; font-weight: 800; color: var(--ink); margin: 0; padding: 0; line-height: 1.15; }
.dg-header p { margin: .1rem 0 0; font-size: .9rem; font-weight: 600; color: var(--ink); }

.dg-upload-head { text-align: center; margin: .6rem 0 .2rem; }
.dg-upload-head .dg-cloud { font-size: 1.8rem; color: var(--purple); line-height: 1; }
.dg-upload-head h3 { margin: .1rem 0 0; font-size: 1.15rem; font-weight: 800; color: var(--ink); }
.dg-upload-head p { margin: .1rem 0 .5rem; font-size: .85rem; color: var(--purple-dark); }

/* ---------- File uploader ---------- */
[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed #c4a6ee;
    border-radius: 20px;
    background: #f1e6fd;
    padding: 1.1rem;
}
[data-testid="stFileUploaderDropzone"] button {
    border-radius: 999px;
    border: 1.5px solid var(--lav-300);
    background: #fbf3ff;
    color: var(--purple-dark);
    font-weight: 700;
}
[data-testid="stFileUploaderFile"] {
    background: var(--lav-100);
    border-radius: 14px;
    padding: .2rem .5rem;
}

/* ---------- Buttons ---------- */
.stButton > button {
    border-radius: 999px;
    border: 1.5px solid var(--lav-300);
    background: linear-gradient(135deg, #f8c9ea, #d9bdf8);
    color: var(--purple-dark);
    font-weight: 800;
    padding: .4rem 1rem;
    font-size: .9rem;
    box-shadow: 0 4px 12px rgba(124,77,189,.08);
    transition: transform .12s ease, box-shadow .12s ease;
}
.stButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 16px rgba(124,77,189,.14);
    border-color: var(--purple);
    color: var(--purple-dark);
}
.stButton > button:focus-visible { outline: 3px solid var(--purple); outline-offset: 2px; }

/* ---------- Alerts / captions ---------- */
[data-testid="stAlert"] { border-radius: 18px; border: 1px solid var(--lav-200); }
[data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]) {
    background: var(--mint-bg); border: 1px solid var(--mint-border);
}
[data-testid="stAlertContentSuccess"], [data-testid="stAlertContentSuccess"] * { color: var(--mint-ink) !important; }
[data-testid="stCaptionContainer"] { color: var(--purple-dark); }

/* ---------- Empty state ---------- */
.dg-empty { text-align: center; margin: 1.6rem 0 1rem; position: relative; }
.dg-empty svg, .dg-empty img { width: 165px; max-width: 50%; height: auto; }
.dg-empty p { font-weight: 700; font-size: .9rem; margin: .2rem 0 0; color: var(--ink); }
.dg-bubble {
    position: absolute; top: -6px; left: calc(50% + 50px);
    background: #f8efff; border: 2px solid var(--lav-300); border-radius: 18px;
    padding: .1rem .7rem; color: var(--purple-dark); font-weight: 800; letter-spacing: 2px;
}

/* ---------- Chat ---------- */
[data-testid="stChatMessage"] {
    background: linear-gradient(135deg, #f6ecff, #fcecf7);
    border: 1px solid var(--lav-300);
    border-radius: 22px;
    padding: .8rem 1rem;
    font-size: .93rem;
    box-shadow: 0 4px 12px rgba(124,77,189,.05);
}
[data-testid="stChatMessageAvatarUser"] { background: #cfb5f2; }
[data-testid="stChatMessageAvatarAssistant"] { background: var(--pink-deep); }

[data-testid="stBottom"], [data-testid="stBottom"] > div,
[data-testid="stBottomBlockContainer"] { background: transparent !important; }

[data-testid="stChatInput"] {
    border-radius: 999px;
    border: 1.5px solid var(--lav-300);
    background: #f8efff;
    box-shadow: 0 6px 18px rgba(124,77,189,.08);
}
[data-testid="stChatInput"] > div, [data-testid="stChatInput"] textarea { background: transparent; border: none; }
[data-testid="stChatInputSubmitButton"] {
    background: linear-gradient(135deg, #b894ea, #9a72d6);
    color: #fff; border-radius: 50%;
}

/* ---------- Floating decorations ---------- */
.dg-note {
    position: fixed; top: 150px; right: 30px; z-index: 1; pointer-events: none;
    font-family: 'Caveat', cursive; font-size: 1.5rem; line-height: 1.1;
    color: var(--purple-dark); transform: rotate(-8deg); text-align: center;
}
.dg-books { position: fixed; bottom: 84px; right: 12px; z-index: 1; pointer-events: none; text-align: center; }
.dg-books svg, .dg-books img { width: 95px; height: auto; display: block; margin: 0 auto; }
.dg-stars span { position: fixed; z-index: 0; pointer-events: none; color: var(--purple); opacity: .4; }

@media (max-width: 1500px) {
    .dg-note, .dg-books { display: none; }
}
@media (max-width: 900px) {
    .block-container, [data-testid="stMainBlockContainer"] { padding: 1.2rem !important; border-radius: 24px; }
}
@media (prefers-reduced-motion: reduce) {
    * { transition: none !important; }
}
/* ---------- Richer colours (override: nothing is near-white) ---------- */
:root {
    --lav-100: #e9d9f9;
    --lav-200: #dcc8f5;
    --lav-300: #c3a5ec;
    --purple: #8a5fcc;
    --purple-dark: #4f2f8a;
    --ink: #2f1f52;
    --pink: #f5c2e8;
    --pink-deep: #ee9fdc;
    --mint-bg: #cfeedd;
    --mint-border: #9ed8ba;
    --mint-ink: #1f6445;
}
.stApp {
    background:
        radial-gradient(900px 500px at 15% 0%, rgba(205,176,245,.6), transparent 60%),
        radial-gradient(800px 500px at 100% 100%, rgba(240,178,222,.5), transparent 60%),
        linear-gradient(rgba(140,100,210,.07) 1px, transparent 1px),
        linear-gradient(90deg, rgba(140,100,210,.07) 1px, transparent 1px),
        linear-gradient(160deg, #e9dafa 0%, #e0cdf6 50%, #f3d9ee 100%) !important;
    background-size: auto, auto, 38px 38px, 38px 38px, auto !important;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ddc8f4 0%, #ecd0ea 55%, #ddc8f4 100%) !important;
    border-right: 1.5px solid rgba(140,100,210,.35);
}
.block-container, [data-testid="stMainBlockContainer"] {
    background: linear-gradient(160deg, rgba(234,218,251,.94), rgba(246,222,240,.94)) !important;
    border: 1.5px solid rgba(140,100,210,.35);
}
[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed #a47fdc;
    background: #dcc8f5 !important;
}
[data-testid="stFileUploaderDropzone"] button {
    background: #f1dcf6 !important;
    border: 1.5px solid #b08ae6;
}
[data-testid="stFileUploaderFile"] { background: #e4d2f8 !important; }
.stButton > button {
    background: linear-gradient(135deg, #f3b4e3, #c9a4f3) !important;
    border: 1.5px solid #b08ae6;
}
[data-testid="stChatMessage"] {
    background: linear-gradient(135deg, #dfcaf6, #ecd0ee) !important;
    border: 1px solid #c3a5ec;
}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {
    background: linear-gradient(135deg, #f5c6e8, #efb8e2) !important;
}
[data-testid="stChatMessageAvatarUser"] { background: #a67fe0; }
[data-testid="stChatInput"] {
    background: #dfcaf6 !important;
    border: 1.5px solid #a47fdc;
}
[data-testid="stChatInput"] textarea { color: var(--ink); }
[data-testid="stChatInput"] textarea::placeholder { color: #6e4fa3; opacity: 1; }
[data-testid="stChatInputSubmitButton"] { background: linear-gradient(135deg, #a67ade, #7d4fc2) !important; }
.dg-bubble { background: #e4d2f8; border-color: #b08ae6; }
.dg-stars span { opacity: .5; }

/* ---------- Remove dark boxes so text is always readable ---------- */
:root { color-scheme: light; }
[data-testid="stChatInput"] div,
[data-testid="stChatInput"] [data-baseweb],
[data-testid="stChatInput"] textarea,
[data-testid="stChatInputContainer"],
[data-testid="stChatInputContainer"] div {
    background: transparent !important;
    background-color: transparent !important;
    box-shadow: none !important;
    color: var(--ink) !important;
}
[data-testid="stChatInput"] textarea { caret-color: var(--purple-dark); }
[data-testid="stChatInput"] textarea::placeholder { color: #6e4fa3 !important; opacity: 1 !important; }
[data-testid="stBottom"], [data-testid="stBottom"] *:not(button):not(svg):not(path):not(textarea) {
    background-color: transparent !important;
}
[data-testid="stChatInput"] { background: #dfcaf6 !important; }
[data-testid="stChatInputSubmitButton"] { background: linear-gradient(135deg, #a67ade, #7d4fc2) !important; color: #fff !important; }
[data-testid="stChatMessage"], [data-testid="stChatMessage"] p,
[data-testid="stChatMessage"] li, [data-testid="stChatMessage"] span,
[data-testid="stChatMessage"] h1, [data-testid="stChatMessage"] h2,
[data-testid="stChatMessage"] h3, [data-testid="stChatMessage"] h4 { color: var(--ink); }
[data-testid="stChatMessage"] pre, [data-testid="stChatMessage"] code,
[data-testid="stCode"], [data-testid="stCode"] pre {
    background: #eadcfa !important;
    color: var(--purple-dark) !important;
}
[data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label { color: var(--ink); }

/* ---------- Chat input: no dark strip (strongest override) ---------- */
[data-testid="stChatInput"],
[data-testid="stChatInput"] *,
[data-testid="stChatInputTextArea"],
[data-baseweb="textarea"],
[data-baseweb="base-input"],
[data-testid="stBottom"] textarea {
    background: transparent !important;
    background-color: transparent !important;
    background-image: none !important;
    border-color: transparent !important;
    outline: none !important;
    -webkit-appearance: none;
    appearance: none;
}
[data-testid="stChatInput"] {
    background: #dfcaf6 !important;
    background-color: #dfcaf6 !important;
    border: 1.5px solid #a47fdc !important;
    border-radius: 999px !important;
}
/* paint lavender over anything dark inside the bar */
[data-testid="stChatInput"] div,
[data-testid="stChatInput"] textarea,
[data-testid="stChatInputTextArea"] {
    box-shadow: inset 0 0 0 1000px #dfcaf6 !important;
    -webkit-box-shadow: inset 0 0 0 1000px #dfcaf6 !important;
    color: #2f1f52 !important;
    -webkit-text-fill-color: #2f1f52 !important;
    border-radius: 999px;
}
[data-testid="stChatInput"] textarea::placeholder {
    color: #6e4fa3 !important;
    -webkit-text-fill-color: #6e4fa3 !important;
    opacity: 1 !important;
}
[data-testid="stChatInputSubmitButton"],
[data-testid="stChatInputSubmitButton"] * {
    background: linear-gradient(135deg, #a67ade, #7d4fc2) !important;
    box-shadow: none !important;
    -webkit-text-fill-color: #ffffff !important;
    color: #ffffff !important;
    border-radius: 50% !important;
}

"""


def bear_svg(width=150, headphones=True, laptop=False, sleeping=False):
    """Built-in cartoon bear used when no custom artwork is provided."""

    if sleeping:
        eyes = (
            '<path d="M62 108 q10 8 20 0" fill="none" stroke="#3b2a2a" stroke-width="3.5" stroke-linecap="round"/>'
            '<path d="M118 108 q10 8 20 0" fill="none" stroke="#3b2a2a" stroke-width="3.5" stroke-linecap="round"/>'
        )
    else:
        eyes = (
            '<circle cx="72" cy="108" r="6" fill="#3b2a2a"/><circle cx="128" cy="108" r="6" fill="#3b2a2a"/>'
            '<circle cx="74" cy="106" r="1.8" fill="#fff"/><circle cx="130" cy="106" r="1.8" fill="#fff"/>'
        )

    phones = ""
    if headphones:
        phones = (
            '<path d="M34 104 C30 28 170 28 166 104" fill="none" stroke="#b794e8" stroke-width="9" stroke-linecap="round"/>'
            '<rect x="19" y="86" width="25" height="46" rx="12" fill="#d6c2f5" stroke="#8b5cf6" stroke-width="3"/>'
            '<rect x="156" y="86" width="25" height="46" rx="12" fill="#d6c2f5" stroke="#8b5cf6" stroke-width="3"/>'
        )

    laptop_svg = ""
    if laptop:
        laptop_svg = (
            '<g transform="rotate(-6 150 178)">'
            '<rect x="104" y="152" width="88" height="52" rx="7" fill="#bfa9e6" stroke="#6b4fa0" stroke-width="3"/>'
            '<path d="M148 188 c-9 -7 -12 -13 -7 -17 c3 -3 7 -1 7 2 c0 -3 4 -5 7 -2 c5 4 2 10 -7 17z" fill="#fff"/>'
            '</g>'
        )

    return (
        f'<svg viewBox="0 0 200 215" width="{width}" xmlns="http://www.w3.org/2000/svg">'
        '<ellipse cx="100" cy="190" rx="56" ry="28" fill="#deae90" stroke="#8a5a44" stroke-width="3"/>'
        '<circle cx="48" cy="56" r="23" fill="#c98f6e" stroke="#8a5a44" stroke-width="3"/>'
        '<circle cx="152" cy="56" r="23" fill="#c98f6e" stroke="#8a5a44" stroke-width="3"/>'
        '<circle cx="48" cy="56" r="12" fill="#f0c4a8"/><circle cx="152" cy="56" r="12" fill="#f0c4a8"/>'
        '<ellipse cx="100" cy="108" rx="68" ry="60" fill="#deae90" stroke="#8a5a44" stroke-width="3"/>'
        '<ellipse cx="100" cy="126" rx="25" ry="19" fill="#f6dcc8"/>'
        f'{eyes}'
        '<ellipse cx="54" cy="125" rx="11" ry="7" fill="#f5a3b8" opacity=".65"/>'
        '<ellipse cx="146" cy="125" rx="11" ry="7" fill="#f5a3b8" opacity=".65"/>'
        '<ellipse cx="100" cy="119" rx="7.5" ry="5.5" fill="#5a3a34"/>'
        '<path d="M100 125 q-6 8 -13 3 M100 125 q6 8 13 3" fill="none" stroke="#5a3a34" stroke-width="3" stroke-linecap="round"/>'
        f'{phones}{laptop_svg}'
        '</svg>'
    )


def books_svg():
    return (
        '<svg viewBox="0 0 170 80" width="105" xmlns="http://www.w3.org/2000/svg">'
        '<g font-family="Caveat, cursive" font-size="16" fill="#ffffff">'
        '<rect x="20" y="4" width="130" height="23" rx="8" fill="#7c5cbf"/><text x="34" y="21">Dream</text>'
        '<rect x="8" y="29" width="144" height="23" rx="8" fill="#9570d6"/><text x="26" y="46">Code</text>'
        '<rect x="20" y="54" width="130" height="23" rx="8" fill="#7c5cbf"/><text x="38" y="71">Learn ♡</text>'
        '</g></svg>'
    )


def art(filename, fallback_svg):
    """Use custom artwork from assets/ if present, otherwise the built-in SVG."""

    path = ASSET_DIR / filename

    if path.exists():
        encoded = base64.b64encode(path.read_bytes()).decode()
        return f'<img src="data:image/png;base64,{encoded}" alt="">'

    return fallback_svg


def inject_theme():

    st.markdown(
        "<style>" + THEME_CSS + "</style>",
        unsafe_allow_html=True
    )

    stars = (
        '<div class="dg-stars">'
        '<span style="top:70px;left:46%;font-size:2rem;">✦</span>'
        '<span style="top:140px;left:30%;font-size:1.1rem;">✦</span>'
        '<span style="top:120px;right:90px;font-size:1.4rem;">✦</span>'
        '<span style="bottom:38%;right:70px;font-size:1.8rem;">✦</span>'
        '</div>'
    )

    note = '<div class="dg-note">Good<br>Ideas<br>Here! ♡</div>'

    books = (
        '<div class="dg-books">'
        + art("books_bear.png", bear_svg(90, sleeping=True) + books_svg())
        + '</div>'
    )

    st.markdown(stars + note + books, unsafe_allow_html=True)


inject_theme()

if not API_KEY:
    st.error("GEMINI_API not found. Check your .env file.")
    st.stop()

client = genai.Client(api_key=API_KEY)


# SESSION STATE

defaults = {
    "chunks": [],
    "index": None,
    "chat_history": [],
    "processed_files": []
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# Independent chat and dashboard workspaces
st.session_state.setdefault("doc_chat_sessions", [])
st.session_state.setdefault("doc_active_chat_id", None)
st.session_state.setdefault("dq_dashboard_sessions", [])
st.session_state.setdefault("dq_active_dashboard_id", None)


def new_session_id():
    return uuid.uuid4().hex[:10]


def save_document_chat():
    active_id = st.session_state.get("doc_active_chat_id")
    if not active_id:
        return
    for session in st.session_state.doc_chat_sessions:
        if session["id"] == active_id:
            session["history"] = list(st.session_state.chat_history)
            break


def load_document_chat(session_id):
    session = next((x for x in st.session_state.doc_chat_sessions if x["id"] == session_id), None)
    if session:
        st.session_state.doc_active_chat_id = session_id
        st.session_state.chat_history = list(session["history"])


def save_dashboard_session():
    active_id = st.session_state.get("dq_active_dashboard_id")
    if not active_id:
        return
    for session in st.session_state.dq_dashboard_sessions:
        if session["id"] == active_id:
            current_spec = st.session_state.get("dq_spec")
            session.update({
                "spec": current_spec,
                "history": list(st.session_state.get("dq_history", [])),
                "requests": list(st.session_state.get("dq_requests", [])),
                "insights": st.session_state.get("dq_insights", ""),
                "tables": st.session_state.get("dq_tables", {}),
            })
            if current_spec and session.get("title") in ("New dashboard", ""):
                session["title"] = current_spec.get("title", "My Dashboard")
            break


def load_dashboard_session(session_id):
    session = next((x for x in st.session_state.dq_dashboard_sessions if x["id"] == session_id), None)
    if session:
        st.session_state.dq_active_dashboard_id = session_id
        st.session_state.dq_spec = session["spec"]
        st.session_state.dq_history = list(session["history"])
        st.session_state.dq_requests = list(session["requests"])
        st.session_state.dq_insights = session["insights"]
        st.session_state.dq_tables = session.get("tables", {})


# FILE EXTRACTION

def extract_pdf(data, filename):

    reader = PdfReader(BytesIO(data))
    sections = []

    for page_number, page in enumerate(reader.pages, 1):

        text = page.extract_text() or ""

        if text.strip():
            sections.append(
                f"[Source: {filename} | Page {page_number}]\n{text}"
            )

    return sections


def extract_docx(data, filename):

    doc = Document(BytesIO(data))
    sections = []

    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            sections.append(paragraph.text)

    for table_number, table in enumerate(doc.tables, 1):

        rows = []

        for row in table.rows:
            rows.append(" | ".join(cell.text for cell in row.cells))

        sections.append(
            f"Table {table_number}:\n" + "\n".join(rows)
        )

    if not sections:
        raise ValueError("No readable text found.")

    return [
        f"[Source: {filename}]\n" + "\n".join(sections)
    ]


def extract_pptx(data, filename):

    presentation = Presentation(BytesIO(data))
    sections = []

    for slide_number, slide in enumerate(presentation.slides, 1):

        slide_text = []

        for shape in slide.shapes:

            if shape.has_text_frame:
                slide_text.append(shape.text)

            if shape.has_table:

                for row in shape.table.rows:
                    slide_text.append(
                        " | ".join(cell.text for cell in row.cells)
                    )

        if slide_text:
            sections.append(
                f"[Source: {filename} | Slide {slide_number}]\n"
                + "\n".join(slide_text)
            )

    if not sections:
        raise ValueError("No readable text found.")

    return sections


def extract_txt(data, filename):

    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        text = data.decode("latin-1")

    if not text.strip():
        raise ValueError("Text file is empty.")

    return [f"[Source: {filename}]\n{text}"]


def extract_csv(data, filename):

    try:
        df = pd.read_csv(BytesIO(data), encoding="utf-8")
    except UnicodeDecodeError:
        df = pd.read_csv(BytesIO(data), encoding="latin-1")

    if df.empty:
        raise ValueError("CSV file is empty.")

    return [
        f"[Source: {filename}]\n"
        f"Columns: {', '.join(map(str, df.columns))}\n"
        f"Data:\n{df.head(1000).to_string(index=False)}"
    ]


def extract_xlsx(data, filename):

    excel = pd.ExcelFile(BytesIO(data))
    sections = []

    for sheet_name in excel.sheet_names:

        df = pd.read_excel(
            excel,
            sheet_name=sheet_name
        )

        if not df.empty:

            sections.append(
                f"[Source: {filename} | Sheet: {sheet_name}]\n"
                f"Columns: {', '.join(map(str, df.columns))}\n"
                f"Data:\n{df.head(1000).to_string(index=False)}"
            )

    if not sections:
        raise ValueError("No readable data found in Excel file.")

    return sections


def extract_file(uploaded_file):

    filename = uploaded_file.name
    extension = filename.rsplit(".", 1)[-1].lower()
    data = uploaded_file.getvalue()

    if extension == "pdf":
        return extract_pdf(data, filename)

    elif extension == "docx":
        return extract_docx(data, filename)

    elif extension == "pptx":
        return extract_pptx(data, filename)

    elif extension == "txt":
        return extract_txt(data, filename)

    elif extension == "csv":
        return extract_csv(data, filename)

    elif extension == "xlsx":
        return extract_xlsx(data, filename)

    raise ValueError(f"Unsupported file type: {extension}")


# EMBEDDINGS

def create_embeddings(texts, task_type):

    embeddings = []

    for start in range(0, len(texts), 50):

        batch = texts[start:start + 50]

        response = client.models.embed_content(
            model=EMBED_MODEL,
            contents=batch,
            config=types.EmbedContentConfig(
                task_type=task_type
            )
        )

        embeddings.extend(
            [item.values for item in response.embeddings]
        )

    return np.array(embeddings, dtype="float32")


# PROCESS MULTIPLE FILES

def process_files(files):

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP
    )

    all_chunks = []
    processed_names = []

    for uploaded_file in files:

        sections = extract_file(uploaded_file)

        for section in sections:

            split_chunks = splitter.split_text(section)

            for chunk in split_chunks:
                all_chunks.append({
                    "text": chunk,
                    "source": uploaded_file.name
                })

        processed_names.append(uploaded_file.name)

    if not all_chunks:
        raise ValueError("No readable content found.")

    texts = [chunk["text"] for chunk in all_chunks]

    embeddings = create_embeddings(
        texts,
        "RETRIEVAL_DOCUMENT"
    )

    faiss.normalize_L2(embeddings)

    index = faiss.IndexFlatIP(embeddings.shape[1])
    index.add(embeddings)

    return all_chunks, index, processed_names


# DOCUMENT SEARCH

def search_documents(question):

    if st.session_state.index is None:
        return ""

    query_embedding = create_embeddings(
        [question],
        "RETRIEVAL_QUERY"
    )

    faiss.normalize_L2(query_embedding)

    scores, indices = st.session_state.index.search(
        query_embedding,
        min(TOP_K, len(st.session_state.chunks))
    )

    results = []

    for idx in indices[0]:

        if idx >= 0:
            results.append(
                st.session_state.chunks[idx]["text"]
            )

    return "\n\n".join(results)


# QUESTION CLASSIFICATION

def classify_question(question):

    prompt = f"""
Classify the user's question into exactly ONE category:

DOCUMENT: Asks about uploaded files.

GENERAL: Asks general knowledge or concepts.

HYBRID: Requires both uploaded file information
and general knowledge.

FOLLOWUP: Refers to previous conversation.

Return only one word:
DOCUMENT, GENERAL, HYBRID or FOLLOWUP.

Question: {question}
"""

    for model in CHAT_MODELS:

        try:

            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0,
                    max_output_tokens=20
                )
            )

            category = response.text.strip().upper()

            for option in [
                "DOCUMENT", "GENERAL", "HYBRID", "FOLLOWUP"
            ]:

                if option in category:
                    return option

        except Exception:
            continue

    return (
        "HYBRID"
        if st.session_state.index is not None
        else "GENERAL"
    )


# ANSWER GENERATION

def generate_answer(question, context, category):

    history = ""

    for message in st.session_state.chat_history[-8:]:

        history += (
            f"\n{message['role'].upper()}: "
            f"{message['content']}\n"
        )

    prompt = f"""
You are DocGenius, an intelligent AI document assistant.

QUESTION TYPE: {category}

RULES:

1. DOCUMENT questions:
Answer using uploaded documents.
Never invent document information.

2. GENERAL questions:
Use general knowledge.

3. HYBRID questions:
Combine document information and general knowledge.
Clearly distinguish the two.

4. FOLLOWUP questions:
Use previous conversation and document context.

5. Mention the filename and page, slide or sheet
when that information is available in the context.

6. If information is missing, clearly say so.

7. For Excel and CSV files:
Analyze the available data carefully.
Do not invent calculations or values.

8. Use simple language and structured answers.

PREVIOUS CONVERSATION:
{history}

RETRIEVED DOCUMENT CONTEXT:
{context if context else "No document context retrieved."}

USER QUESTION:
{question}

ANSWER:
"""

    last_error = None

    for model in CHAT_MODELS:

        for attempt in range(2):

            try:

                response = client.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=0.2,
                        max_output_tokens=2048
                    )
                )

                if response.text:
                    return response.text

                raise ValueError("Empty Gemini response.")

            except Exception as error:

                last_error = error
                error_message = str(error).lower()

                temporary_error = any(
                    term in error_message
                    for term in [
                        "503", "429", "500",
                        "service_unavailable",
                        "high demand",
                        "overloaded",
                        "unavailable",
                        "internal"
                    ]
                )

                if not temporary_error:
                    raise

                if attempt == 0:
                    time.sleep(2)

    raise RuntimeError(f"Gemini failed: {last_error}")




# ---------- DASHQUEST DATA ANALYTICS FUNCTIONS ----------

# ---------- DATA LOADING & CLEANING ----------

def kind_of(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return "date"
    return "number" if pd.api.types.is_numeric_dtype(s) else "text"


def clean_df(df):
    """Fix common messy-Excel problems: empty rows/cols, spaces, text numbers, text dates."""
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    strip = r"[,$₹€£%\s]"
    for c in df.columns:
        if kind_of(df[c]) != "text":
            continue
        s = df[c].dropna().astype(str)
        if s.empty:
            continue
        if pd.to_numeric(s.str.replace(strip, "", regex=True), errors="coerce").notna().mean() > 0.9:
            df[c] = pd.to_numeric(df[c].astype(str).str.replace(strip, "", regex=True), errors="coerce")
            continue
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            if pd.to_datetime(s.head(50), errors="coerce").notna().mean() > 0.9:
                df[c] = pd.to_datetime(df[c], errors="coerce")
    return df.reset_index(drop=True)


def load_files(files):
    tables = {}
    for f in files:
        stem = f.name.rsplit(".", 1)[0]
        data = f.getvalue()
        if f.name.lower().endswith(".csv"):
            try:
                raw = {stem: pd.read_csv(io.BytesIO(data), encoding="utf-8")}
            except UnicodeDecodeError:
                raw = {stem: pd.read_csv(io.BytesIO(data), encoding="latin-1")}
        else:
            xl = pd.ExcelFile(io.BytesIO(data))
            raw = {f"{stem}_{sh}": xl.parse(sh) for sh in xl.sheet_names}
        for name, df in raw.items():
            df = clean_df(df)
            if not df.empty:
                tables[name] = df
    if not tables:
        raise ValueError("No readable data found.")
    return tables


def prepared(tables, dedupe, fill):
    out = {}
    for name, df in tables.items():
        d = df.drop_duplicates() if dedupe else df
        if fill:
            d = d.copy()
            nums = d.select_dtypes("number").columns
            d[nums] = d[nums].fillna(d[nums].median())
        out[name] = d
    return out


def profile(tables):
    out = {}
    for name, df in tables.items():
        cols = []
        for c in df.columns:
            s, k = df[c], kind_of(df[c])
            info = {"name": c, "type": k}
            if k == "number":
                info["min"], info["max"] = float(s.min()), float(s.max())
            elif k == "date":
                info["min"], info["max"] = str(s.min()), str(s.max())
            else:
                info["unique"] = int(s.nunique())
                info["examples"] = [str(v) for v in s.dropna().unique()[:4]]
            cols.append(info)
        out[name] = {"rows": len(df), "columns": cols}
    return json.dumps(out, default=str)


# ---------- AI: PLAIN ENGLISH -> DASHBOARD SPEC ----------

SPEC_PROMPT = """You are a BI dashboard designer. Return ONLY valid JSON in this shape:
{"title": str, "widgets": [{"type": "kpi|bar|line|pie|scatter|hist|table", "title": str, "table": str,
"x": column|null, "y": column|null, "agg": "sum|mean|count|min|max|nunique", "color": column|null,
"top_n": int|null, "sort": "asc|desc|null", "date_grain": "day|month|year|null", "columns": [column]|null}]}
Rules: use ONLY tables and columns from DATA PROFILE, spelled exactly. kpi uses y + agg (no x).
bar/line/pie need x (category or date) plus y + agg (use agg "count" when no y). scatter needs numeric x and y.
hist needs numeric x. table uses "columns". Use date_grain for line charts over dates. Use 4-8 widgets, KPIs first.
If CURRENT DASHBOARD is not "none", apply the request to it and return the FULL updated dashboard.
If the request cannot be answered from the data, return {"title": "", "widgets": [], "message": "<why>"}.
DATA PROFILE: <<PROFILE>>
CURRENT DASHBOARD: <<CURRENT>>
USER REQUEST: <<REQUEST>>"""


def ask_ai(prompt, json_mode=False):
    last = None
    for model in CHAT_MODELS:
        for _ in range(2):
            try:
                r = client.models.generate_content(
                    model=model, contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=0.2, max_output_tokens=4096,
                        response_mime_type="application/json" if json_mode else None))
                if r.text:
                    return r.text
            except Exception as e:
                last = e
                time.sleep(1.5)
    raise RuntimeError(f"Gemini failed: {last}")


def validate(spec, tables):
    good, warn = [], []
    for w in spec.get("widgets", []):
        title = w.get("title", "widget")
        t = w.get("table") or next(iter(tables))
        if t not in tables or w.get("type") not in KINDS:
            warn.append(f"Skipped '{title}': unknown table or chart type.")
            continue
        cols = set(tables[t].columns)
        bad = [k for k in ("x", "y", "color") if w.get(k) and w[k] not in cols]
        bad += [c for c in (w.get("columns") or []) if c not in cols]
        if bad or any(not w.get(k) for k in NEEDS.get(w["type"], [])):
            warn.append(f"Skipped '{title}': missing or unknown columns {bad}.")
            continue
        if w.get("agg") not in AGGS or (not w.get("y") and w.get("agg") != "nunique"):
            w["agg"] = "count" if not w.get("y") else "sum"
        w["table"] = t
        good.append(w)
    return {"title": spec.get("title") or "My Dashboard", "widgets": good}, warn


def run_request(text, tables):
    cur = st.session_state.dq_spec
    prompt = (SPEC_PROMPT.replace("<<PROFILE>>", profile(tables))
              .replace("<<CURRENT>>", json.dumps(cur) if cur else "none")
              .replace("<<REQUEST>>", text))
    raw = ask_ai(prompt, json_mode=True).strip().removeprefix("```json").removeprefix("```").removesuffix("```")
    data = json.loads(raw)
    spec, warn = validate(data, tables)
    if not spec["widgets"]:
        st.warning(data.get("message") or "I couldn't build that from this data. Try rephrasing.")
        return
    if cur:
        st.session_state.dq_history.append(cur)
    st.session_state.dq_spec, st.session_state.dq_insights = spec, ""
    st.session_state.dq_requests.append(text)
    for w in warn:
        st.warning(w)


# ---------- COMPUTE & DRAW ----------

def aggregate(df, keys, y, agg):
    if agg == "count" or not y or y in keys:
        return df.groupby(keys, dropna=False).size().reset_index(name="count"), "count"
    return df.groupby(keys, dropna=False)[y].agg(agg).reset_index(), y


def compute(w, df):
    kind, x, y, agg, color = w["type"], w.get("x"), w.get("y"), w.get("agg", "sum"), w.get("color")
    if kind == "kpi":
        if agg == "count" or not y:
            return {"value": len(df) if not y else int(df[y].count())}
        return {"value": df[y].agg(agg)}
    if kind == "table":
        return {"df": df[w.get("columns") or list(df.columns)].head(200)}
    if kind in ("hist", "scatter"):
        keep = [c for c in dict.fromkeys([x, y, color]) if c]
        return {"df": df[keep].dropna().head(5000)}
    grain = w.get("date_grain")
    if grain in ("day", "month", "year") and kind_of(df[x]) == "date":
        df = df.copy()
        df[x] = df[x].dt.to_period({"day": "D", "month": "M", "year": "Y"}[grain]).dt.to_timestamp()
    g, ycol = aggregate(df, [x] + ([color] if color and color != x else []), y, agg)
    if kind == "line":
        g = g.sort_values(x)
    else:
        g = g.sort_values(ycol, ascending=w.get("sort") == "asc")
        if w.get("top_n"):
            g = g.head(int(w["top_n"]))
    return {"df": g, "ycol": ycol}


def make_fig(w, res):
    kind, x, color, title = w["type"], w.get("x"), w.get("color"), w.get("title", "")
    d = res["df"]
    if kind == "bar":
        fig = px.bar(d, x=x, y=res["ycol"], color=color, title=title, barmode="group")
    elif kind == "line":
        fig = px.line(d, x=x, y=res["ycol"], color=color, title=title, markers=True)
    elif kind == "pie":
        fig = px.pie(d, names=x, values=res["ycol"], title=title, hole=0.4)
    elif kind == "scatter":
        fig = px.scatter(d, x=x, y=w["y"], color=color, title=title)
    else:
        fig = px.histogram(d, x=x, color=color, title=title)
    fig.update_layout(template="plotly_white", colorway=PALETTE, paper_bgcolor="rgba(0,0,0,0)",
                      plot_bgcolor="rgba(255,255,255,.6)", margin=dict(l=10, r=10, t=50, b=10),
                      font=dict(color="#2f1f52"))
    return fig


def fmt(v):
    try:
        return f"{v:,.2f}" if isinstance(v, float) and not float(v).is_integer() else f"{int(v):,}"
    except Exception:
        return str(v)


def apply_filters(df, f):
    for c, v in f.items():
        if c not in df.columns:
            continue
        if isinstance(v, tuple):
            df = df[(df[c] >= pd.Timestamp(v[0])) & (df[c] < pd.Timestamp(v[1]) + pd.Timedelta(days=1))]
        elif v:
            df = df[df[c].astype(str).isin(v)]
    return df


def render_dashboard(spec, tables):
    results = []
    kpis = [w for w in spec["widgets"] if w["type"] == "kpi"]
    others = [w for w in spec["widgets"] if w["type"] != "kpi"]

    for i in range(0, len(kpis), 4):
        row = kpis[i:i + 4]
        for col, w in zip(st.columns(len(row)), row):
            try:
                res = compute(w, tables[w["table"]])
                col.metric(f"{w.get('title', '')} ({w.get('agg', 'sum')})", fmt(res["value"]))
                results.append((w, res))
            except Exception as e:
                col.warning(f"{w.get('title')}: {e}")

    for i in range(0, len(others), 2):
        row = others[i:i + 2]
        for n, (col, w) in enumerate(zip(st.columns(2), row)):
            with col.container(border=True):
                try:
                    res = compute(w, tables[w["table"]])
                    if w["type"] == "table":
                        st.markdown(f"**{w.get('title', '')}**")
                        st.dataframe(res["df"], use_container_width=True)
                    else:
                        res["fig"] = make_fig(w, res)
                        st.plotly_chart(res["fig"], use_container_width=True, key=f"w{i}_{n}")
                    with st.expander("How was this calculated?"):
                        st.json(w)
                        st.dataframe(res["df"].head(20), use_container_width=True)
                    results.append((w, res))
                except Exception as e:
                    st.warning(f"{w.get('title')}: could not build this widget ({e}).")
    return results


def build_html(spec, results):
    parts = [f"<h1>{html.escape(spec['title'])}</h1>"]
    first = True
    for w, res in results:
        t = html.escape(w.get("title", ""))
        if "value" in res:
            parts.append(f"<div class='k'><small>{t}</small><b>{fmt(res['value'])}</b></div>")
        elif "fig" in res:
            parts.append("<div class='c'>" + res["fig"].to_html(
                full_html=False, include_plotlyjs="cdn" if first else False) + "</div>")
            first = False
        else:
            parts.append(f"<div class='c'><h3>{t}</h3>{res['df'].to_html(index=False)}</div>")
    style = ("body{font-family:sans-serif;background:#efe3fb;color:#2f1f52;padding:24px}"
             ".k{display:inline-block;background:#dfcaf6;border-radius:14px;padding:12px 18px;margin:6px}"
             ".k b{display:block;font-size:1.6rem}.c{background:#f3e9fd;border-radius:16px;padding:12px;margin:12px 0}")
    return f"<html><head><meta charset='utf-8'><style>{style}</style></head><body>{''.join(parts)}</body></html>"


def make_insights(results):
    blocks = []
    for w, res in results:
        if "value" in res:
            blocks.append(f"KPI {w.get('title')} ({w.get('agg')}): {res['value']}")
        elif w["type"] in ("bar", "line", "pie"):
            blocks.append(f"{w.get('title')}\n{res['df'].head(20).to_csv(index=False)}")
    prompt = ("You are a data analyst. Using ONLY the numbers below, write 4-6 short bullet insights "
              "(trends, top/bottom items, outliers, comparisons). Do not invent numbers.\n\n" + "\n\n".join(blocks))
    return ask_ai(prompt)




# ---------- DOCUMENT ASSISTANT WORKSPACE ----------

def render_document_workspace():
    # MAIN HEADER (UI)

    st.markdown(
        '<div id="dg-top" class="dg-header">'
        '<div class="dg-spark">✦</div>'
        '<div><h1>DocQuest</h1>'
        '<p>Upload your documents and start asking questions! ♡</p></div>'
        '</div>',
        unsafe_allow_html=True
    )


    # UPLOAD AREA (moved from sidebar into the main card; same logic as before)

    st.markdown(
        '<div id="dg-upload" class="dg-upload-head">'
        '<div class="dg-cloud">☁</div>'
        '<h3>Upload Your Documents</h3>'
        '<p>Drag &amp; drop your files here or click to browse</p>'
        '</div>',
        unsafe_allow_html=True
    )

    uploaded_files = st.file_uploader(
        "Choose your files",
        type=SUPPORTED_FILES,
        accept_multiple_files=True,
        label_visibility="collapsed",
        key="dq_document_uploader"
    )

    st.caption(
        "Supported: PDF, DOCX, PPTX, TXT, CSV, XLSX"
    )

    if uploaded_files:

        if st.button(
            "Process Documents",
            use_container_width=True
        ):

            with st.spinner("Reading and processing files..."):

                try:

                    chunks, index, names = process_files(
                        uploaded_files
                    )

                    st.session_state.chunks = chunks
                    st.session_state.index = index
                    st.session_state.processed_files = names
                    st.session_state.chat_history = []
                    save_document_chat()

                    st.success("Documents processed successfully!")

                except Exception as error:

                    st.error(f"Processing failed: {error}")


    # SIDEBAR

    with st.sidebar:

        st.markdown("### 💬 Document chats")
        if st.button("＋ New chat", use_container_width=True, key="doc_new_chat"):
            save_document_chat()
            new_id = new_session_id()
            st.session_state.doc_chat_sessions.append({
                "id": new_id, "title": "New chat", "history": []
            })
            st.session_state.doc_active_chat_id = new_id
            st.session_state.chat_history = []
            st.session_state["doc_chat_picker"] = new_id
            st.rerun()

        if not st.session_state.doc_chat_sessions:
            first_id = new_session_id()
            st.session_state.doc_chat_sessions.append({
                "id": first_id, "title": "New chat", "history": []
            })
            st.session_state.doc_active_chat_id = first_id

        chat_options = {x["id"]: x["title"] for x in st.session_state.doc_chat_sessions}
        active_chat = st.session_state.doc_active_chat_id
        if active_chat not in chat_options:
            active_chat = st.session_state.doc_chat_sessions[0]["id"]
            load_document_chat(active_chat)
        selected_chat = st.selectbox(
            "Your chats",
            options=list(chat_options),
            format_func=lambda value: chat_options[value],
            index=list(chat_options).index(active_chat),
            key="doc_chat_picker"
        )
        if selected_chat != st.session_state.doc_active_chat_id:
            save_document_chat()
            load_document_chat(selected_chat)
            st.rerun()

        st.divider()
        st.markdown(
            '<div class="dg-side-top">'
            + art("sidebar_bear.png", bear_svg(130, laptop=True))
            + '<div class="dg-brand">DocQuest</div>'
            '<div class="dg-tag">Ask • Learn • Explore ♡</div>'
            '</div>'
            '<div class="dg-nav">'
            '<a class="active" href="#dg-top">🏠&nbsp; Home</a>'
            '<a href="#dg-upload">☁️&nbsp; Upload</a>'
            '<a href="#dg-chat">💬&nbsp; Chat</a>'
            '<details class="dg-about"><summary>ℹ️&nbsp; About</summary>'
            '<p>DocQuest reads your PDF, Word, PowerPoint, text, CSV and Excel '
            'files and answers your questions about them.</p></details>'
            '</div>',
            unsafe_allow_html=True
        )

        if st.session_state.index is not None:

            st.success("Documents Ready")

            st.write("**Uploaded files:**")

            for filename in st.session_state.processed_files:
                st.write(f"📄 {filename}")

            st.write(
                f"Total chunks: {len(st.session_state.chunks)}"
            )

            if st.button("Clear Current Chat", key="doc_clear_current_chat"):
                st.session_state.chat_history = []
                save_document_chat()
                st.rerun()

        if st.button("Remove All Documents"):

            st.session_state.chunks = []
            st.session_state.index = None
            st.session_state.processed_files = []
            st.session_state.chat_history = []
            save_document_chat()

            st.rerun()

        st.markdown(
            '<div class="dg-side-foot">'
            '<div class="dg-hand">Stay Curious<br>Stay Smart ♡</div>'
            + art("footer_bear.png", bear_svg(105, headphones=False))
            + '</div>',
            unsafe_allow_html=True
        )


    # MAIN INTERFACE

    if st.session_state.index is None:

        st.markdown(
            '<div class="dg-empty">'
            '<div class="dg-bubble">•••</div>'
            + art("sleeping_bear.png", bear_svg(165, laptop=True, sleeping=True))
            + '<p>Upload a document and ask your first question! ♡</p>'
            '</div>',
            unsafe_allow_html=True
        )

    else:

        st.caption(
            f"Active documents: "
            f"{len(st.session_state.processed_files)}"
        )


    # CHAT HISTORY

    st.markdown('<div id="dg-chat"></div>', unsafe_allow_html=True)

    for message in st.session_state.chat_history:

        with st.chat_message(message["role"]):
            st.markdown(message["content"])


    # CHAT INPUT

    question = st.chat_input(
        "Ask anything about your documents..."
    )

    if question:

        st.session_state.chat_history.append({
            "role": "user",
            "content": question
        })
        active_chat = next(
            (x for x in st.session_state.doc_chat_sessions
             if x["id"] == st.session_state.doc_active_chat_id), None
        )
        if active_chat and active_chat["title"] == "New chat":
            active_chat["title"] = question[:36] + ("…" if len(question) > 36 else "")
        save_document_chat()

        with st.chat_message("user"):
            st.markdown(question)

        with st.chat_message("assistant"):

            with st.spinner("Analyzing your question..."):

                try:

                    category = classify_question(question)

                    if category in [
                        "DOCUMENT", "HYBRID", "FOLLOWUP"
                    ]:
                        context = search_documents(question)
                    else:
                        context = ""

                    answer = generate_answer(
                        question,
                        context,
                        category
                    )

                    st.markdown(answer)

                    st.session_state.chat_history.append({
                        "role": "assistant",
                        "content": answer
                    })
                    save_document_chat()

                except Exception as error:

                    st.error(f"Could not generate answer: {error}")

                    st.session_state.chat_history.pop()


# ---------- DATA ANALYTICS WORKSPACE ----------

def render_analytics_workspace():
    for key, value in {
        "dq_tables": {},
        "dq_spec": None,
        "dq_history": [],
        "dq_requests": [],
        "dq_insights": "",
    }.items():
        st.session_state.setdefault(key, value)

    if not st.session_state.dq_dashboard_sessions:
        first_id = new_session_id()
        st.session_state.dq_dashboard_sessions.append({
            "id": first_id, "title": "New dashboard", "spec": None,
            "history": [], "requests": [], "insights": "",
            "tables": dict(st.session_state.get("dq_tables", {}))
        })
        st.session_state.dq_active_dashboard_id = first_id

    # ---------- SIDEBAR ----------

    with st.sidebar:
        st.markdown("### 🧸 DocQuest · Data Analytics")
        if st.button("＋ New dashboard", use_container_width=True, key="dq_new_dashboard"):
            save_dashboard_session()
            new_id = new_session_id()
            st.session_state.dq_dashboard_sessions.append({
                "id": new_id, "title": "New dashboard", "spec": None,
                "history": [], "requests": [], "insights": "",
                "tables": dict(st.session_state.get("dq_tables", {}))
            })
            st.session_state.dq_active_dashboard_id = new_id
            st.session_state.dq_spec = None
            st.session_state.dq_history = []
            st.session_state.dq_requests = []
            st.session_state.dq_insights = ""
            st.session_state["dq_dashboard_picker"] = new_id
            st.rerun()

        dashboard_options = {x["id"]: x["title"] for x in st.session_state.dq_dashboard_sessions}
        active_dashboard = st.session_state.dq_active_dashboard_id
        if active_dashboard not in dashboard_options:
            active_dashboard = st.session_state.dq_dashboard_sessions[0]["id"]
            load_dashboard_session(active_dashboard)
        selected_dashboard = st.selectbox(
            "My dashboards",
            options=list(dashboard_options),
            format_func=lambda value: dashboard_options[value],
            index=list(dashboard_options).index(active_dashboard),
            key="dq_dashboard_picker"
        )
        if selected_dashboard != st.session_state.dq_active_dashboard_id:
            save_dashboard_session()
            load_dashboard_session(selected_dashboard)
            st.rerun()

        files = st.file_uploader("Upload data", type=["csv", "xlsx"], accept_multiple_files=True, key="dq_data_uploader")
        if files and st.button("Load data", use_container_width=True):
            try:
                with st.spinner("Reading and cleaning your data..."):
                    st.session_state.dq_tables = load_files(files)
                st.session_state.dq_spec, st.session_state.dq_history = None, []
                st.session_state.dq_requests, st.session_state.dq_insights = [], ""
                current_dashboard = next(
                    (x for x in st.session_state.dq_dashboard_sessions
                     if x["id"] == st.session_state.dq_active_dashboard_id), None
                )
                if current_dashboard:
                    current_dashboard.update({
                        "title": ", ".join(st.session_state.dq_tables.keys())[:40] or "New dashboard",
                        "spec": None, "history": [], "requests": [], "insights": "",
                        "tables": dict(st.session_state.dq_tables)
                    })
                save_dashboard_session()
                st.success("Data loaded!")
            except Exception as e:
                st.error(f"Could not read files: {e}")

        base = st.session_state.dq_tables
        dedupe = fill = False
        filters = {}
        if base:
            st.markdown("**Data tools**")
            dedupe = st.checkbox("Remove duplicate rows")
            fill = st.checkbox("Fill missing numbers with median")
            tables = prepared(base, dedupe, fill)

            st.markdown("**Filters** (apply to every chart)")
            cols = {}
            for df in tables.values():
                for c in df.columns:
                    cols.setdefault(c, df[c])
            cands = [c for c, s in cols.items() if kind_of(s) == "date" or (kind_of(s) == "text" and s.nunique() <= 50)]
            for c in st.multiselect("Filter by", cands):
                s = pd.concat([df[c] for df in tables.values() if c in df.columns])
                if kind_of(s) == "date":
                    r = st.date_input(c, (s.min().date(), s.max().date()))
                    if len(r) == 2:
                        filters[c] = r
                else:
                    filters[c] = st.multiselect(c, sorted(s.dropna().astype(str).unique()))

            st.markdown("**Save / load dashboard**")
            if st.session_state.dq_spec:
                st.download_button("Download dashboard file", json.dumps(st.session_state.dq_spec, indent=2),
                                   "dashboard.json", "application/json", use_container_width=True)
            spec_file = st.file_uploader("Load dashboard file", type=["json"], key="dq_json_uploader")
            if spec_file and st.button("Open dashboard file", use_container_width=True):
                try:
                    spec, warn = validate(json.loads(spec_file.getvalue()), tables)
                    st.session_state.dq_spec = spec
                    save_dashboard_session()
                    for w in warn:
                        st.warning(w)
                except Exception as e:
                    st.error(f"Invalid dashboard file: {e}")

            if st.session_state.dq_requests:
                with st.expander("Your requests"):
                    for r in st.session_state.dq_requests:
                        st.write(f"• {r}")


    # ---------- MAIN ----------

    st.markdown('<div class="dq-head"><h1>📊 Data Analytics</h1>'
                '<p>Explore your CSV and Excel data with AI-powered dashboards ♡</p></div>',
                unsafe_allow_html=True)

    if not st.session_state.dq_tables:
        st.info("Upload one or more Excel or CSV files in the sidebar, then click **Load data**. "
                "Then try: *monthly revenue by region*, *top 5 products by sales* or *show average price as a KPI*.")
        return

    working = {n: apply_filters(df, filters) for n, df in tables.items()}

    c1, c2, c3 = st.columns(3)
    suggest = c1.button("✨ Suggest a starter dashboard", use_container_width=True)
    undo = c2.button("↩ Undo last change", use_container_width=True, disabled=not st.session_state.dq_history)
    reset = c3.button("🗑 Clear dashboard", use_container_width=True, disabled=not st.session_state.dq_spec)

    if undo:
        st.session_state.dq_spec = st.session_state.dq_history.pop()
        save_dashboard_session()
    if reset:
        st.session_state.dq_spec, st.session_state.dq_history, st.session_state.dq_insights = None, [], ""
        save_dashboard_session()

    request = st.chat_input("Ask for a chart, KPI or change, e.g. 'monthly sales by region as a line chart'")
    if suggest:
        request = "Create a useful starter dashboard that gives an overview of this data."

    if request:
        with st.spinner("Designing your dashboard..."):
            try:
                run_request(request, tables)
                current_dashboard = next(
                    (x for x in st.session_state.dq_dashboard_sessions
                     if x["id"] == st.session_state.dq_active_dashboard_id), None
                )
                if current_dashboard and current_dashboard["title"] == "New dashboard":
                    current_dashboard["title"] = st.session_state.dq_spec.get("title", "My Dashboard")
                save_dashboard_session()
            except Exception as e:
                st.error(f"Could not build that: {e}")

    tab_dash, tab_data = st.tabs(["📊 Dashboard", "🔍 Data & profile"])

    with tab_dash:
        spec = st.session_state.dq_spec
        if not spec:
            st.markdown("#### No dashboard yet\nClick **Suggest a starter dashboard** or type a request below ♡")
        else:
            st.subheader(spec["title"])
            results = render_dashboard(spec, working)

            a, b = st.columns(2)
            if a.button("💡 Explain this dashboard (AI insights)", use_container_width=True):
                with st.spinner("Analysing the numbers..."):
                    try:
                        st.session_state.dq_insights = make_insights(results)
                        save_dashboard_session()
                    except Exception as e:
                        st.error(f"Could not generate insights: {e}")
            b.download_button("⬇ Download as HTML", build_html(spec, results), "dashboard.html",
                              "text/html", use_container_width=True)
            if st.session_state.dq_insights:
                with st.container(border=True):
                    st.markdown("**Insights**")
                    st.markdown(st.session_state.dq_insights)

    with tab_data:
        for name, df in working.items():
            with st.expander(f"{name}  ({len(df):,} rows × {len(df.columns)} columns)", expanded=len(working) == 1):
                st.dataframe(df.head(50), use_container_width=True)
                info = pd.DataFrame({"type": [kind_of(df[c]) for c in df.columns],
                                     "missing": df.isna().sum().values,
                                     "unique": df.nunique().values}, index=df.columns)
                st.dataframe(info, use_container_width=True)
                st.caption(f"Duplicate rows: {int(df.duplicated().sum()):,}")
                st.download_button("Download filtered data (CSV)", df.to_csv(index=False), f"{name}.csv",
                                   "text/csv", key=f"dl_{name}")


# ---------- WORKSPACE NAVIGATION ----------

with st.sidebar:
    st.markdown(
        '<div class="dg-side-top">'
        + art("sidebar_bear.png", bear_svg(115, laptop=True))
        + '<div class="dg-brand">DocQuest</div>'
        '<div class="dg-tag">Ask • Analyze • Explore ♡</div>'
        '</div>',
        unsafe_allow_html=True
    )
    workspace = st.radio(
        "Choose workspace",
        ["Document Assistant", "Data Analytics"],
        key="dq_workspace"
    )
    st.divider()

if workspace == "Document Assistant":
    render_document_workspace()
else:
    render_analytics_workspace()
