import base64
import binascii
import hashlib
import html
import io
import json
from pathlib import Path
import sys
import altair as alt
import pandas as pd
import streamlit as st
from analytics_calendar import (
    build_calendar_frame,
    build_weekly_summary,
    classify_polish_day,
    get_last_completed_reference_week,
)
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from release_loader import compare_releases as compare_release_frames
from release_loader import load_release as load_release_file


THRESHOLD = 15
MAX_MATRIX_STYLE_CELLS = 50000
BRAND_NAME = "Pjoter Development"
BASE_DIR = Path(__file__).resolve().parent
RUNTIME_ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else BASE_DIR


def resolve_runtime_path(relative_path):
    try:
        external_path = RUNTIME_ROOT / relative_path
        internal_path = BASE_DIR / relative_path
        return external_path if external_path.exists() else internal_path
    except Exception:
        return BASE_DIR / relative_path


LOGO_PATH = resolve_runtime_path(Path("assets") / "logo.png")
AUTH_USERS_PATH = resolve_runtime_path(Path("config") / "users.json")
DATE_OPTIONS = ["Receipt Date", "Ship Date"]
DATE_LABELS = {
    "Receipt Date": "Data odbioru",
    "Ship Date": "Data wysyłki",
}
CHANGE_DIRECTION_LABELS = {
    "Increase": "Wzrost",
    "Decrease": "Spadek",
    "No Change": "Bez zmian",
}
MATRIX_METRIC_LABELS = {
    "Current Quantity": "Aktualna ilość",
    "Previous Quantity": "Poprzednia ilość",
    "Delta": "Zmiana ilości",
    "Percent Change": "Zmiana procentowa",
}
VIEW_MODE_LABELS = {
    "chart": "Wykres",
    "table": "Dane",
}


st.set_page_config(
    page_title="Pjoter Development | Analiza zamówień i wysyłek",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.markdown(
    """
    <style>
    :root {
        --ink: #f5f7fb;
        --navy: #e8eef9;
        --slate: #b5c0d4;
        --muted: #8190a8;
        --line: rgba(148, 163, 184, 0.16);
        --line-strong: rgba(148, 163, 184, 0.28);
        --panel: rgba(14, 19, 29, 0.90);
        --panel-strong: rgba(16, 22, 34, 0.97);
        --panel-soft: rgba(12, 17, 26, 0.82);
        --mint: #43d3b1;
        --rose: #ff6f6f;
        --steel: #7dd3fc;
        --steel-strong: #4cc9f0;
        --accent: #9dd7ff;
        --canvas-a: #03060b;
        --canvas-b: #08111a;
        --canvas-c: #101927;
        --shadow-xl: 0 38px 120px rgba(0, 0, 0, 0.42);
        --shadow-lg: 0 26px 70px rgba(0, 0, 0, 0.34);
        --shadow-md: 0 18px 44px rgba(0, 0, 0, 0.24);
    }
    html, body, [class*="css"]  {
        font-family: "Segoe UI", "Aptos", "Helvetica Neue", Arial, sans-serif;
    }
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(76, 201, 240, 0.13), transparent 22%),
            radial-gradient(circle at top right, rgba(125, 211, 252, 0.08), transparent 20%),
            radial-gradient(circle at bottom right, rgba(67, 211, 177, 0.06), transparent 24%),
            linear-gradient(180deg, var(--canvas-a) 0%, var(--canvas-b) 52%, var(--canvas-c) 100%);
        color: var(--ink);
    }
    .block-container {
        padding-top: 1.1rem;
        padding-bottom: 2.6rem;
        max-width: 1540px;
    }
    section[data-testid="stSidebar"] {
        display: none !important;
    }
    [data-testid="collapsedControl"],
    [data-testid="stSidebarCollapseButton"],
    button[aria-label="Close sidebar"],
    button[aria-label="Open sidebar"] {
        display: none !important;
    }
    .filter-panel-shell {
        border: 1px solid var(--line);
        border-radius: 28px;
        padding: 1rem 1rem 1.1rem 1rem;
        background:
            linear-gradient(180deg, rgba(6,10,16,0.98), rgba(10,15,24,0.96)),
            radial-gradient(circle at top, rgba(76, 201, 240, 0.08), transparent 26%);
        backdrop-filter: blur(18px);
        box-shadow: var(--shadow-lg);
        position: sticky;
        top: 1rem;
    }
    .file-type-banner {
        width: 100%;
        border-radius: 22px;
        border: 1px solid rgba(255, 255, 255, 0.09);
        display: flex;
        align-items: center;
        justify-content: center;
        text-align: center;
        overflow: hidden;
        position: relative;
        box-shadow: var(--shadow-md);
        background:
            linear-gradient(135deg, rgba(255,255,255,0.06), rgba(255,255,255,0.01)),
            linear-gradient(180deg, rgba(20, 27, 39, 0.96), rgba(11, 16, 24, 0.96));
    }
    .file-type-banner::before {
        content: "";
        position: absolute;
        inset: 0;
        background:
            linear-gradient(90deg, rgba(255,255,255,0.06), transparent 26%, transparent 74%, rgba(255,255,255,0.04)),
            radial-gradient(circle at top left, rgba(255,255,255,0.10), transparent 34%);
        pointer-events: none;
        opacity: 0.85;
    }
    .file-type-banner--sidebar {
        min-height: 108px;
        margin: 0 0 0.95rem 0;
        padding: 0.9rem 1rem;
    }
    .file-type-banner--header {
        min-height: 118px;
        padding: 1rem 1.15rem;
    }
    .file-type-banner--tesla {
        background:
            linear-gradient(135deg, rgba(193, 199, 208, 0.20), rgba(87, 95, 107, 0.12)),
            linear-gradient(180deg, rgba(86, 91, 100, 0.96), rgba(51, 56, 64, 0.96));
    }
    .file-type-banner--mercedes {
        background:
            linear-gradient(135deg, rgba(170, 182, 198, 0.16), rgba(69, 80, 94, 0.10)),
            linear-gradient(180deg, rgba(41, 52, 66, 0.97), rgba(24, 31, 40, 0.97));
    }
    .file-type-banner--audi {
        background:
            linear-gradient(135deg, rgba(186, 197, 214, 0.16), rgba(95, 111, 135, 0.10)),
            linear-gradient(180deg, rgba(47, 57, 72, 0.97), rgba(26, 33, 43, 0.97));
    }
    .file-type-banner--default {
        background:
            linear-gradient(135deg, rgba(165, 185, 210, 0.13), rgba(71, 92, 122, 0.08)),
            linear-gradient(180deg, rgba(32, 42, 57, 0.97), rgba(16, 23, 34, 0.97));
    }
    .file-type-banner__text {
        position: relative;
        z-index: 1;
        color: #f8fbff;
        font-family: "Aptos Display", "Segoe UI", "Aptos", "Helvetica Neue", Arial, sans-serif;
        font-size: clamp(1.15rem, 1.05rem + 0.65vw, 1.85rem);
        font-weight: 800;
        letter-spacing: 0.18em;
        line-height: 1.1;
        text-transform: uppercase;
        text-wrap: balance;
    }
    .side-panel-divider {
        border: 0;
        border-top: 1px solid var(--line);
        margin: 0.9rem 0;
    }
    .filter-panel-kicker {
        font-size: 0.74rem;
        text-transform: uppercase;
        letter-spacing: 0.14em;
        color: var(--steel);
        font-weight: 800;
        margin-bottom: 0.35rem;
    }
    .filter-panel-title {
        color: var(--ink);
        font-size: 1.15rem;
        font-weight: 800;
        margin-bottom: 0.3rem;
    }
    .filter-panel-copy {
        color: var(--slate);
        font-size: 0.88rem;
        line-height: 1.6;
        margin-bottom: 0.8rem;
    }
    .filter-panel-shell .stRadio > label,
    .filter-panel-shell .stMultiSelect label,
    .filter-panel-shell .stTextInput label,
    .filter-panel-shell .stDateInput label {
        color: var(--ink);
        font-weight: 700;
        letter-spacing: 0.01em;
    }
    .filter-panel-shell h2,
    .filter-panel-shell h3 {
        color: var(--ink);
    }
    div[data-testid="stMetric"] {
        border: 1px solid var(--line);
        background:
            linear-gradient(180deg, rgba(15, 22, 34, 0.98) 0%, rgba(10, 16, 25, 0.96) 100%),
            radial-gradient(circle at top right, rgba(76, 201, 240, 0.06), transparent 30%);
        border-radius: 24px;
        padding: 1rem;
        box-shadow: var(--shadow-md);
    }
    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
        color: var(--slate);
    }
    div[data-testid="stMetricValue"] {
        color: var(--ink);
    }
    div[data-testid="stMetricDelta"] {
        color: var(--steel);
    }
    h1, h2, h3 {
        letter-spacing: -0.02em;
        color: var(--ink);
    }
    .hero-card {
        border: 1px solid var(--line);
        border-radius: 28px;
        padding: 1.7rem 1.7rem;
        background:
            radial-gradient(circle at top right, rgba(76, 201, 240, 0.10), transparent 34%),
            linear-gradient(180deg, rgba(14, 20, 31, 0.99), rgba(10, 15, 24, 0.97));
        box-shadow: var(--shadow-lg);
    }
    .hero-logo {
        width: 168px;
        max-width: 100%;
        height: auto;
        display: block;
        margin-bottom: 1rem;
        filter: drop-shadow(0 18px 34px rgba(0, 0, 0, 0.34));
    }
    .hero-kicker {
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 0.14em;
        color: var(--steel);
        font-weight: 700;
        margin-bottom: 0.6rem;
    }
    .hero-title {
        font-size: 2.65rem;
        line-height: 0.98;
        font-weight: 800;
        color: var(--ink);
        margin-bottom: 0.6rem;
    }
    .hero-copy {
        color: var(--slate);
        font-size: 0.98rem;
        line-height: 1.78;
        margin-bottom: 0;
    }
    .hero-stat-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.8rem;
        margin-top: 1rem;
    }
    .hero-stat {
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 0.9rem 1rem;
        background: rgba(18, 26, 40, 0.72);
    }
    .hero-stat-label {
        font-size: 0.74rem;
        text-transform: uppercase;
        letter-spacing: 0.12em;
        color: var(--muted);
        margin-bottom: 0.35rem;
        font-weight: 700;
    }
    .hero-stat-value {
        color: var(--ink);
        font-size: 1.25rem;
        font-weight: 800;
        line-height: 1.1;
    }
    .compact-header {
        display: grid;
        grid-template-columns: minmax(0, 1.45fr) auto;
        gap: 1rem;
        align-items: center;
        border: 1px solid var(--line);
        border-radius: 28px;
        padding: 1.15rem 1.25rem;
        background:
            radial-gradient(circle at top right, rgba(76, 201, 240, 0.10), transparent 32%),
            linear-gradient(180deg, rgba(14, 20, 31, 0.99), rgba(10, 15, 24, 0.96));
        box-shadow: var(--shadow-lg);
        margin-bottom: 1rem;
    }
    .compact-header-kicker {
        font-size: 0.73rem;
        text-transform: uppercase;
        letter-spacing: 0.14em;
        color: var(--steel);
        font-weight: 800;
        margin-bottom: 0.35rem;
    }
    .compact-header-title {
        color: var(--ink);
        font-size: 1.75rem;
        line-height: 1.02;
        font-weight: 800;
        letter-spacing: -0.03em;
        margin-bottom: 0.45rem;
    }
    .compact-header-copy {
        color: var(--slate);
        font-size: 0.94rem;
        line-height: 1.65;
        margin-bottom: 0.8rem;
    }
    .compact-pill-row {
        display: flex;
        flex-wrap: wrap;
        gap: 0.55rem;
    }
    .compact-pill {
        border: 1px solid var(--line);
        border-radius: 999px;
        padding: 0.4rem 0.78rem;
        background: rgba(17, 24, 39, 0.82);
        color: var(--navy);
        font-size: 0.79rem;
        font-weight: 700;
        line-height: 1;
    }
    .compact-brand-box {
        min-width: 248px;
        border: 1px solid var(--line);
        border-radius: 24px;
        padding: 0.9rem;
        background: rgba(17, 24, 39, 0.72);
        display: grid;
        gap: 0.55rem;
        align-content: start;
    }
    .compact-brand-copy {
        color: var(--muted);
        font-size: 0.78rem;
        line-height: 1.45;
        text-align: center;
    }
    .brand-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        border-radius: 999px;
        padding: 0.35rem 0.8rem;
        background: rgba(17, 24, 39, 0.96);
        color: #ffffff;
        font-size: 0.82rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 0.85rem;
    }
    .upload-card {
        border: 1px solid var(--line);
        border-radius: 24px;
        padding: 1.1rem 1.15rem;
        background:
            radial-gradient(circle at top right, rgba(125, 211, 252, 0.07), transparent 34%),
            linear-gradient(180deg, rgba(15, 21, 31, 0.98), rgba(10, 15, 23, 0.94));
        box-shadow: var(--shadow-md);
        min-height: 132px;
        margin-bottom: 0.55rem;
    }
    .upload-step {
        display: inline-flex;
        align-items: center;
        border-radius: 999px;
        padding: 0.28rem 0.62rem;
        background: rgba(125, 211, 252, 0.11);
        color: var(--accent);
        font-size: 0.75rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.12em;
        margin-bottom: 0.7rem;
    }
    .upload-title {
        color: var(--ink);
        font-size: 1.18rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin-bottom: 0.35rem;
    }
    .upload-copy {
        color: var(--slate);
        font-size: 0.93rem;
        line-height: 1.65;
    }
    section[data-testid="stFileUploader"] {
        border: 1px solid var(--line-strong);
        border-radius: 20px;
        background: rgba(15, 21, 32, 0.86);
        padding: 0.3rem 0.55rem 0.55rem 0.55rem;
        box-shadow: var(--shadow-md);
    }
    div[data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed rgba(125, 211, 252, 0.24);
        border-radius: 18px;
        background: linear-gradient(180deg, rgba(14, 19, 29, 0.96), rgba(10, 15, 24, 0.92));
        padding: 1.15rem 1rem;
    }
    div[data-testid="stFileUploaderDropzoneInstructions"] span {
        color: var(--slate);
        font-weight: 600;
    }
    .quick-card {
        border: 1px solid var(--line);
        border-radius: 22px;
        padding: 1.05rem 1.1rem;
        background: linear-gradient(180deg, rgba(14, 20, 31, 0.96), rgba(10, 15, 24, 0.92));
        box-shadow: var(--shadow-md);
        min-height: 156px;
    }
    .login-shell {
        max-width: 1080px;
        margin: 0 auto 1rem auto;
        padding: 0.25rem 0 1rem 0;
    }
    .login-grid {
        display: grid;
        grid-template-columns: 1.15fr 0.95fr;
        gap: 1rem;
        align-items: stretch;
    }
    .login-brand-card, .login-form-card {
        border: 1px solid var(--line);
        border-radius: 28px;
        padding: 1.55rem 1.6rem;
        background: linear-gradient(180deg, rgba(14, 20, 31, 0.98), rgba(10, 15, 24, 0.95));
        box-shadow: var(--shadow-xl);
        min-height: 460px;
    }
    .login-brand-card {
        background:
            radial-gradient(circle at top right, rgba(76, 201, 240, 0.12), transparent 28%),
            linear-gradient(135deg, rgba(14, 20, 31, 0.98), rgba(10, 15, 24, 0.95));
    }
    .login-kicker {
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 0.14em;
        color: var(--steel);
        font-weight: 800;
        margin-bottom: 0.6rem;
    }
    .login-title {
        font-size: 2.45rem;
        line-height: 0.98;
        letter-spacing: -0.03em;
        font-weight: 800;
        color: var(--ink);
        margin-bottom: 0.75rem;
    }
    .login-copy {
        color: var(--slate);
        font-size: 1rem;
        line-height: 1.75;
        margin-bottom: 1rem;
    }
    .login-points {
        display: grid;
        gap: 0.8rem;
        margin-top: 1rem;
    }
    .login-point {
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 0.9rem 1rem;
        background: rgba(19, 27, 41, 0.70);
    }
    .login-point-title {
        color: var(--ink);
        font-size: 0.96rem;
        font-weight: 800;
        margin-bottom: 0.3rem;
    }
    .login-point-copy {
        color: var(--muted);
        font-size: 0.88rem;
        line-height: 1.6;
    }
    .login-form-heading {
        color: var(--ink);
        font-size: 1.55rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin-bottom: 0.35rem;
    }
    .login-form-copy {
        color: var(--muted);
        font-size: 0.92rem;
        line-height: 1.65;
        margin-bottom: 0.9rem;
    }
    .sidebar-user-card {
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 0.95rem 1rem;
        background: rgba(16, 23, 35, 0.90);
        box-shadow: var(--shadow-md);
        margin-bottom: 0.8rem;
    }
    .sidebar-user-label {
        color: var(--muted);
        text-transform: uppercase;
        font-size: 0.72rem;
        letter-spacing: 0.12em;
        font-weight: 800;
        margin-bottom: 0.25rem;
    }
    .sidebar-user-name {
        color: var(--ink);
        font-size: 1rem;
        font-weight: 800;
        margin-bottom: 0.15rem;
    }
    .sidebar-user-role {
        color: var(--slate);
        font-size: 0.82rem;
    }
    .quick-title {
        color: var(--ink);
        font-size: 1rem;
        font-weight: 800;
        margin-bottom: 0.45rem;
        letter-spacing: -0.01em;
    }
    .quick-copy {
        color: var(--slate);
        font-size: 0.9rem;
        line-height: 1.65;
    }
    .meta-card {
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 1rem 1.15rem;
        background: linear-gradient(180deg, rgba(16, 22, 34, 0.98), rgba(11, 16, 25, 0.94));
        box-shadow: var(--shadow-md);
    }
    .meta-label {
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.12em;
        color: var(--slate);
        margin-bottom: 0.4rem;
    }
    .meta-value {
        font-size: 1rem;
        line-height: 1.6;
        color: var(--ink);
    }
    .pill-row {
        display: flex;
        flex-wrap: wrap;
        gap: 0.55rem;
        margin: 0.5rem 0 0.25rem 0;
    }
    .pill {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        border-radius: 999px;
        padding: 0.45rem 0.8rem;
        font-size: 0.86rem;
        font-weight: 600;
        border: 1px solid var(--line);
        background: rgba(17, 24, 37, 0.88);
        color: var(--ink);
    }
    .pill-positive {
        color: var(--mint);
        border-color: rgba(67, 211, 177, 0.22);
        background: rgba(8, 41, 37, 0.9);
    }
    .pill-negative {
        color: var(--rose);
        border-color: rgba(255, 111, 111, 0.18);
        background: rgba(47, 17, 17, 0.92);
    }
    .pill-neutral {
        color: var(--steel);
        border-color: rgba(125, 211, 252, 0.18);
        background: rgba(13, 29, 43, 0.92);
    }
    .finding-card {
        border: 1px solid var(--line);
        border-radius: 24px;
        padding: 1.2rem 1.25rem;
        background:
            radial-gradient(circle at top right, rgba(76, 201, 240, 0.08), transparent 36%),
            linear-gradient(180deg, rgba(15, 21, 32, 0.98), rgba(10, 15, 24, 0.95));
        min-height: 196px;
        box-shadow: var(--shadow-md);
    }
    .finding-label {
        font-size: 0.76rem;
        text-transform: uppercase;
        letter-spacing: 0.12em;
        color: var(--accent);
        margin-bottom: 0.55rem;
        font-weight: 700;
    }
    .finding-title {
        font-size: 1.52rem;
        color: var(--ink);
        font-weight: 800;
        margin-bottom: 0.5rem;
        line-height: 1.2;
        letter-spacing: -0.02em;
    }
    .finding-copy {
        font-size: 0.82rem;
        line-height: 1.62;
        color: var(--slate);
    }
    div[data-testid="stVegaLiteChart"],
    div[data-testid="stDataFrame"] {
        border-radius: 24px;
        overflow: hidden;
        border: 1px solid var(--line);
        box-shadow: var(--shadow-md);
        background: linear-gradient(180deg, rgba(14, 20, 31, 0.98), rgba(10, 15, 24, 0.96));
    }
    div[data-testid="stAlert"] {
        border-radius: 18px;
        border: 1px solid var(--line);
        background: rgba(14, 20, 31, 0.92);
    }
    .section-banner {
        margin: 0.25rem 0 0.95rem 0;
    }
    .section-kicker {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.14em;
        color: var(--accent);
        font-weight: 800;
        margin-bottom: 0.25rem;
    }
    .section-copy {
        color: var(--slate);
        font-size: 0.96rem;
        line-height: 1.72;
    }
    [data-testid="stMarkdownContainer"] p {
        color: var(--slate);
    }
    .stMarkdown a {
        color: var(--accent);
    }
    button[kind="secondary"] {
        border-radius: 14px;
        border: 1px solid var(--line-strong);
        background: linear-gradient(180deg, rgba(19, 27, 41, 0.96), rgba(13, 20, 31, 0.94));
        color: var(--ink);
    }
    button[kind="secondary"]:hover,
    button[kind="primary"]:hover {
        border-color: rgba(125, 211, 252, 0.28);
        color: #ffffff;
    }
    button[kind="primary"] {
        border-radius: 16px;
        border: 1px solid rgba(125, 211, 252, 0.28);
        background: linear-gradient(135deg, rgba(18, 40, 56, 0.96), rgba(17, 31, 46, 0.94));
        color: #ffffff;
        box-shadow: 0 14px 32px rgba(7, 13, 20, 0.34);
    }
    div[data-baseweb="tab-list"] {
        gap: 0.5rem;
    }
    button[data-baseweb="tab"] {
        border-radius: 999px;
        background: rgba(16, 24, 37, 0.82);
        border: 1px solid var(--line);
        padding: 0.5rem 0.95rem;
        box-shadow: none;
        color: var(--slate);
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, rgba(19, 44, 61, 0.96), rgba(12, 28, 41, 0.92));
        color: white;
        border-color: rgba(125, 211, 252, 0.28);
    }
    div[data-baseweb="input"] > div,
    div[data-baseweb="base-input"] > div,
    div[data-baseweb="select"] > div,
    .stDateInput > div > div,
    .stMultiSelect [data-baseweb="tag"],
    .stTextInput > div > div > input {
        background: rgba(15, 21, 32, 0.90);
        color: var(--ink);
        border-color: var(--line-strong);
    }
    .stTextInput input,
    .stDateInput input {
        color: var(--ink) !important;
    }
    .stCheckbox label,
    .stRadio label {
        color: var(--slate) !important;
    }
    .stSelectbox label,
    .stMultiSelect label,
    .stTextInput label,
    .stDateInput label {
        color: var(--ink) !important;
    }
    .stDownloadButton button {
        width: 100%;
    }
    [data-testid="stToolbar"] {
        display: none;
    }
    /* We provide our own Wykres / Dane switch, so hide Vega/Altair action menus. */
    .vega-embed details,
    .vega-embed .vega-actions,
    .stVegaLiteChart details,
    .stVegaLiteChart summary {
        display: none !important;
    }
    /* Hide Streamlit's fullscreen toolbar for Vega/Altair cards to keep charts clean. */
    div[data-testid="stFullScreenFrame"]:has(div[data-testid="stVegaLiteChart"])
        [data-testid="stElementToolbar"] {
        display: none !important;
    }
    @media (max-width: 920px) {
        .hero-title {
            font-size: 2rem;
        }
        .hero-stat-grid {
            grid-template-columns: 1fr;
        }
        .login-grid {
            grid-template-columns: 1fr;
        }
        .compact-header {
            grid-template-columns: 1fr;
        }
        .compact-brand-box {
            justify-items: start;
            text-align: left;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    """
    <style>
    :root {
        --bg-primary: #0b1120;
        --bg-secondary: #111827;
        --bg-tertiary: #162033;
        --bg-elevated: rgba(17, 24, 39, 0.88);
        --surface-soft: rgba(15, 23, 42, 0.72);
        --text-primary: #f8fafc;
        --text-secondary: #cbd5e1;
        --text-muted: #94a3b8;
        --border-soft: rgba(148, 163, 184, 0.16);
        --border-strong: rgba(148, 163, 184, 0.28);
        --accent-blue: #38bdf8;
        --accent-green: #34d399;
        --accent-red: #f87171;
        --accent-amber: #fbbf24;
        --shadow-soft: 0 18px 40px rgba(2, 6, 23, 0.24);
    }
    .stApp {
        background: linear-gradient(180deg, #09101e 0%, #0d1528 48%, #101829 100%) !important;
        color: var(--text-primary) !important;
    }
    .block-container {
        max-width: 1680px !important;
        padding-top: 0.9rem !important;
        padding-bottom: 2.2rem !important;
    }
    h1, h2, h3, h4 {
        color: var(--text-primary) !important;
        font-family: "Aptos Display", "Segoe UI", "Aptos", "Helvetica Neue", Arial, sans-serif !important;
        letter-spacing: -0.03em !important;
        line-height: 1.05 !important;
        text-wrap: balance;
    }
    p, label, span, div {
        text-wrap: pretty;
    }
    .app-shell-header {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        gap: 0.85rem;
        align-items: end;
        border: 1px solid var(--border-soft);
        border-radius: 22px;
        padding: 0.95rem 1.05rem;
        margin-bottom: 0.9rem;
        background: linear-gradient(180deg, rgba(15, 23, 42, 0.88), rgba(15, 23, 42, 0.62));
        box-shadow: none;
    }
    .app-shell-kicker,
    .app-header__eyebrow {
        font-size: 0.74rem;
        font-weight: 800;
        letter-spacing: 0.16em;
        text-transform: uppercase;
        color: var(--accent-blue);
        margin-bottom: 0.45rem;
    }
    .app-shell-title,
    .app-header__title {
        font-size: clamp(1.45rem, 1.18rem + 0.8vw, 2.35rem);
        font-weight: 800;
        color: var(--text-primary);
        margin-bottom: 0.4rem;
    }
    .app-shell-copy,
    .app-header__subtitle {
        color: var(--text-secondary);
        font-size: 0.93rem;
        line-height: 1.6;
        max-width: 70ch;
    }
    .app-shell-chip {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border-radius: 999px;
        border: 1px solid rgba(56, 189, 248, 0.22);
        background: rgba(56, 189, 248, 0.08);
        color: var(--text-primary);
        padding: 0.48rem 0.82rem;
        font-size: 0.76rem;
        font-weight: 700;
        white-space: nowrap;
    }
    .app-header {
        display: grid;
        grid-template-columns: minmax(0, 1.55fr) minmax(220px, 0.75fr);
        gap: 1rem;
        align-items: start;
        border: 1px solid var(--border-soft);
        border-radius: 24px;
        padding: 1.2rem;
        margin-bottom: 1rem;
        background: linear-gradient(180deg, rgba(15, 23, 42, 0.9), rgba(15, 23, 42, 0.72));
        box-shadow: var(--shadow-soft);
    }
    .app-header__banner {
        display: grid;
        gap: 0.5rem;
    }
    .app-header-caption {
        color: var(--text-muted);
        font-size: 0.8rem;
        line-height: 1.5;
        text-align: center;
    }
    .context-chip-row {
        display: flex;
        flex-wrap: wrap;
        gap: 0.5rem;
        margin-top: 0.85rem;
    }
    .context-chip,
    .compact-pill {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        border-radius: 999px;
        padding: 0.42rem 0.78rem;
        border: 1px solid var(--border-soft);
        background: rgba(148, 163, 184, 0.08);
        color: var(--text-secondary);
        font-size: 0.78rem;
        font-weight: 700;
        line-height: 1;
    }
    .file-type-banner {
        border-radius: 20px !important;
        border: 1px solid var(--border-soft) !important;
        box-shadow: none !important;
        background: linear-gradient(180deg, rgba(30, 41, 59, 0.88), rgba(15, 23, 42, 0.92)) !important;
    }
    .file-type-banner::before {
        opacity: 0.35 !important;
    }
    .file-type-banner--tesla {
        background: linear-gradient(180deg, rgba(100, 116, 139, 0.32), rgba(51, 65, 85, 0.92)) !important;
    }
    .file-type-banner--mercedes {
        background: linear-gradient(180deg, rgba(71, 85, 105, 0.4), rgba(15, 23, 42, 0.92)) !important;
    }
    .file-type-banner--audi {
        background: linear-gradient(180deg, rgba(59, 72, 89, 0.42), rgba(15, 23, 42, 0.92)) !important;
    }
    .file-type-banner--default {
        background: linear-gradient(180deg, rgba(30, 41, 59, 0.88), rgba(15, 23, 42, 0.92)) !important;
    }
    .file-type-banner__text {
        letter-spacing: 0.16em !important;
        font-size: clamp(1.05rem, 0.96rem + 0.55vw, 1.55rem) !important;
    }
    .empty-state-shell {
        display: grid;
        grid-template-columns: minmax(0, 1.25fr) minmax(220px, 280px);
        gap: 1rem;
        align-items: stretch;
        border: 1px solid var(--border-soft);
        border-radius: 24px;
        padding: 1.15rem 1.2rem;
        margin-bottom: 1rem;
        background: linear-gradient(180deg, rgba(15, 23, 42, 0.84), rgba(15, 23, 42, 0.62));
    }
    .empty-state-copy {
        display: flex;
        flex-direction: column;
        gap: 0.72rem;
        justify-content: center;
        min-width: 0;
    }
    .empty-state-kicker {
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.16em;
        color: var(--accent-blue);
        font-weight: 800;
    }
    .empty-state-title {
        color: var(--text-primary);
        font-size: clamp(1.35rem, 1.12rem + 0.75vw, 2.05rem);
        font-weight: 800;
        line-height: 1.2;
        max-width: 18ch;
    }
    .empty-state-subtitle {
        color: var(--text-secondary);
        font-size: 0.94rem;
        line-height: 1.65;
        max-width: 68ch;
    }
    .empty-state-banner {
        display: flex;
        align-items: center;
    }
    .empty-state-banner .file-type-banner {
        width: 100%;
        min-height: 120px;
    }
    .filter-panel-shell,
    .upload-card,
    .quick-card,
    .meta-card,
    .finding-card,
    .compact-brand-box,
    .sidebar-user-card {
        border: 1px solid var(--border-soft) !important;
        background: linear-gradient(180deg, rgba(15, 23, 42, 0.9), rgba(15, 23, 42, 0.74)) !important;
        box-shadow: none !important;
        border-radius: 20px !important;
    }
    .filter-panel-shell {
        padding: 1rem 1rem 1.05rem 1rem !important;
        margin-bottom: 0.9rem;
        position: sticky;
        top: 0.85rem;
        backdrop-filter: blur(12px);
    }
    .filter-panel-title,
    .upload-title,
    .quick-title,
    .finding-title {
        color: var(--text-primary) !important;
    }
    .filter-panel-copy,
    .upload-copy,
    .quick-copy,
    .finding-copy,
    .meta-value,
    .sidebar-user-role {
        color: var(--text-secondary) !important;
    }
    .section-head {
        margin: 0.2rem 0 0.85rem 0;
    }
    .section-title {
        color: var(--text-primary);
        font-size: 1.18rem;
        font-weight: 800;
        margin-bottom: 0.25rem;
        letter-spacing: -0.02em;
    }
    .section-copy {
        color: var(--text-muted);
        font-size: 0.92rem;
        line-height: 1.65;
        max-width: 72ch;
    }
    .report-metadata-grid {
        display: grid;
        grid-template-columns: repeat(5, minmax(0, 1fr));
        gap: 0.75rem;
        margin-bottom: 1rem;
    }
    .report-meta-card {
        border: 1px solid var(--border-soft);
        border-radius: 18px;
        padding: 0.9rem 1rem;
        background: rgba(15, 23, 42, 0.72);
    }
    .report-meta-label {
        color: var(--text-muted);
        font-size: 0.72rem;
        font-weight: 800;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        margin-bottom: 0.35rem;
    }
    .report-meta-value {
        color: var(--text-primary);
        font-size: 0.97rem;
        font-weight: 700;
        line-height: 1.45;
    }
    .kpi-card {
        border: 1px solid var(--border-soft);
        border-radius: 18px;
        padding: 1rem 1rem 0.95rem 1rem;
        background: rgba(15, 23, 42, 0.78);
        min-height: 138px;
    }
    .kpi-card--positive {
        border-color: rgba(52, 211, 153, 0.24);
    }
    .kpi-card--negative {
        border-color: rgba(248, 113, 113, 0.24);
    }
    .kpi-label {
        color: var(--text-muted);
        font-size: 0.78rem;
        font-weight: 800;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 0.5rem;
    }
    .kpi-value {
        color: var(--text-primary);
        font-size: 1.65rem;
        font-weight: 800;
        letter-spacing: -0.04em;
        margin-bottom: 0.45rem;
        line-height: 1;
    }
    .kpi-copy {
        color: var(--text-secondary);
        font-size: 0.84rem;
        line-height: 1.55;
    }
    .insight-card {
        border: 1px solid var(--border-soft);
        border-radius: 18px;
        padding: 1rem;
        background: rgba(15, 23, 42, 0.76);
        min-height: 190px;
    }
    .insight-card--critical {
        border-color: rgba(248, 113, 113, 0.26);
        background: linear-gradient(180deg, rgba(69, 20, 26, 0.42), rgba(15, 23, 42, 0.78));
    }
    .insight-card--positive {
        border-color: rgba(52, 211, 153, 0.24);
    }
    .insight-badge {
        display: inline-flex;
        align-items: center;
        border-radius: 999px;
        padding: 0.34rem 0.68rem;
        margin-bottom: 0.75rem;
        background: rgba(148, 163, 184, 0.1);
        color: var(--text-primary);
        font-size: 0.74rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }
    .insight-title {
        color: var(--text-primary);
        font-size: 1.08rem;
        font-weight: 800;
        line-height: 1.25;
        margin-bottom: 0.45rem;
    }
    .insight-copy {
        color: var(--text-secondary);
        font-size: 0.88rem;
        line-height: 1.6;
    }
    .upload-status-grid {
        display: grid;
        gap: 0.7rem;
        margin-top: 0.85rem;
        margin-bottom: 0.85rem;
    }
    .upload-status-card {
        border: 1px solid var(--border-soft);
        border-radius: 18px;
        padding: 0.9rem 1rem;
        background: rgba(15, 23, 42, 0.72);
    }
    .upload-status-card--ready {
        border-color: rgba(52, 211, 153, 0.24);
    }
    .upload-status-card--pending {
        border-color: rgba(56, 189, 248, 0.22);
    }
    .upload-status-label {
        color: var(--text-muted);
        font-size: 0.72rem;
        font-weight: 800;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        margin-bottom: 0.35rem;
    }
    .upload-status-name {
        color: var(--text-primary);
        font-size: 0.95rem;
        font-weight: 800;
        line-height: 1.35;
        margin-bottom: 0.25rem;
        word-break: break-word;
    }
    .upload-status-meta {
        color: var(--text-secondary);
        font-size: 0.82rem;
        line-height: 1.5;
        margin-bottom: 0.2rem;
    }
    .upload-status-caption {
        color: var(--text-muted);
        font-size: 0.78rem;
        line-height: 1.5;
    }
    div[data-testid="stMetric"] {
        background: rgba(15, 23, 42, 0.76) !important;
        border: 1px solid var(--border-soft) !important;
        box-shadow: none !important;
        border-radius: 18px !important;
    }
    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
        color: var(--text-muted) !important;
    }
    div[data-testid="stMetricValue"] {
        color: var(--text-primary) !important;
    }
    div[data-testid="stMetricDelta"] {
        color: var(--accent-blue) !important;
    }
    div[data-testid="stAlert"] {
        border-radius: 18px !important;
        border: 1px solid var(--border-soft) !important;
        background: rgba(15, 23, 42, 0.84) !important;
        box-shadow: none !important;
    }
    div[data-testid="stVegaLiteChart"],
    div[data-testid="stDataFrame"] {
        border-radius: 20px !important;
        border: 1px solid var(--border-soft) !important;
        background: rgba(15, 23, 42, 0.78) !important;
        box-shadow: none !important;
    }
    section[data-testid="stFileUploader"] {
        border: 1px solid var(--border-soft) !important;
        border-radius: 18px !important;
        background: rgba(15, 23, 42, 0.74) !important;
        box-shadow: none !important;
        padding: 0.35rem 0.5rem 0.55rem 0.5rem !important;
    }
    div[data-testid="stFileUploaderDropzone"] {
        border: 1px dashed rgba(148, 163, 184, 0.26) !important;
        border-radius: 14px !important;
        background: transparent !important;
        padding: 1rem 0.9rem !important;
    }
    div[data-baseweb="input"] > div,
    div[data-baseweb="base-input"] > div,
    div[data-baseweb="select"] > div,
    .stDateInput > div > div,
    .stMultiSelect [data-baseweb="tag"],
    .stTextInput > div > div > input {
        background: rgba(15, 23, 42, 0.78) !important;
        border-color: var(--border-soft) !important;
        color: var(--text-primary) !important;
        border-radius: 14px !important;
    }
    [data-testid="stButtonGroup"] {
        width: 100%;
        margin-bottom: 0.2rem;
    }
    [data-testid="stButtonGroup"] > div {
        width: 100%;
    }
    [data-testid="stButtonGroup"] button {
        border-radius: 12px !important;
        border: 1px solid var(--border-soft) !important;
        background: rgba(15, 23, 42, 0.78) !important;
        color: var(--text-secondary) !important;
    }
    [data-testid="stButtonGroup"] button[kind*="Active"] {
        background: rgba(56, 189, 248, 0.14) !important;
        border-color: rgba(56, 189, 248, 0.24) !important;
        color: var(--text-primary) !important;
    }
    div[data-baseweb="tab-list"] {
        gap: 0.4rem !important;
        margin-bottom: 0.7rem;
    }
    button[data-baseweb="tab"] {
        border-radius: 999px !important;
        background: rgba(15, 23, 42, 0.76) !important;
        border: 1px solid var(--border-soft) !important;
        color: var(--text-secondary) !important;
        padding: 0.48rem 0.92rem !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background: rgba(56, 189, 248, 0.14) !important;
        border-color: rgba(56, 189, 248, 0.22) !important;
        color: var(--text-primary) !important;
    }
    button[kind="primary"],
    button[kind="secondary"] {
        border-radius: 14px !important;
        box-shadow: none !important;
        border: 1px solid var(--border-soft) !important;
    }
    button[kind="primary"] {
        background: rgba(56, 189, 248, 0.14) !important;
        color: var(--text-primary) !important;
        border-color: rgba(56, 189, 248, 0.24) !important;
    }
    button[kind="secondary"] {
        background: rgba(15, 23, 42, 0.76) !important;
        color: var(--text-primary) !important;
    }
    .stDownloadButton button {
        width: 100%;
    }
    @media (max-width: 1200px) {
        .report-metadata-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr));
        }
        .app-header {
            grid-template-columns: 1fr;
        }
    }
    @media (max-width: 920px) {
        .app-shell-header {
            grid-template-columns: 1fr;
        }
        .empty-state-shell {
            grid-template-columns: 1fr;
        }
        .report-metadata-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
    }
    @media (max-width: 640px) {
        .report-metadata-grid {
            grid-template-columns: 1fr;
        }
    }
    </style>
    <div class="app-shell-header">
        <div>
            <div class="app-shell-kicker">Release Intelligence</div>
            <div class="app-shell-title">Dashboard porownan release'ow</div>
            <div class="app-shell-copy">
                Upload, filtry i eksport pozostaja w jednym miejscu do codziennej analizy planistycznej
                i logistycznej.
            </div>
        </div>
        <div class="app-shell-chip">Workspace</div>
    </div>
    """,
    unsafe_allow_html=True,
)


def first_non_empty(series):
    values = series.dropna().astype(str).str.strip()
    values = values[values.ne("")]
    return values.iloc[0] if not values.empty else "n/a"


def format_date(value):
    if pd.isna(value):
        return "n/a"
    return pd.Timestamp(value).strftime("%Y-%m-%d")


def format_signed_int(value):
    return f"{float(value):+,.0f}"


def format_signed_pct(value):
    return f"{float(value):+,.1f}%"


def get_date_label(value):
    return DATE_LABELS.get(value, value)


def get_change_label(value):
    return CHANGE_DIRECTION_LABELS.get(value, value)


def get_metric_label(value):
    return MATRIX_METRIC_LABELS.get(value, value)


def get_view_mode_label(value):
    return VIEW_MODE_LABELS.get(value, value)


def format_release_label(meta):
    release_version = str(meta.get("release_version", "")).strip()
    release_date = meta.get("release_date")
    if release_version and release_version.lower() != "n/a":
        return f"v{release_version}"
    if not pd.isna(release_date):
        return f"Snapshot {format_date(release_date)}"
    return "n/a"


def format_release_summary(meta):
    release_label = format_release_label(meta)
    release_date = meta.get("release_date")
    if release_label == "n/a":
        return "n/a"
    if pd.isna(release_date):
        return release_label
    return f"{release_label} / {format_date(release_date)}"


def available_detail_columns(dataframe):
    preferred_columns = [
        "PO Number",
        "Origin Doc",
        "Item",
        "Ship To",
        "Part Number",
        "Part Description",
        "Customer Material",
        "Unrestricted Qty",
        "Unloading Point",
        "Ship Date",
        "Receipt Date",
        "Unit of Measure",
        "CumQty",
        "Quantity_Prev",
        "Quantity_Curr",
        "Delta",
        "Percent Change",
        "Demand Status",
        "Change Direction",
        "Alert",
    ]
    return [column for column in preferred_columns if column in dataframe.columns]


def format_chart_source_table(dataframe):
    source_table = dataframe.copy()
    for column in source_table.columns:
        if pd.api.types.is_datetime64_any_dtype(source_table[column]):
            source_table[column] = source_table[column].dt.strftime("%Y-%m-%d")
        elif pd.api.types.is_bool_dtype(source_table[column]):
            source_table[column] = source_table[column].map(
                lambda value: "Tak" if value else "Nie"
            )
    return source_table


def render_chart_table_switch(
    key,
    chart,
    source_df,
    *,
    chart_empty_message="Brak danych do wykresu.",
    table_empty_message="Brak danych źródłowych.",
    table_height=320,
):
    state_key = f"{key}_view_mode"
    st.session_state.setdefault(state_key, "chart")
    selected_view = st.segmented_control(
        "Widok sekcji",
        options=["chart", "table"],
        selection_mode="single",
        default=st.session_state[state_key],
        required=True,
        key=state_key,
        label_visibility="collapsed",
        format_func=get_view_mode_label,
        width="stretch",
    )
    selected_view = selected_view or st.session_state[state_key]

    if selected_view == "chart":
        if chart is None:
            st.info(chart_empty_message)
        else:
            st.altair_chart(chart, use_container_width=True)
        return

    source_table = format_chart_source_table(source_df)
    if source_table.empty:
        st.info(table_empty_message)
    else:
        st.dataframe(source_table, use_container_width=True, height=table_height)


def render_meta_card(title, body_lines):
    body_html = "<br>".join(body_lines)
    st.markdown(
        f"""
        <div class="meta-card">
            <div class="meta-label">{title}</div>
            <div class="meta-value">{body_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_status_pills(total_delta, alert_count, products_changed):
    delta_class = "pill-positive" if total_delta > 0 else "pill-negative" if total_delta < 0 else "pill-neutral"
    st.markdown(
        f"""
        <div class="pill-row">
            <div class="pill {delta_class}">Bilans zmian {format_signed_int(total_delta)}</div>
            <div class="pill pill-neutral">Alerty {alert_count:,}</div>
            <div class="pill pill-neutral">Zmienne produkty {products_changed:,}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_finding_card(label, title, copy):
    st.markdown(
        f"""
        <div class="finding-card">
            <div class="finding-label">{html.escape(str(label))}</div>
            <div class="finding-title">{html.escape(str(title))}</div>
            <div class="finding-copy">{html.escape(str(copy))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_upload_card(step, title, copy):
    st.markdown(
        f"""
        <div class="upload-card">
            <div class="upload-step">{step}</div>
            <div class="upload-title">{title}</div>
            <div class="upload-copy">{copy}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_quick_card(title, copy):
    st.markdown(
        f"""
        <div class="quick-card">
            <div class="quick-title">{title}</div>
            <div class="quick-copy">{copy}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def format_file_type_label(file_type):
    mapping = {
        "legacy_wide": "Tesla / ReleaseData",
        "vl10e_block": "Mercedes / VL10E",
        "cw_weekly_pivot": "Audi Q7/Q9 weekly",
    }
    return mapping.get(str(file_type or "").strip().lower(), "Nieznany format")


def guess_file_type_label(file_name):
    normalized = str(file_name or "").strip().lower()
    if any(keyword in normalized for keyword in ["vl10e", "mercedes", "merc"]):
        return "Mercedes / VL10E"
    if any(keyword in normalized for keyword in ["audi", "q7", "q9", "megatech", "cw17"]):
        return "Audi Q7/Q9 weekly"
    if any(keyword in normalized for keyword in ["tesla", "releasedata", "raw"]):
        return "Tesla / ReleaseData"
    return "Oczekuje na rozpoznanie parsera"


def count_status_matches(dataframe, *keywords):
    if dataframe.empty or "Demand Status" not in dataframe.columns:
        return 0
    normalized = dataframe["Demand Status"].fillna("").astype(str).str.lower()
    return int(
        normalized.apply(lambda value: any(keyword in value for keyword in keywords)).sum()
    )


def render_section_header(kicker, title, copy=None):
    copy_html = (
        f'<div class="section-copy">{html.escape(str(copy))}</div>'
        if copy
        else ""
    )
    markup = (
        '<div class="section-head">'
        f'<div class="section-kicker">{html.escape(str(kicker))}</div>'
        f'<div class="section-title">{html.escape(str(title))}</div>'
        f"{copy_html}"
        "</div>"
    )
    st.markdown(markup, unsafe_allow_html=True)


def render_app_header(brand_context, title, subtitle, meta_items=None, file_caption=""):
    meta_items = meta_items or []
    chips_html = "".join(
        f'<div class="context-chip">{html.escape(str(item))}</div>'
        for item in meta_items
        if item
    )
    chips_html = (
        f'<div class="context-chip-row">{chips_html}</div>' if chips_html else ""
    )
    banner_html = build_file_type_banner_markup(brand_context, variant="header")
    caption_html = (
        f'<div class="app-header-caption">{html.escape(str(file_caption))}</div>'
        if file_caption
        else ""
    )
    header_markup = (
        '<div class="app-header">'
        '<div class="app-header__copy">'
        '<div class="app-header__eyebrow">Pjoter Development Analytics</div>'
        f'<div class="app-header__title">{html.escape(str(title))}</div>'
        f'<div class="app-header__subtitle">{html.escape(str(subtitle))}</div>'
        f"{chips_html}"
        "</div>"
        '<div class="app-header__banner">'
        f"{banner_html}"
        f"{caption_html}"
        "</div>"
        "</div>"
    )
    st.markdown(header_markup, unsafe_allow_html=True)


def render_empty_state_header(brand_context, title, subtitle, meta_items=None):
    meta_items = meta_items or []
    chips_html = "".join(
        f'<div class="context-chip">{html.escape(str(item))}</div>'
        for item in meta_items
        if item
    )
    chips_html = (
        f'<div class="context-chip-row">{chips_html}</div>' if chips_html else ""
    )
    markup = (
        '<div class="empty-state-shell">'
        '<div class="empty-state-copy">'
        '<div class="empty-state-kicker">Workspace status</div>'
        f'<div class="empty-state-title">{html.escape(str(title))}</div>'
        f'<div class="empty-state-subtitle">{html.escape(str(subtitle))}</div>'
        f"{chips_html}"
        "</div>"
        '<div class="empty-state-banner">'
        f'{build_file_type_banner_markup(brand_context, variant="header")}'
        "</div>"
        "</div>"
    )
    st.markdown(markup, unsafe_allow_html=True)


def render_report_metadata(items):
    cards_html = "".join(
        (
            '<div class="report-meta-card">'
            f'<div class="report-meta-label">{html.escape(str(item.get("label", "")))}</div>'
            f'<div class="report-meta-value">{html.escape(str(item.get("value", "n/a")))}</div>'
            "</div>"
        )
        for item in items
    )
    st.markdown(
        f'<div class="report-metadata-grid">{cards_html}</div>',
        unsafe_allow_html=True,
    )


def render_kpi_cards(metrics):
    metric_cols = st.columns(len(metrics), gap="medium")
    for index, metric in enumerate(metrics):
        tone = html.escape(str(metric.get("tone", "neutral")))
        with metric_cols[index]:
            st.markdown(
                f"""
                <div class="kpi-card kpi-card--{tone}">
                    <div class="kpi-label">{html.escape(str(metric.get('label', '')))}</div>
                    <div class="kpi-value">{html.escape(str(metric.get('value', '0')))}</div>
                    <div class="kpi-copy">{html.escape(str(metric.get('copy', '')))}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def build_kpi_metrics(filtered_df, product_summary):
    increase_total = filtered_df.loc[filtered_df["Delta"] > 0, "Delta"].sum()
    decrease_total = abs(filtered_df.loc[filtered_df["Delta"] < 0, "Delta"].sum())
    changed_rows = int((filtered_df["Delta"] != 0).sum())
    new_positions = count_status_matches(filtered_df, "new")
    removed_positions = count_status_matches(filtered_df, "removed", "delete")
    percent_series = pd.to_numeric(filtered_df["Percent Change"], errors="coerce")
    finite_percent = percent_series[pd.notna(percent_series) & percent_series.ne(float("inf")) & percent_series.ne(float("-inf"))]
    largest_percent = (
        format_signed_pct(finite_percent.loc[finite_percent.abs().idxmax()])
        if not finite_percent.empty
        else "n/a"
    )

    return [
        {
            "label": "Liczba zmian",
            "value": f"{changed_rows:,}",
            "copy": "Wiersze z inną ilością niż w poprzednim release.",
            "tone": "neutral",
        },
        {
            "label": "Łączny wzrost",
            "value": f"{increase_total:,.0f}",
            "copy": "Suma dodatnich zmian w aktualnym widoku.",
            "tone": "positive",
        },
        {
            "label": "Łączny spadek",
            "value": f"{decrease_total:,.0f}",
            "copy": "Suma spadków wymagających weryfikacji.",
            "tone": "negative",
        },
        {
            "label": "Nowe pozycje",
            "value": f"{new_positions:,}",
            "copy": "Wiersze oznaczone jako nowy demand.",
            "tone": "neutral",
        },
        {
            "label": "Usunięte pozycje",
            "value": f"{removed_positions:,}",
            "copy": "Wiersze oznaczone jako removed demand.",
            "tone": "negative",
        },
        {
            "label": "Największa zmiana %",
            "value": largest_percent,
            "copy": f"Największy ruch procentowy w {product_summary['Part Number'].nunique():,} produktach.",
            "tone": "neutral",
        },
    ]


def build_alert_items(filtered_df, key_findings):
    alert_items = []
    alert_count = int(filtered_df["Alert"].sum())
    if alert_count:
        alert_items.append(
            {
                "badge": "Wysoki priorytet",
                "title": f"{alert_count:,} pozycji przekracza próg {THRESHOLD}%",
                "copy": "Te wiersze mają największy potencjał wpływu na plan i warto je sprawdzić w pierwszej kolejności.",
                "tone": "critical",
            }
        )

    new_positions = count_status_matches(filtered_df, "new")
    if new_positions:
        alert_items.append(
            {
                "badge": "Nowy demand",
                "title": f"{new_positions:,} nowych pozycji w aktualnym zakresie",
                "copy": "Pojawiły się nowe linie zapotrzebowania, które nie występowały w poprzednim release.",
                "tone": "positive",
            }
        )

    removed_positions = count_status_matches(filtered_df, "removed", "delete")
    if removed_positions:
        alert_items.append(
            {
                "badge": "Removed demand",
                "title": f"{removed_positions:,} pozycji zostało usuniętych",
                "copy": "Warto potwierdzić, czy zniknięcie tych pozycji jest oczekiwane biznesowo.",
                "tone": "negative",
            }
        )

    for finding in key_findings:
        alert_items.append(
            {
                "badge": finding["label"],
                "title": finding["title"],
                "copy": finding["copy"],
                "tone": "neutral",
            }
        )

    return alert_items[:4]


def render_alerts(alert_items):
    if not alert_items:
        st.markdown(
            """
            <div class="insight-card insight-card--neutral">
                <div class="insight-badge">Stabilny zakres</div>
                <div class="insight-title">Brak istotnych alertów w aktywnym widoku</div>
                <div class="insight-copy">Po zastosowanych filtrach nie ma sygnałów, które przekraczałyby próg ostrzegawczy.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    insight_cols = st.columns(len(alert_items), gap="medium")
    for index, item in enumerate(alert_items):
        tone = html.escape(str(item.get("tone", "neutral")))
        with insight_cols[index]:
            st.markdown(
                f"""
                <div class="insight-card insight-card--{tone}">
                    <div class="insight-badge">{html.escape(str(item.get('badge', 'Insight')))}</div>
                    <div class="insight-title">{html.escape(str(item.get('title', '')))}</div>
                    <div class="insight-copy">{html.escape(str(item.get('copy', '')))}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def build_file_slot_payload(slot_label, file_obj=None, meta=None):
    if meta:
        return {
            "slot": slot_label,
            "status": "Załadowany",
            "name": meta.get("file_name", "n/a"),
            "detail": format_file_type_label(meta.get("file_type")),
            "caption": f"Release {format_release_summary(meta)}",
            "tone": "ready",
        }
    if file_obj is not None:
        return {
            "slot": slot_label,
            "status": "Plik dodany",
            "name": file_obj.name,
            "detail": guess_file_type_label(file_obj.name),
            "caption": "Plik czeka na wspólne uruchomienie analizy.",
            "tone": "pending",
        }
    return {
        "slot": slot_label,
        "status": "Oczekiwanie",
        "name": "Brak pliku",
        "detail": "Dodaj plik wejściowy",
        "caption": "Sekcja uzupełni się po dodaniu pliku.",
        "tone": "empty",
    }


def render_file_slot_cards(prev_file=None, current_file=None, prev_meta=None, curr_meta=None):
    slots = [
        build_file_slot_payload("Poprzedni plik", prev_file, prev_meta),
        build_file_slot_payload("Aktualny plik", current_file, curr_meta),
    ]
    markup = "".join(
        (
            f'<div class="upload-status-card upload-status-card--{html.escape(str(slot["tone"]))}">'
            f'<div class="upload-status-label">{html.escape(str(slot["slot"]))}</div>'
            f'<div class="upload-status-name">{html.escape(str(slot["name"]))}</div>'
            f'<div class="upload-status-meta">{html.escape(str(slot["status"]))} | {html.escape(str(slot["detail"]))}</div>'
            f'<div class="upload-status-caption">{html.escape(str(slot["caption"]))}</div>'
            "</div>"
        )
        for slot in slots
    )
    st.markdown(
        f'<div class="upload-status-grid">{markup}</div>',
        unsafe_allow_html=True,
    )


def render_upload_section():
    render_section_header(
        "Workspace",
        "Pliki wejściowe",
        "Dodaj poprzedni i aktualny release. Panel po lewej utrzymuje cały kontekst analizy w jednym miejscu.",
    )
    render_upload_card(
        "Poprzedni",
        "Baseline release",
        "Plik referencyjny, do którego porównywany będzie aktualny stan planu i wysyłek.",
    )
    prev_file = st.file_uploader(
        "Upload Previous Release",
        type=["xlsx"],
        key="previous_release_upload",
        label_visibility="collapsed",
    )
    render_upload_card(
        "Aktualny",
        "Current release",
        "Nowy plik wejściowy, z którego aplikacja policzy zmiany, alerty i bilans wolumenu.",
    )
    current_file = st.file_uploader(
        "Upload Current Release",
        type=["xlsx"],
        key="current_release_upload",
        label_visibility="collapsed",
    )
    return prev_file, current_file


def render_export_actions(csv_bytes, excel_bytes):
    render_section_header(
        "Eksport",
        "Pobierz wyniki",
        "Pobierz przefiltrowane dane albo pełny raport Excel bez opuszczania panelu roboczego.",
    )
    st.download_button(
        "Pobierz filtrowane dane CSV",
        data=csv_bytes,
        file_name="pjoter_development_release_change_filtered.csv",
        mime="text/csv",
        use_container_width=True,
    )
    st.download_button(
        "Pobierz raport Excel",
        data=excel_bytes,
        file_name="pjoter_development_release_change_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def render_welcome_state(prev_file, current_file):
    brand_context = detect_brand_context(
        *(meta for meta in [
            {"file_name": prev_file.name} if prev_file is not None else None,
            {"file_name": current_file.name} if current_file is not None else None,
        ] if meta)
    )
    has_any_file = prev_file is not None or current_file is not None
    title = (
        "Dodaj dwa pliki, aby uruchomic porownanie release'ow"
        if not has_any_file
        else "Dodaj drugi plik, aby dokonczyc analize"
    )
    subtitle = (
        "Lewa kolumna sluzy do uploadu, filtrow i eksportu. Po dodaniu kompletu plikow "
        "dashboard automatycznie pokaze KPI, alerty, wykresy i tabele szczegolowe."
    )
    meta_items = [
        "Upload po lewej stronie",
        "Porownanie daily i weekly",
        "Eksport CSV / Excel",
        "1 / 2 plikow gotowe" if has_any_file else "0 / 2 plikow gotowe",
    ]
    render_empty_state_header(brand_context, title, subtitle, meta_items)

    render_section_header(
        "Po uruchomieniu analizy",
        "Co zobaczysz w raporcie",
        "Po lewej stronie zostaje sterowanie analiza, a glowna sekcja skupia sie na wynikach i najwazniejszych sygnalach.",
    )
    quick_cols = st.columns(3, gap="medium")
    with quick_cols[0]:
        render_quick_card(
            "Szybkie KPI",
            "Najwazniejsze liczby i sygnaly beda zawsze na gorze raportu, gotowe do szybkiego odczytu.",
        )
    with quick_cols[1]:
        render_quick_card(
            "Alerty i insighty",
            "Sekcja alertow porzadkuje anomalie, nowe pozycje i zmiany przekraczajace ustalony prog.",
        )
    with quick_cols[2]:
        render_quick_card(
            "Staly kontekst pracy",
            "Filtry, upload i eksport pozostaja w jednym miejscu, dzieki czemu dashboard nie gubi kontekstu pracy.",
        )

    if prev_file is None and current_file is None:
        st.info("Dodaj dwa pliki Excel w panelu po lewej, aby uruchomic porownanie release'ow.")
    else:
        missing_label = "poprzedni" if prev_file is None else "aktualny"
        st.info(
            f"Jeden plik jest juz gotowy. Dodaj jeszcze plik {missing_label}, aby uruchomic pelna analize."
        )
    return
    title = "Premium dashboard porównawczy dla release'ów"
    subtitle = (
        "Po załadowaniu dwóch plików aplikacja zbuduje raport KPI, alerty, widoki tygodniowe, "
        "tabele szczegółowe i eksport gotowy do dalszej pracy operacyjnej."
    )
    meta_items = [
        "Upload po lewej stronie",
        "Porównanie daily i weekly",
        "Eksport CSV / Excel",
    ]
    if prev_file is not None or current_file is not None:
        meta_items.append("1 / 2 plików gotowe" if prev_file is None or current_file is None else "2 / 2 plików gotowe")
    render_app_header(brand_context, title, subtitle, meta_items)

    render_section_header(
        "Jak działa workspace",
        "Jeden panel do uploadu, filtrów i eksportu",
        "Lewa kolumna utrzymuje stały kontekst pracy, a prawa część skupia się wyłącznie na wynikach i analizie.",
    )
    quick_cols = st.columns(3, gap="medium")
    with quick_cols[0]:
        render_quick_card(
            "Szybkie KPI",
            "Najważniejsze liczby i sygnały są zawsze na górze raportu, gotowe do szybkiego odczytu.",
        )
    with quick_cols[1]:
        render_quick_card(
            "Insighty i alerty",
            "Sekcja alertów porządkuje anomalie, nowe pozycje i zmiany przekraczające ustalony próg.",
        )
    with quick_cols[2]:
        render_quick_card(
            "Stabilny kontekst analizy",
            "Filtry, upload i eksport pozostają w jednym miejscu, dzięki czemu dashboard nie gubi kontekstu pracy.",
        )

    if prev_file is None and current_file is None:
        st.info("Dodaj dwa pliki Excel w panelu po lewej, aby uruchomić porównanie release'ów.")
    else:
        missing_label = "poprzedni" if prev_file is None else "aktualny"
        st.info(
            f"Jeden plik jest już gotowy. Dodaj jeszcze plik {missing_label}, aby uruchomić pełną analizę."
        )


def build_detail_export_table(dataframe):
    detail_table = dataframe[available_detail_columns(dataframe)].copy()
    if "Ship Date" in detail_table.columns:
        detail_table["Ship Date"] = detail_table["Ship Date"].dt.strftime("%Y-%m-%d")
    if "Receipt Date" in detail_table.columns:
        detail_table["Receipt Date"] = detail_table["Receipt Date"].dt.strftime("%Y-%m-%d")
    if "Change Direction" in detail_table.columns:
        detail_table["Change Direction"] = detail_table["Change Direction"].map(
            get_change_label
        )
    if "Alert" in detail_table.columns:
        detail_table["Alert"] = detail_table["Alert"].map(
            lambda value: "Tak" if value else "Nie"
        )
    return detail_table.rename(
        columns={
            "PO Number": "Numer PO",
            "Origin Doc": "Origin Doc",
            "Item": "Pozycja",
            "Ship To": "Ship-to",
            "Part Number": "Numer części",
            "Part Description": "Opis produktu",
            "Customer Material": "Materiał klienta",
            "Unrestricted Qty": "Ilość unrestr.",
            "Unloading Point": "Punkt rozładunku",
            "Ship Date": "Data wysyłki",
            "Receipt Date": "Data odbioru",
            "Unit of Measure": "JM",
            "CumQty": "CumQty",
            "Quantity_Prev": "Poprzednia ilość",
            "Quantity_Curr": "Aktualna ilość",
            "Delta": "Zmiana ilości",
            "Percent Change": "Zmiana %",
            "Demand Status": "Status popytu",
            "Change Direction": "Kierunek zmiany",
            "Alert": "Alert",
        }
    )


def load_auth_config():
    if not AUTH_USERS_PATH.exists():
        # Default user for Streamlit Cloud deployment
        return [
            {
                "username": "admin",
                "display_name": "Administrator",
                "role": "Admin",
                "active": True,
                "salt": "c6b02c39a66d2460b6a3a3885b467ad0",
                "password_hash": "f951c24eead1d41496fc80c791f5ac802af477002998494b058dde362f1e2dda"
            }
        ]
    try:
        with AUTH_USERS_PATH.open("r", encoding="utf-8") as file:
            payload = json.load(file)
        return payload.get("users", [])
    except Exception:
        # Return default user if JSON is corrupted
        return [
            {
                "username": "admin",
                "display_name": "Administrator",
                "role": "Admin",
                "active": True,
                "salt": "c6b02c39a66d2460b6a3a3885b467ad0",
                "password_hash": "f951c24eead1d41496fc80c791f5ac802af477002998494b058dde362f1e2dda"
            }
        ]


def verify_password(password, salt_hex, password_hash_hex):
    salt = binascii.unhexlify(salt_hex.encode("utf-8"))
    computed_hash = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt, 120000
    )
    return binascii.hexlify(computed_hash).decode("utf-8") == password_hash_hex


def init_auth_state():
    st.session_state.setdefault("authenticated", False)
    st.session_state.setdefault("auth_user", None)


def attempt_login(username, password):
    users = load_auth_config()
    normalized_username = username.strip().lower()
    for user in users:
        if not user.get("active", True):
            continue
        if str(user.get("username", "")).strip().lower() != normalized_username:
            continue
        if verify_password(
            password,
            user.get("salt", ""),
            user.get("password_hash", ""),
        ):
            st.session_state["authenticated"] = True
            st.session_state["auth_user"] = {
                "username": user.get("username", ""),
                "display_name": user.get("display_name", user.get("username", "")),
                "role": user.get("role", "User"),
            }
            return True
    return False


def logout_user():
    st.session_state["authenticated"] = False
    st.session_state["auth_user"] = None


def _legacy_render_sidebar_user():
    auth_user = st.session_state.get("auth_user") or {}
    st.sidebar.markdown(
        f"""
        <div class="sidebar-user-card">
            <div class="sidebar-user-label">Aktywna sesja</div>
            <div class="sidebar-user-name">{auth_user.get('display_name', 'User')}</div>
            <div class="sidebar-user-role">{auth_user.get('role', 'User')} · {auth_user.get('username', '')}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.sidebar.button("Wyloguj", use_container_width=True):
        logout_user()
        st.rerun()


def render_filter_panel_shell():
    st.markdown(
        """
        <div class="filter-panel-shell">
            <div class="filter-panel-kicker">Panel Nawigacji</div>
            <div class="filter-panel-title">Filtry i kontekst analizy</div>
            <div class="filter-panel-copy">
                Ten panel pozostaje stale widoczny, aby filtrowanie, kalendarz i kontrola zakresu
                były zawsze pod ręką podczas pracy z dashboardem.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_login_screen():
    st.markdown('<div class="login-shell">', unsafe_allow_html=True)
    left_col, right_col = st.columns([1.18, 0.92], gap="large")
    with left_col:
        logo_html = (
            f'<img class="hero-logo" src="{logo_data_uri()}" alt="{BRAND_NAME} logo" />'
            if logo_available()
            else f'<div class="brand-badge">{BRAND_NAME}</div>'
        )
        st.markdown(
            f"""
            <div class="login-brand-card">
                {logo_html}
                <div class="login-kicker">Secure Access</div>
                <div class="login-title">Analiza zamówień i wysyłek</div>
                <div class="login-copy">
                    Zaloguj się, aby otworzyć dashboard, porównywać dwa pliki Excel i generować raporty
                    dla planowania, logistyki oraz przeglądów managerskich.
                </div>
                <div class="login-points">
                    <div class="login-point">
                        <div class="login-point-title">Premium dashboard analytics</div>
                        <div class="login-point-copy">Czytelne KPI, alerty, wykresy i macierz zmian z wyraźną hierarchią informacji.</div>
                    </div>
                    <div class="login-point">
                        <div class="login-point-title">Raport gotowy do eksportu</div>
                        <div class="login-point-copy">Filtrowane dane CSV i biznesowy raport Excel przygotowany do dalszej pracy.</div>
                    </div>
                    <div class="login-point">
                        <div class="login-point-title">Dostęp na wielu komputerach</div>
                        <div class="login-point-copy">Ta sama aplikacja może działać lokalnie, w sieci LAN lub jako uruchamiany launcher EXE.</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right_col:
        st.markdown(
            """
            <div class="login-form-card">
                <div class="login-form-heading">Zaloguj się do aplikacji</div>
                <div class="login-form-copy">
                    Użyj swojego loginu i hasła, aby otworzyć panel analityczny. Dane dostępowe są trzymane
                    lokalnie w konfiguracji aplikacji.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Login")
            password = st.text_input("Hasło", type="password")
            submitted = st.form_submit_button("Zaloguj", use_container_width=True)
        if submitted:
            if attempt_login(username, password):
                st.success("Logowanie zakończone powodzeniem.")
                st.rerun()
            st.error("Nieprawidłowy login lub hasło.")
        st.info(
            "Domyślny login jest zapisany w pliku config/users.json. Po wdrożeniu zmień hasło administratora."
        )
    st.markdown("</div>", unsafe_allow_html=True)


def logo_available():
    return LOGO_PATH.exists()


def asset_data_uri(path):
    try:
        if not path.exists():
            return ""
        extension = path.suffix.lower()
        if extension == ".svg":
            mime_type = "image/svg+xml"
        elif extension in {".jpg", ".jpeg"}:
            mime_type = "image/jpeg"
        elif extension == ".webp":
            mime_type = "image/webp"
        else:
            mime_type = "image/png"
        encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
        return f"data:{mime_type};base64,{encoded}"
    except Exception:
        return ""


def logo_data_uri():
    return asset_data_uri(LOGO_PATH)


def render_sidebar_user(target=st.sidebar):
    auth_user = st.session_state.get("auth_user") or {}
    target.markdown(
        f"""
        <div class="sidebar-user-card">
            <div class="sidebar-user-label">Aktywna sesja</div>
            <div class="sidebar-user-name">{html.escape(str(auth_user.get('display_name', 'User')))}</div>
            <div class="sidebar-user-role">{html.escape(str(auth_user.get('role', 'User')))} &middot; {html.escape(str(auth_user.get('username', '')))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if target.button("Wyloguj", use_container_width=True):
        logout_user()
        st.rerun()


def detect_brand_context(*metas):
    file_names = [
        str(meta.get("file_name", "")).lower()
        for meta in metas
        if isinstance(meta, dict)
    ]
    file_types = {
        str(meta.get("file_type", "")).strip().lower()
        for meta in metas
        if isinstance(meta, dict)
    }
    joined_names = " ".join(file_names)

    if "cw_weekly_pivot" in file_types or any(
        keyword in joined_names for keyword in ["audi", "q7", "q9", "megatech"]
    ):
        return {
            "brand_key": "audi",
            "label": "Audi Q7/Q9",
            "status": "Klient: Audi Q7/Q9",
            "format_copy": "Rozpoznano tygodniowy format CW / Audi Q7/Q9.",
            "banner_text": "AUDI Q7/Q9",
            "banner_theme": "audi",
        }

    if any(keyword in joined_names for keyword in ["mercedes", "merc", "vl10e"]) or "vl10e_block" in file_types:
        return {
            "brand_key": "mercedes",
            "label": "Mercedes-Benz",
            "status": "Klient: Mercedes-Benz",
            "format_copy": "Rozpoznano plik VL10E / Mercedes.",
            "banner_text": "MERCEDES-BENZ",
            "banner_theme": "mercedes",
        }

    if any(keyword in joined_names for keyword in ["releasedata", "tesla"]) or (
        "legacy_wide" in file_types and "releasedata" in joined_names
    ):
        return {
            "brand_key": "tesla",
            "label": "Tesla",
            "status": "Klient: Tesla / ReleaseData",
            "format_copy": "Rozpoznano plik ReleaseData / legacy wide.",
            "banner_text": "TESLA",
            "banner_theme": "tesla",
        }

    return {
        "brand_key": "default",
        "label": "Analytics Dashboard",
        "status": "Klient: neutralny",
        "format_copy": "Brak dedykowanej marki dla załadowanego pliku.",
        "banner_text": "ANALYTICS DASHBOARD",
        "banner_theme": "default",
    }


def describe_format_context(*metas):
    file_types = [
        str(meta.get("file_type", "")).strip().lower()
        for meta in metas
        if isinstance(meta, dict) and meta.get("file_type")
    ]
    if not file_types:
        return "Oczekiwanie na dwa pliki wejściowe."
    unique_types = sorted(set(file_types))
    if unique_types == ["cw_weekly_pivot"]:
        return "Format: Weekly pivot"
    if unique_types == ["vl10e_block"]:
        return "Format: VL10E block"
    if unique_types == ["legacy_wide"]:
        return "Format: Legacy wide"
    if set(unique_types) == {"cw_weekly_pivot", "legacy_wide"}:
        return "Format: daily + weekly"
    return "Format: mixed / mixed release sources"


def build_file_type_banner_markup(brand_context, variant="sidebar"):
    banner_text = html.escape(str(brand_context.get("banner_text", "ANALYTICS DASHBOARD")))
    banner_theme = html.escape(str(brand_context.get("banner_theme", "default")))
    banner_variant = "header" if variant == "header" else "sidebar"
    return (
        f'<div class="file-type-banner file-type-banner--{banner_variant} '
        f'file-type-banner--{banner_theme}">'
        f'<div class="file-type-banner__text">{banner_text}</div>'
        "</div>"
    )


def render_file_type_banner(brand_context, target=st, variant="sidebar"):
    target.markdown(
        build_file_type_banner_markup(brand_context, variant=variant),
        unsafe_allow_html=True,
    )


def render_side_panel_brand(brand_context):
    render_file_type_banner(brand_context, variant="sidebar")


def render_compact_header(brand_context, prev_meta, curr_meta, date_basis, selected_start_date, selected_end_date):
    format_context = describe_format_context(prev_meta, curr_meta)
    render_app_header(
        brand_context,
        f"Raport zmian dla PO {curr_meta.get('po_number', 'n/a')}",
        (
            f"{brand_context.get('status', 'Klient: neutralny')}. "
            f"{brand_context.get('format_copy', '')} "
            f"Zakres analizy: {selected_start_date:%Y-%m-%d} do {selected_end_date:%Y-%m-%d} "
            f"na osi {get_date_label(date_basis)}."
        ),
        [
            format_context,
            f"Poprzedni: {format_release_label(prev_meta)}",
            f"Aktualny: {format_release_label(curr_meta)}",
        ],
        curr_meta.get("file_name", ""),
    )


def apply_chart_theme(chart):
    return (
        chart.configure_view(strokeOpacity=0)
        .configure(background="transparent")
        .configure_axis(
            grid=False,
            domainColor="#334155",
            tickColor="#64748b",
            labelColor="#94a3b8",
            titleColor="#e2e8f0",
            labelFontSize=12,
            titleFontSize=13,
            tickSize=6,
            labelPadding=10,
            titlePadding=12,
        )
        .configure_axisX(
            grid=True,
            gridColor="rgba(148, 163, 184, 0.12)",
            gridDash=[2, 6],
            domain=False,
            tickColor="#64748b",
            labelColor="#94a3b8",
        )
        .configure_axisY(
            grid=True,
            gridColor="rgba(148, 163, 184, 0.12)",
            gridDash=[2, 6],
            domain=False,
            tickColor="#64748b",
            labelColor="#94a3b8",
        )
        .configure_legend(
            labelColor="#cbd5e1",
            titleColor="#e2e8f0",
            labelFontSize=12,
            titleFontSize=13,
            symbolType="circle",
        )
        .configure_title(color="#f8fafc", fontSize=16, fontWeight="bold", anchor="start")
    )


def normalize_date_selection(selection, default_start, default_end):
    if isinstance(selection, tuple):
        values = list(selection)
    elif isinstance(selection, list):
        values = selection
    else:
        values = [selection]

    if len(values) == 0:
        return default_start, default_end
    if len(values) == 1:
        return values[0], values[0]
    return values[0], values[1]


def render_filter_panel_shell(
    kicker="Panel Nawigacji",
    title="Filtry i kontekst analizy",
    copy=(
        "Ten panel pozostaje stale widoczny, aby filtrowanie, kalendarz i kontrola zakresu "
        "byly zawsze pod reka podczas pracy z dashboardem."
    ),
):
    st.markdown(
        f"""
        <div class="filter-panel-shell">
            <div class="filter-panel-kicker">{html.escape(str(kicker))}</div>
            <div class="filter-panel-title">{html.escape(str(title))}</div>
            <div class="filter-panel-copy">
                {html.escape(str(copy))}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_filter_controls(result):
    render_section_header(
        "Filtry",
        "Zakres i oś analizy",
        "Utrzymuj ten sam kontekst pracy podczas przeglądania wszystkich zakładek raportu.",
    )
    date_basis = st.segmented_control(
        "Oś dat",
        DATE_OPTIONS,
        selection_mode="single",
        default=DATE_OPTIONS[0],
        required=True,
        key="analysis_date_basis",
        format_func=get_date_label,
        width="stretch",
    )
    date_basis = date_basis or DATE_OPTIONS[0]

    available_dates = result[date_basis].dropna().sort_values()
    min_date = available_dates.min().date()
    max_date = available_dates.max().date()

    st.markdown("###### Zakres czasowy")
    selected_date_input = st.date_input(
        "Wybierz przedział dat:",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        help="Kliknij, aby wybrać pojedynczy dzień lub zakres dat do analizy.",
        label_visibility="collapsed",
    )
    selected_start_date, selected_end_date = normalize_date_selection(
        selected_date_input, min_date, max_date
    )
    swapped_dates = selected_start_date > selected_end_date
    if swapped_dates:
        selected_start_date, selected_end_date = selected_end_date, selected_start_date
        st.warning("Zamieniono kolejność dat, aby zachować poprawny zakres analizy.")

    st.caption(
        f"Zakres aktywnej analizy: {selected_start_date.strftime('%Y-%m-%d')} — {selected_end_date.strftime('%Y-%m-%d')}"
    )

    full_product_summary = summarize_products(result)
    all_products = full_product_summary["Product Label"].tolist()
    st.markdown("###### Zakres produktów")
    selected_products = st.multiselect(
        "Produkty",
        options=all_products,
        default=all_products,
    )
    search_term = st.text_input("Szukaj po numerze lub opisie")
    st.markdown("###### Kierunek zmiany")
    selected_change_directions = st.segmented_control(
        "Kierunek zmiany",
        options=["Increase", "Decrease", "No Change"],
        selection_mode="multi",
        default=["Increase", "Decrease", "No Change"],
        key="analysis_change_direction",
        format_func=get_change_label,
        width="stretch",
    )
    selected_change_directions = selected_change_directions or ["Increase", "Decrease", "No Change"]
    only_alerts = st.checkbox(f"Tylko alerty >= {THRESHOLD}%")

    return {
        "date_basis": date_basis,
        "selected_start_date": selected_start_date,
        "selected_end_date": selected_end_date,
        "selected_products": selected_products,
        "search_term": search_term,
        "selected_change_directions": selected_change_directions,
        "only_alerts": only_alerts,
        "full_product_summary": full_product_summary,
    }


def render_welcome_side_panel(prev_file, current_file):
    pending_meta = []
    if prev_file is not None:
        pending_meta.append({"file_name": prev_file.name})
    if current_file is not None:
        pending_meta.append({"file_name": current_file.name})
    brand_context = detect_brand_context(*pending_meta) if pending_meta else detect_brand_context()

    render_filter_panel_shell(
        kicker="Workspace",
        title="Panel aplikacji",
        copy=(
            "Upload, status plikow i pozniejszy panel filtrow pozostaja w jednej stalej kolumnie roboczej."
        ),
    )
    render_side_panel_brand(brand_context)
    render_file_slot_cards(prev_file=prev_file, current_file=current_file)
    st.info(
        "Po zaladowaniu obu plikow ten panel zostanie uzupelniony o filtry, status formatu i akcje eksportu."
    )


def render_analysis_side_panel(result, brand_context, prev_meta=None, curr_meta=None):
    render_filter_panel_shell(
        kicker="Analysis Controls",
        title="Filtry i status analizy",
        copy="Lewa kolumna utrzymuje upload, status plikow, filtry oraz eksport w jednym miejscu pracy.",
    )
    render_side_panel_brand(brand_context)
    st.caption(brand_context.get("format_copy", ""))
    render_file_slot_cards(prev_meta=prev_meta, curr_meta=curr_meta)
    return render_filter_controls(result)


def render_analysis_main(
    filtered_df,
    product_summary,
    date_summary,
    weekly_summary,
    key_findings,
    prev_meta,
    curr_meta,
    date_basis,
    selected_start_date,
    selected_end_date,
    excel_bytes=None,
    csv_bytes=None,
):
    brand_context = detect_brand_context(prev_meta, curr_meta)
    render_compact_header(
        brand_context,
        prev_meta,
        curr_meta,
        date_basis,
        selected_start_date,
        selected_end_date,
    )

    reference_week = get_last_completed_reference_week(selected_end_date)
    reference_row, previous_week_row = get_reference_week_rows(weekly_summary)
    reference_week_label = (
        reference_row["Week Label"] if reference_row is not None else reference_week.week_label
    )
    reference_range_label = (
        format_week_range(reference_row["Week Start"], reference_row["Week End"])
        if reference_row is not None
        else format_week_range(reference_week.week_start, reference_week.week_end)
    )
    reference_release_delta = (
        format_signed_int(reference_row["Delta"]) if reference_row is not None else "+0"
    )
    reference_release_pct = (
        format_percent_display(reference_row["Release Percent Label"])
        if reference_row is not None
        else "n/a"
    )
    reference_wow_delta = (
        format_signed_int(reference_row["WoW Delta"]) if reference_row is not None else "+0"
    )
    reference_wow_pct = (
        format_percent_display(reference_row["WoW Percent Label"])
        if reference_row is not None
        else "n/a"
    )
    reference_working_days = (
        int(reference_row["Working_Days_PL"]) if reference_row is not None else 0
    )
    reference_per_day = (
        "n/a"
        if reference_row is None or pd.isna(reference_row["Avg Current / Working Day"])
        else f"{float(reference_row['Avg Current / Working Day']):,.2f} / dzien"
    )
    previous_week_label = (
        previous_week_row["Week Label"] if previous_week_row is not None else "brak"
    )

    report_metadata = [
        {"label": "Klient", "value": brand_context.get("label", "n/a")},
        {"label": "Format", "value": describe_format_context(prev_meta, curr_meta)},
        {"label": "Numer PO", "value": curr_meta.get("po_number", "n/a")},
        {"label": "Planista", "value": curr_meta.get("planner_name", "n/a")},
        {"label": "E-mail", "value": curr_meta.get("planner_email", "n/a")},
        {"label": "Oś analizy", "value": get_date_label(date_basis)},
        {
            "label": "Zakres analizy",
            "value": f"{selected_start_date:%Y-%m-%d} — {selected_end_date:%Y-%m-%d}",
        },
        {"label": "Referencyjny tydzień", "value": reference_week_label},
        {"label": "Poprzedni release", "value": format_release_summary(prev_meta)},
        {"label": "Aktualny release", "value": format_release_summary(curr_meta)},
    ]
    render_report_metadata(report_metadata)

    if filtered_df.empty:
        st.warning(
            "Po zastosowaniu filtrów nie ma danych do pokazania. Poszerz zakres dat albo przywróć produkty w panelu po lewej stronie."
        )
        return

    render_section_header(
        "KPI",
        "Najważniejsze wskaźniki",
        "Karty poniżej pokazują główne liczby do szybkiego odczytu bez przeskakiwania między zakładkami.",
    )
    render_kpi_cards(build_kpi_metrics(filtered_df, product_summary))

    render_section_header(
        "Alerts & Insights",
        "Priorytety do sprawdzenia",
        "Najważniejsze sygnały, które warto zweryfikować w pierwszej kolejności.",
    )
    render_alerts(build_alert_items(filtered_df, key_findings))

    render_section_header(
        "Reference Week",
        "Szybki odczyt tygodniowy",
        (
            f"Analiza tygodniowa odnosi się do {reference_week_label} ({reference_range_label}). "
            f"Data referencyjna: {selected_end_date:%Y-%m-%d}."
        ),
    )
    render_kpi_cards(
        [
            {
                "label": "Wolumen tygodnia",
                "value": f"{float(reference_row['Quantity_Curr']):,.0f}" if reference_row is not None else "0",
                "copy": f"Bilans release: {reference_release_delta}",
                "tone": "neutral",
            },
            {
                "label": "Zmiana vs poprzedni release",
                "value": reference_release_pct,
                "copy": f"Poprzedni wolumen: {float(reference_row['Quantity_Prev']):,.0f}" if reference_row is not None else "Poprzedni wolumen: 0",
                "tone": "neutral",
            },
            {
                "label": "Zmiana WoW",
                "value": reference_wow_delta,
                "copy": f"{reference_wow_pct} względem {previous_week_label}",
                "tone": "neutral",
            },
            {
                "label": "Dni robocze PL",
                "value": f"{reference_working_days}",
                "copy": reference_per_day,
                "tone": "neutral",
            },
        ]
    )

    dashboard_tab, weekly_tab, product_tab, matrix_tab, detail_tab = st.tabs(
        ["Dashboard", "Analiza tygodniowa", "Raport produktu", "Macierz release'u", "Dane szczegolowe"]
    )

    with dashboard_tab:
        render_section_header(
            "Dashboard",
            f"Trend zmian według osi: {get_date_label(date_basis)}",
            "Widok główny zbiera najważniejsze wykresy, strukturę zmian oraz szybki podgląd produktów z największym ruchem.",
        )
        render_chart_table_switch(
            "dashboard_trend",
            build_quantity_chart(date_summary, get_date_label(date_basis)),
            date_summary,
            table_height=360,
        )

        trend_left, trend_right = st.columns([1.45, 1], gap="large")
        with trend_left:
            render_chart_table_switch(
                "dashboard_delta",
                build_delta_chart(date_summary, get_date_label(date_basis)),
                date_summary,
                table_height=320,
            )
        with trend_right:
            st.subheader("Struktura zmian")
            render_chart_table_switch(
                "dashboard_mix",
                build_change_mix_chart(filtered_df),
                build_change_mix_source(filtered_df),
                table_height=240,
            )

        increase_chart, increase_title = build_product_bar_chart(product_summary, "increase")
        decrease_chart, decrease_title = build_product_bar_chart(product_summary, "decrease")
        dashboard_left, dashboard_right = st.columns(2)

        with dashboard_left:
            st.subheader(increase_title)
            if increase_chart is None:
                st.info("Brak produktow ze wzrostem w aktualnym filtrowaniu.")
            else:
                render_chart_table_switch(
                    "dashboard_increase",
                    increase_chart,
                    build_product_bar_source(product_summary, "increase"),
                    table_height=340,
                )

        with dashboard_right:
            st.subheader(decrease_title)
            if decrease_chart is None:
                st.info("Brak produktow ze spadkiem w aktualnym filtrowaniu.")
            else:
                render_chart_table_switch(
                    "dashboard_decrease",
                    decrease_chart,
                    build_product_bar_source(product_summary, "decrease"),
                    table_height=340,
                )

        st.subheader("Najwazniejsze zmiany")
        highlight_table = (
            product_summary.assign(Abs_Delta=product_summary["Delta"].abs())
            .sort_values("Abs_Delta", ascending=False)
            .drop(columns=["Abs_Delta"])
            .head(10)
        )
        highlight_table["Quantity_Prev"] = highlight_table["Quantity_Prev"].map(
            lambda value: f"{value:,.0f}"
        )
        highlight_table["Quantity_Curr"] = highlight_table["Quantity_Curr"].map(
            lambda value: f"{value:,.0f}"
        )
        highlight_table["Delta"] = highlight_table["Delta"].map(format_signed_int)
        highlight_table = highlight_table.rename(
            columns={
                "Part Number": "Numer czesci",
                "Part Description": "Opis produktu",
                "Quantity_Prev": "Poprzednia ilosc",
                "Quantity_Curr": "Aktualna ilosc",
                "Delta": "Zmiana ilosci",
                "Alert_Count": "Liczba alertow",
                "Change Direction": "Kierunek zmiany",
            }
        )
        st.dataframe(highlight_table, use_container_width=True, height=360)

        st.subheader("Tygodnie ISO")
        weekly_chart = build_weekly_quantity_chart(weekly_summary)
        weekly_preview = prepare_weekly_display_table(weekly_summary).tail(8)
        render_chart_table_switch(
            "dashboard_weekly",
            weekly_chart,
            weekly_preview,
            chart_empty_message="Brak danych tygodniowych do wykresu.",
            table_height=320,
        )

    with weekly_tab:
        render_section_header(
            "Weekly View",
            "Analiza tygodniowa oparta na datach",
            "Ten widok agreguje dzienne dane do poziomu tygodni ISO i ułatwia porównanie release-over-release oraz week-over-week.",
        )
        weekly_partial = weekly_summary[
            weekly_summary["Is Partial Range"] | ~weekly_summary["Is Closed Week"]
        ]
        if not weekly_partial.empty:
            st.info(
                "W tabeli i wykresach tygodnie oznaczone jako 'Partial range' lub 'Open week' "
                "obejmuja niepelny zakres albo nie byly jeszcze zakonczone wzgledem daty referencyjnej."
            )

        weekly_qty_chart = build_weekly_quantity_chart(weekly_summary)
        render_chart_table_switch(
            "weekly_quantity",
            weekly_qty_chart,
            prepare_weekly_display_table(weekly_summary),
            chart_empty_message="Brak danych tygodniowych do wykresu.",
            table_height=360,
        )

        weekly_left, weekly_right = st.columns([1.3, 1], gap="large")
        with weekly_left:
            weekly_delta_chart = build_weekly_delta_chart(weekly_summary)
            render_chart_table_switch(
                "weekly_delta",
                weekly_delta_chart,
                prepare_weekly_display_table(weekly_summary),
                chart_empty_message="Brak danych tygodniowych do wykresu delta.",
                table_height=320,
            )
        with weekly_right:
            weekly_focus = pd.DataFrame(
                [
                    {
                        "Widok": "Referencyjny tydzien",
                        "Tydzien ISO": reference_week_label,
                        "Aktualny release": (
                            f"{float(reference_row['Quantity_Curr']):,.0f}"
                            if reference_row is not None
                            else "0"
                        ),
                        "Poprzedni release": (
                            f"{float(reference_row['Quantity_Prev']):,.0f}"
                            if reference_row is not None
                            else "0"
                        ),
                        "Delta release": reference_release_delta,
                        "Zmiana release %": reference_release_pct,
                        "Delta WoW": reference_wow_delta,
                        "Zmiana WoW %": reference_wow_pct,
                    },
                    {
                        "Widok": "Poprzedni tydzien",
                        "Tydzien ISO": previous_week_label,
                        "Aktualny release": (
                            f"{float(previous_week_row['Quantity_Curr']):,.0f}"
                            if previous_week_row is not None
                            else "0"
                        ),
                        "Poprzedni release": (
                            f"{float(previous_week_row['Quantity_Prev']):,.0f}"
                            if previous_week_row is not None
                            else "0"
                        ),
                        "Delta release": (
                            format_signed_int(previous_week_row["Delta"])
                            if previous_week_row is not None
                            else "+0"
                        ),
                        "Zmiana release %": (
                            format_percent_display(previous_week_row["Release Percent Label"])
                            if previous_week_row is not None
                            else "n/a"
                        ),
                        "Delta WoW": (
                            format_signed_int(previous_week_row["WoW Delta"])
                            if previous_week_row is not None
                            else "+0"
                        ),
                        "Zmiana WoW %": (
                            format_percent_display(previous_week_row["WoW Percent Label"])
                            if previous_week_row is not None
                            else "n/a"
                        ),
                    },
                ]
            )
            st.subheader("Porownanie tygodni")
            st.dataframe(weekly_focus, use_container_width=True, height=240)

        weekly_table = prepare_weekly_display_table(weekly_summary)
        st.subheader("Tabela tygodniowa")
        st.dataframe(weekly_table, use_container_width=True, height=420)

    with product_tab:
        render_section_header(
            "Product Drilldown",
            "Analiza wybranego produktu",
            "Skup się na jednym materiale i prześledź jego ruch po dniach oraz tygodniach bez utraty kontekstu filtrowania.",
        )
        selected_product_label = st.selectbox(
            "Wybierz produkt",
            options=product_summary["Product Label"].tolist(),
        )
        product_detail = filtered_df[
            filtered_df["Product Label"] == selected_product_label
        ].sort_values(date_basis)
        product_date_summary = summarize_dates(product_detail, date_basis)

        product_metrics = st.columns(4)
        product_metrics[0].metric(
            "Poprzednia ilosc", f"{product_detail['Quantity_Prev'].sum():,.0f}"
        )
        product_metrics[1].metric(
            "Aktualna ilosc", f"{product_detail['Quantity_Curr'].sum():,.0f}"
        )
        product_metrics[2].metric(
            "Bilans zmian", f"{product_detail['Delta'].sum():+,.0f}"
        )
        product_metrics[3].metric("Liczba alertow", int(product_detail["Alert"].sum()))

        render_chart_table_switch(
            "product_quantity",
            build_quantity_chart(product_date_summary, get_date_label(date_basis)),
            product_date_summary,
            table_height=320,
        )
        render_chart_table_switch(
            "product_delta",
            build_delta_chart(product_date_summary, get_date_label(date_basis)),
            product_date_summary,
            table_height=320,
        )

        product_weekly_summary = build_weekly_summary(
            product_detail,
            date_basis,
            selected_start_date,
            selected_end_date,
            selected_end_date,
            THRESHOLD,
        )
        st.subheader("Tygodnie ISO dla produktu")
        product_weekly_chart = build_weekly_quantity_chart(product_weekly_summary)
        render_chart_table_switch(
            "product_weekly",
            product_weekly_chart,
            prepare_weekly_display_table(product_weekly_summary),
            chart_empty_message="Brak danych tygodniowych dla wybranego produktu.",
            table_height=280,
        )

        product_table = product_detail[available_detail_columns(product_detail)].copy()
        product_table["Ship Date"] = product_table["Ship Date"].dt.strftime("%Y-%m-%d")
        product_table["Receipt Date"] = product_table["Receipt Date"].dt.strftime("%Y-%m-%d")
        product_table["Change Direction"] = product_table["Change Direction"].map(
            get_change_label
        )
        product_table["Alert"] = product_table["Alert"].map(
            lambda value: "Tak" if value else "Nie"
        )
        product_table = product_table.rename(
            columns={
                "Part Number": "Numer czesci",
                "Part Description": "Opis produktu",
                "Origin Doc": "Origin Doc",
                "Item": "Pozycja",
                "Ship To": "Ship-to",
                "Customer Material": "Material klienta",
                "Unrestricted Qty": "Ilosc unrestr.",
                "Unloading Point": "Punkt rozladunku",
                "Ship Date": "Data wysylki",
                "Receipt Date": "Data odbioru",
                "Unit of Measure": "JM",
                "CumQty": "CumQty",
                "Quantity_Prev": "Poprzednia ilosc",
                "Quantity_Curr": "Aktualna ilosc",
                "Delta": "Zmiana ilosci",
                "Percent Change": "Zmiana %",
                "Demand Status": "Status popytu",
                "Change Direction": "Kierunek zmiany",
                "Alert": "Alert",
            }
        )
        st.dataframe(product_table, use_container_width=True, height=360)

    with matrix_tab:
        render_section_header(
            "Release Matrix",
            "Macierz podobna do arkusza release'u",
            "Macierz zachowuje układ bliski pracy w Excelu, ale pozostaje spójna wizualnie z całym dashboardem.",
        )
        matrix_metric = st.segmented_control(
            "Metryka",
            options=["Current Quantity", "Previous Quantity", "Delta", "Percent Change"],
            selection_mode="single",
            default="Current Quantity",
            required=True,
            format_func=get_metric_label,
            width="stretch",
        )
        matrix_metric = matrix_metric or "Current Quantity"
        matrix = build_matrix(filtered_df, date_basis, matrix_metric)
        matrix_cells = matrix.shape[0] * max(matrix.shape[1], 1)

        if matrix.empty:
            st.info("Brak danych do macierzy.")
        elif matrix_cells <= MAX_MATRIX_STYLE_CELLS:
            st.dataframe(
                style_matrix(matrix, matrix_metric),
                use_container_width=True,
                height=520,
            )
        else:
            st.info(
                "Macierz jest zbyt duza do stylowania, dlatego pokazuje ja bez dodatkowego formatowania."
            )
            st.dataframe(matrix, use_container_width=True, height=520)

    with detail_tab:
        render_section_header(
            "Detailed Data",
            "Dane szczegolowe",
            "Pełny podgląd przefiltrowanych wierszy do szybkiej walidacji oraz eksportu do dalszej pracy operacyjnej.",
        )
        preview_limit = st.selectbox(
            "Liczba wierszy w podgladzie",
            options=[100, 250, 500, 1000],
            index=2,
        )
        detail_table = build_detail_export_table(filtered_df)

        if len(detail_table) > preview_limit:
            st.info(
                f"Pokazuje pierwsze {preview_limit} z {len(detail_table)} wierszy. "
                "Pelny raport jest dostepny do pobrania."
            )
        st.dataframe(
            detail_table.head(preview_limit),
            use_container_width=True,
            height=420,
        )

        if excel_bytes is None:
            current_matrix_for_export = build_matrix(filtered_df, date_basis, "Current Quantity")
            delta_matrix_for_export = build_matrix(filtered_df, date_basis, "Delta")
            excel_bytes = to_excel_bytes(
                filtered_df,
                weekly_summary,
                current_matrix_for_export,
                delta_matrix_for_export,
                prev_meta,
                curr_meta,
                product_summary,
                date_basis,
                selected_start_date,
                selected_end_date,
                key_findings,
            )
        if csv_bytes is None:
            csv_bytes = detail_table.to_csv(index=False).encode("utf-8")

        download_left, download_right = st.columns(2)
        with download_left:
            st.download_button(
                "Pobierz filtrowane dane CSV",
                data=csv_bytes,
                file_name="pjoter_development_release_change_filtered.csv",
                mime="text/csv",
            )
        with download_right:
            st.download_button(
                "Pobierz raport Excel",
                data=excel_bytes,
                file_name="pjoter_development_release_change_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


@st.cache_data(show_spinner=False)
def load_release(file_bytes, file_name):
    return load_release_file(file_bytes, file_name)


def compare_releases(prev_df, curr_df):
    return compare_release_frames(prev_df, curr_df, threshold=THRESHOLD)


def summarize_products(dataframe):
    product_summary = (
        dataframe.groupby(["Part Number", "Part Description", "Product Label"], as_index=False)
        .agg(
            Quantity_Prev=("Quantity_Prev", "sum"),
            Quantity_Curr=("Quantity_Curr", "sum"),
            Delta=("Delta", "sum"),
            Abs_Delta=("Abs Delta", "sum"),
            Alert_Count=("Alert", "sum"),
        )
        .sort_values("Delta", ascending=False)
    )
    product_summary["Change Direction"] = product_summary["Delta"].apply(
        lambda value: "Increase" if value > 0 else ("Decrease" if value < 0 else "No Change")
    )
    return product_summary.reset_index(drop=True)


def summarize_dates(dataframe, date_basis):
    if dataframe.empty:
        return pd.DataFrame(
            columns=["Analysis Date", "Quantity_Prev", "Quantity_Curr", "Delta", "Alerts"]
        )

    date_summary = (
        dataframe.groupby(date_basis, as_index=False)
        .agg(
            Quantity_Prev=("Quantity_Prev", "sum"),
            Quantity_Curr=("Quantity_Curr", "sum"),
            Delta=("Delta", "sum"),
            Alerts=("Alert", "sum"),
        )
        .sort_values(date_basis)
        .rename(columns={date_basis: "Analysis Date"})
    )
    return date_summary


def format_percent_display(value):
    return "Nowy tydzien" if value == "new" else value


def format_week_range(start_value, end_value):
    start_label = pd.Timestamp(start_value).strftime("%Y-%m-%d")
    end_label = pd.Timestamp(end_value).strftime("%Y-%m-%d")
    return f"{start_label} - {end_label}"


def get_reference_week_rows(weekly_summary):
    if weekly_summary.empty:
        return None, None

    reference_rows = weekly_summary[weekly_summary["Is Reference Week"]]
    if reference_rows.empty:
        closed_rows = weekly_summary[weekly_summary["Is Closed Week"]]
        if closed_rows.empty:
            return None, None
        reference_rows = closed_rows.tail(1)

    reference_row = reference_rows.iloc[0]
    previous_rows = weekly_summary[weekly_summary["Week Start"] < reference_row["Week Start"]]
    previous_row = previous_rows.tail(1).iloc[0] if not previous_rows.empty else None
    return reference_row, previous_row


def prepare_weekly_display_table(weekly_summary):
    if weekly_summary.empty:
        return pd.DataFrame()

    weekly_table = weekly_summary[
        [
            "Week Label",
            "Week Start",
            "Week End",
            "Week Status",
            "Working_Days_PL",
            "Products",
            "Quantity_Prev",
            "Quantity_Curr",
            "Delta",
            "Release Percent Label",
            "WoW Delta",
            "WoW Percent Label",
            "Avg Current / Working Day",
            "Any Weekly Alert",
        ]
    ].copy()
    weekly_table["Week Range"] = weekly_table.apply(
        lambda row: format_week_range(row["Week Start"], row["Week End"]),
        axis=1,
    )
    weekly_table["Release Delta"] = weekly_table["Delta"].map(format_signed_int)
    weekly_table["WoW Delta"] = weekly_table["WoW Delta"].map(format_signed_int)
    weekly_table["Previous Release"] = weekly_table["Quantity_Prev"].map(lambda value: f"{value:,.0f}")
    weekly_table["Current Release"] = weekly_table["Quantity_Curr"].map(lambda value: f"{value:,.0f}")
    weekly_table["Release Change %"] = weekly_table["Release Percent Label"].map(format_percent_display)
    weekly_table["WoW Change %"] = weekly_table["WoW Percent Label"].map(format_percent_display)
    weekly_table["Current / Working Day"] = weekly_table["Avg Current / Working Day"].map(
        lambda value: "n/a" if pd.isna(value) else f"{float(value):,.2f}"
    )
    weekly_table["Alert"] = weekly_table["Any Weekly Alert"].map(lambda value: "Tak" if value else "Nie")
    return weekly_table[
        [
            "Week Label",
            "Week Range",
            "Week Status",
            "Working_Days_PL",
            "Products",
            "Previous Release",
            "Current Release",
            "Release Delta",
            "Release Change %",
            "WoW Delta",
            "WoW Change %",
            "Current / Working Day",
            "Alert",
        ]
    ].rename(
        columns={
            "Week Label": "Tydzien ISO",
            "Week Range": "Zakres tygodnia",
            "Week Status": "Status",
            "Working_Days_PL": "Dni robocze PL",
            "Products": "Produkty",
        }
    )


def build_weekly_quantity_chart(weekly_summary):
    if weekly_summary.empty:
        return None

    chart_data = weekly_summary.copy()
    chart_data["Week Start"] = pd.to_datetime(chart_data["Week Start"])
    current_area = (
        alt.Chart(chart_data)
        .mark_area(color="#5092ff", opacity=0.18, interpolate="monotone")
        .encode(
            x=alt.X("Week Start:T", title="Tydzien ISO", axis=alt.Axis(labelAngle=-24, labelLimit=120)),
            y=alt.Y("Quantity_Curr:Q", title="Wolumen tygodniowy"),
        )
    )
    prev_line = (
        alt.Chart(chart_data)
        .mark_line(strokeWidth=2.4, interpolate="monotone", color="#7c93c9", opacity=0.9)
        .encode(
            x=alt.X("Week Start:T", title="Tydzien ISO", axis=alt.Axis(labelAngle=-24, labelLimit=120)),
            y=alt.Y("Quantity_Prev:Q", title="Wolumen tygodniowy"),
            tooltip=[
                alt.Tooltip("Week Label:N", title="Tydzien"),
                alt.Tooltip("Quantity_Prev:Q", title="Poprzedni release", format=",.0f"),
                alt.Tooltip("Quantity_Curr:Q", title="Aktualny release", format=",.0f"),
                alt.Tooltip("Delta:Q", title="Delta release", format=",.0f"),
                alt.Tooltip("Working_Days_PL:Q", title="Dni robocze PL"),
                alt.Tooltip("Week Status:N", title="Status"),
            ],
        )
    )
    current_line = (
        alt.Chart(chart_data)
        .mark_line(
            point=alt.OverlayMarkDef(size=90, filled=True, fill="#eef4ff", stroke="#3c78d8", strokeWidth=2.2),
            strokeWidth=3.4,
            interpolate="monotone",
            color="#6cb0ff",
        )
        .encode(
            x=alt.X("Week Start:T", title="Tydzien ISO", axis=alt.Axis(labelAngle=-24, labelLimit=120)),
            y=alt.Y("Quantity_Curr:Q", title="Wolumen tygodniowy"),
            tooltip=[
                alt.Tooltip("Week Label:N", title="Tydzien"),
                alt.Tooltip("Quantity_Prev:Q", title="Poprzedni release", format=",.0f"),
                alt.Tooltip("Quantity_Curr:Q", title="Aktualny release", format=",.0f"),
                alt.Tooltip("Delta:Q", title="Delta release", format=",.0f"),
                alt.Tooltip("Avg Current / Working Day:Q", title="Na dzien roboczy", format=",.2f"),
                alt.Tooltip("Week Status:N", title="Status"),
            ],
        )
    )
    return apply_chart_theme(
        alt.layer(current_area, prev_line, current_line).properties(height=360)
    )


def build_weekly_delta_chart(weekly_summary):
    if weekly_summary.empty:
        return None

    chart_data = weekly_summary.copy()
    chart_data["Week Start"] = pd.to_datetime(chart_data["Week Start"])
    chart_data["WoW Color"] = chart_data["WoW Delta"].apply(
        lambda value: "#43d3b1" if value > 0 else "#ff6f6f" if value < 0 else "#7c93c9"
    )
    bars = (
        alt.Chart(chart_data)
        .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6, opacity=0.88)
        .encode(
            x=alt.X("Week Start:T", title="Tydzien ISO", axis=alt.Axis(labelAngle=-24, labelLimit=120)),
            y=alt.Y("WoW Delta:Q", title="Delta WoW"),
            color=alt.Color("WoW Color:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("Week Label:N", title="Tydzien"),
                alt.Tooltip("WoW Delta:Q", title="Zmiana vs poprzedni tydzien", format=",.0f"),
                alt.Tooltip("WoW Percent Label:N", title="Zmiana WoW %"),
                alt.Tooltip("Working_Days_PL:Q", title="Dni robocze PL"),
                alt.Tooltip("Week Status:N", title="Status"),
            ],
        )
    )
    line = (
        alt.Chart(chart_data)
        .mark_line(color="#d7e7ff", strokeWidth=2, opacity=0.55)
        .encode(
            x=alt.X("Week Start:T", title="Tydzien ISO", axis=alt.Axis(labelAngle=-24, labelLimit=120)),
            y=alt.Y("Delta:Q", title="Delta tygodniowa"),
        )
    )
    return apply_chart_theme(alt.layer(bars, line).properties(height=320))


def build_quantity_chart(date_summary, x_title):
    chart_data = date_summary.sort_values("Analysis Date").copy()
    latest_point = chart_data.tail(1).copy()
    latest_point["Current Label"] = latest_point["Quantity_Curr"].map(lambda value: f"Aktualnie {value:,.0f}")
    latest_point["Previous Label"] = latest_point["Quantity_Prev"].map(lambda value: f"Poprzednio {value:,.0f}")

    prev_line = (
        alt.Chart(chart_data)
        .mark_line(strokeWidth=2.6, interpolate="monotone", color="#6f8ed1", opacity=0.9)
        .encode(
            x=alt.X("Analysis Date:T", title=x_title, axis=alt.Axis(labelAngle=-24, labelLimit=140)),
            y=alt.Y("Quantity_Prev:Q", title="Ilość otwarta"),
            tooltip=[
                alt.Tooltip("Analysis Date:T", title="Data"),
                alt.Tooltip("Quantity_Prev:Q", title="Poprzednia ilość", format=",.0f"),
            ],
        )
    )
    current_area = (
        alt.Chart(chart_data)
        .mark_area(color="#5092ff", opacity=0.18, interpolate="monotone")
        .encode(
            x=alt.X("Analysis Date:T", title=x_title, axis=alt.Axis(labelAngle=-24, labelLimit=140)),
            y=alt.Y("Quantity_Curr:Q", title="Ilość otwarta"),
        )
    )
    current_line = (
        alt.Chart(chart_data)
        .mark_line(
            point=alt.OverlayMarkDef(size=90, filled=True, fill="#eef4ff", stroke="#3c78d8", strokeWidth=2.2),
            strokeWidth=3.6,
            interpolate="monotone",
            color="#6cb0ff",
        )
        .encode(
            x=alt.X("Analysis Date:T", title=x_title, axis=alt.Axis(labelAngle=-24, labelLimit=140)),
            y=alt.Y("Quantity_Curr:Q", title="Ilość otwarta"),
            tooltip=[
                alt.Tooltip("Analysis Date:T", title="Data"),
                alt.Tooltip("Quantity_Prev:Q", title="Poprzednia ilość", format=",.0f"),
                alt.Tooltip("Quantity_Curr:Q", title="Aktualna ilość", format=",.0f"),
                alt.Tooltip("Delta:Q", title="Bilans zmian", format=",.0f"),
            ],
        )
    )
    focus_rule = (
        alt.Chart(latest_point)
        .mark_rule(color="#1f3447", strokeWidth=1.2, opacity=0.7)
        .encode(x="Analysis Date:T")
    )
    current_label = (
        alt.Chart(latest_point)
        .mark_text(
            align="left",
            baseline="middle",
            dx=12,
            dy=-4,
            color="#eef4ff",
            fontSize=13,
            fontWeight="bold",
        )
        .encode(x="Analysis Date:T", y="Quantity_Curr:Q", text="Current Label:N")
    )
    previous_label = (
        alt.Chart(latest_point)
        .mark_text(
            align="left",
            baseline="middle",
            dx=12,
            dy=14,
            color="#b6c8dd",
            fontSize=12,
            fontWeight="bold",
        )
        .encode(x="Analysis Date:T", y="Quantity_Prev:Q", text="Previous Label:N")
    )
    chart = alt.layer(
        current_area,
        prev_line,
        focus_rule,
        current_line,
        current_label,
        previous_label,
    ).properties(
        height=420,
        padding={"left": 6, "right": 22, "top": 12, "bottom": 8},
    )
    return apply_chart_theme(chart)


def build_delta_chart(date_summary, x_title):
    chart_data = date_summary.sort_values("Analysis Date").copy()
    chart_data["Abs Delta"] = chart_data["Delta"].abs()
    label_source = chart_data.nlargest(min(6, len(chart_data)), "Abs Delta").copy()
    label_source["Delta Label"] = label_source["Delta"].map(lambda value: f"{value:+,.0f}")

    zero_line = alt.Chart(pd.DataFrame({"y": [0]})).mark_rule(color="#2b3a4d", strokeWidth=1.2).encode(y="y:Q")
    bars = (
        alt.Chart(chart_data)
        .mark_bar(cornerRadiusTopLeft=7, cornerRadiusTopRight=7, opacity=0.92, size=20)
        .encode(
            x=alt.X("Analysis Date:T", title=x_title, axis=alt.Axis(labelAngle=-24, labelLimit=140)),
            y=alt.Y("Delta:Q", title="Zmiana ilości"),
            color=alt.condition(
                alt.datum.Delta >= 0,
                alt.value("#5f8b75"),
                alt.value("#c56b61"),
            ),
            tooltip=[
                alt.Tooltip("Analysis Date:T", title="Data"),
                alt.Tooltip("Delta:Q", title="Zmiana ilości", format=",.0f"),
                alt.Tooltip("Alerts:Q", title="Liczba alertów"),
            ],
        )
    )
    positive_labels = (
        alt.Chart(label_source[label_source["Delta"] >= 0])
        .mark_text(
            baseline="bottom",
            dy=-6,
            color="#e8f3ed",
            fontWeight="bold",
            fontSize=11,
        )
        .encode(
            x="Analysis Date:T",
            y="Delta:Q",
            text="Delta Label:N",
        )
    )
    negative_labels = (
        alt.Chart(label_source[label_source["Delta"] < 0])
        .mark_text(
            baseline="top",
            dy=8,
            color="#f8d7d3",
            fontWeight="bold",
            fontSize=11,
        )
        .encode(
            x="Analysis Date:T",
            y="Delta:Q",
            text="Delta Label:N",
        )
    )
    chart = alt.layer(zero_line, bars, positive_labels, negative_labels).properties(
        height=320,
        padding={"left": 6, "right": 22, "top": 12, "bottom": 8},
    )
    return apply_chart_theme(chart)


def build_product_bar_source(product_summary, chart_type):
    if chart_type == "increase":
        return (
            product_summary[product_summary["Delta"] > 0]
            .nlargest(10, "Delta")[["Part Number", "Part Description", "Delta"]]
            .reset_index(drop=True)
        )

    return (
        product_summary[product_summary["Delta"] < 0]
        .nsmallest(10, "Delta")[["Part Number", "Part Description", "Delta"]]
        .reset_index(drop=True)
    )


def build_product_bar_chart(product_summary, chart_type):
    source = build_product_bar_source(product_summary, chart_type)
    if chart_type == "increase":
        color = "#5f8b75"
        title = "Największe wzrosty"
    else:
        color = "#c56b61"
        title = "Największe spadki"

    if source.empty:
        return None, title

    source["Display Label"] = source["Part Description"].map(
        lambda value: value if len(str(value)) <= 42 else f"{str(value)[:39]}..."
    )
    source["Delta Label"] = source["Delta"].map(lambda value: f"{value:+,.0f}")
    chart = (
        alt.Chart(source)
        .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6, color=color, opacity=0.94)
        .encode(
            x=alt.X("Delta:Q", title="Zmiana ilości"),
            y=alt.Y(
                "Display Label:N",
                sort="-x",
                title=None,
                axis=alt.Axis(labelLimit=280, labelPadding=12),
            ),
            tooltip=[
                alt.Tooltip("Part Description:N", title="Produkt"),
                alt.Tooltip("Delta:Q", title="Zmiana ilości", format=",.0f"),
            ],
        )
        .properties(height=max(340, len(source) * 34))
    )
    text = (
        alt.Chart(source)
        .mark_text(
            align="left" if chart_type == "increase" else "right",
            dx=8 if chart_type == "increase" else -8,
            color="#e7edf6",
            fontWeight="bold",
            fontSize=11,
        )
        .encode(
            x="Delta:Q",
            y=alt.Y("Display Label:N", sort="-x", title=None),
            text="Delta Label:N",
        )
    )
    try:
        layered = alt.layer(chart, text).properties(height=max(340, len(source) * 34))
        return apply_chart_theme(layered), title
    except Exception:
        return apply_chart_theme(chart), title


def build_change_mix_source(dataframe):
    mix = (
        dataframe.groupby("Change Direction", as_index=False)
        .agg(Rows=("Change Direction", "size"), Total_Delta=("Delta", "sum"))
    )
    mix["Direction Label"] = mix["Change Direction"].map(get_change_label)
    mix["Share"] = mix["Rows"] / max(int(mix["Rows"].sum()), 1)
    return mix


def build_change_mix_chart(dataframe):
    mix = build_change_mix_source(dataframe)
    order = ["Wzrost", "Spadek", "Bez zmian"]
    colors = ["#5f8b75", "#c56b61", "#6f87ab"]
    bars = (
        alt.Chart(mix)
        .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6, size=28)
        .encode(
            x=alt.X("Rows:Q", title="Liczba pozycji"),
            y=alt.Y("Direction Label:N", sort=order, title=None),
            color=alt.Color(
                "Direction Label:N",
                sort=order,
                scale=alt.Scale(domain=order, range=colors),
                legend=None,
            ),
            tooltip=[
                alt.Tooltip("Direction Label:N", title="Kierunek"),
                alt.Tooltip("Rows:Q", title="Liczba pozycji"),
                alt.Tooltip("Share:Q", title="Udział", format=".1%"),
                alt.Tooltip("Total_Delta:Q", title="Bilans zmian", format=",.0f"),
            ],
        )
    )
    labels = (
        alt.Chart(mix)
        .mark_text(align="left", dx=8, color="#e7edf6", fontSize=11, fontWeight="bold")
        .encode(
            x="Rows:Q",
            y=alt.Y("Direction Label:N", sort=order, title=None),
            text=alt.Text("Share:Q", format=".1%"),
        )
    )
    chart = alt.layer(bars, labels).properties(
        height=240,
        padding={"left": 6, "right": 20, "top": 8, "bottom": 8},
    )
    return apply_chart_theme(chart)


def build_key_findings(dataframe, product_summary, date_summary, date_basis):
    findings = []
    if dataframe.empty or product_summary.empty:
        return findings

    ranked = product_summary.assign(Abs_Delta=product_summary["Delta"].abs()).sort_values(
        "Abs_Delta", ascending=False
    )
    largest_move = ranked.iloc[0]
    findings.append(
        {
            "label": "Największa zmiana",
            "title": largest_move["Part Description"],
            "copy": (
                f"Najsilniejszy ruch w badanym oknie: {format_signed_int(largest_move['Delta'])} szt."
            ),
        }
    )

    negative = product_summary[product_summary["Delta"] < 0]
    if not negative.empty:
        largest_drop = negative.nsmallest(1, "Delta").iloc[0]
        findings.append(
            {
                "label": "Największy spadek",
                "title": largest_drop["Part Description"],
                "copy": (
                    f"Ta pozycja notuje najmocniejszy spadek: {format_signed_int(largest_drop['Delta'])} szt."
                ),
            }
        )

    alert_summary = (
        dataframe.groupby(["Part Description", "Part Number"], as_index=False)
        .agg(Alert_Count=("Alert", "sum"), Delta=("Delta", "sum"))
        .sort_values(["Alert_Count", "Delta"], ascending=[False, False])
    )
    if not alert_summary.empty and alert_summary.iloc[0]["Alert_Count"] > 0:
        top_alert = alert_summary.iloc[0]
        findings.append(
            {
                "label": "Najwięcej alertów",
                "title": top_alert["Part Description"],
                "copy": (
                    f"Produkt pojawia się w {int(top_alert['Alert_Count'])} alertach i wymaga szybkiej weryfikacji."
                ),
            }
        )

    if not date_summary.empty:
        peak_row = date_summary.sort_values(["Alerts", "Delta"], ascending=[False, False]).iloc[0]
        day_slice = dataframe[dataframe[date_basis] == peak_row["Analysis Date"]]
        if not day_slice.empty:
            day_product = (
                day_slice.groupby(["Part Description", "Part Number"], as_index=False)["Delta"]
                .sum()
                .assign(Abs_Delta=lambda df: df["Delta"].abs())
                .sort_values("Abs_Delta", ascending=False)
                .iloc[0]
            )
            findings.append(
                {
                    "label": "Kluczowy dzień",
                    "title": day_product["Part Description"],
                    "copy": (
                        f"Dnia {format_date(peak_row['Analysis Date'])} ta pozycja wygenerowała zmianę {format_signed_int(day_product['Delta'])} szt."
                    ),
                }
            )

    return findings[:4]


def build_matrix(dataframe, date_basis, metric_name):
    pivot_prev = pd.pivot_table(
        dataframe,
        index="Product Label",
        columns=date_basis,
        values="Quantity_Prev",
        aggfunc="sum",
        fill_value=0,
    )
    pivot_curr = pd.pivot_table(
        dataframe,
        index="Product Label",
        columns=date_basis,
        values="Quantity_Curr",
        aggfunc="sum",
        fill_value=0,
    )

    if metric_name == "Current Quantity":
        matrix = pivot_curr
    elif metric_name == "Previous Quantity":
        matrix = pivot_prev
    elif metric_name == "Delta":
        matrix = pivot_curr - pivot_prev
    else:
        matrix = ((pivot_curr - pivot_prev) / pivot_prev.replace(0, pd.NA)) * 100
        matrix = matrix.fillna(0)

    matrix = matrix.sort_index(axis=1)
    matrix.columns = [pd.Timestamp(column).strftime("%Y-%m-%d") for column in matrix.columns]
    return matrix


def blend_hex(start_hex, end_hex, ratio):
    ratio = max(0.0, min(1.0, float(ratio)))
    start = tuple(int(start_hex[index : index + 2], 16) for index in (1, 3, 5))
    end = tuple(int(end_hex[index : index + 2], 16) for index in (1, 3, 5))
    blended = tuple(
        round(start[channel] + (end[channel] - start[channel]) * ratio)
        for channel in range(3)
    )
    return "#{:02x}{:02x}{:02x}".format(*blended)


def style_value(value, metric_name, max_value, max_abs):
    base_style = "text-align: center;"

    if pd.isna(value):
        return base_style

    if metric_name == "Current Quantity":
        ratio = 0 if max_value <= 0 else value / max_value
        background = blend_hex("#f8fafc", "#2563eb", ratio)
        text_color = "#ffffff" if ratio > 0.6 else "#0f172a"
    elif metric_name == "Previous Quantity":
        ratio = 0 if max_value <= 0 else value / max_value
        background = blend_hex("#f8fafc", "#64748b", ratio)
        text_color = "#ffffff" if ratio > 0.6 else "#0f172a"
    else:
        ratio = 0 if max_abs <= 0 else abs(value) / max_abs
        if value > 0:
            background = blend_hex("#f0fdf4", "#16a34a", ratio)
            text_color = "#ffffff" if ratio > 0.6 else "#14532d"
        elif value < 0:
            background = blend_hex("#fef2f2", "#dc2626", ratio)
            text_color = "#ffffff" if ratio > 0.6 else "#7f1d1d"
        else:
            background = "#f8fafc"
            text_color = "#334155"

    return f"{base_style} background-color: {background}; color: {text_color};"


def style_matrix(matrix, metric_name):
    max_value = float(matrix.max().max()) if not matrix.empty else 0
    max_abs = float(matrix.abs().max().max()) if not matrix.empty else 0
    style_frame = matrix.map(
        lambda value: style_value(value, metric_name, max_value, max_abs)
    )

    styled = matrix.style.apply(lambda _: style_frame, axis=None)

    if metric_name in ["Current Quantity", "Previous Quantity", "Delta"]:
        return styled.format("{:,.0f}")

    return styled.format("{:+,.1f}%")


def style_excel_header(worksheet, row_number):
    fill = PatternFill(fill_type="solid", fgColor="0F172A")
    font = Font(color="FFFFFF", bold=True)
    alignment = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )
    for cell in worksheet[row_number]:
        cell.fill = fill
        cell.font = font
        cell.alignment = alignment
        cell.border = border


def autosize_worksheet(worksheet, min_width=12, max_width=42):
    for column_cells in worksheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        worksheet.column_dimensions[column_letter].width = min(
            max(max_length + 2, min_width), max_width
        )


def insert_logo(worksheet, cell):
    if not logo_available():
        return

    logo = OpenpyxlImage(str(LOGO_PATH))
    logo.width = 160
    logo.height = 48
    worksheet.add_image(logo, cell)


def decorate_delta_column(worksheet, header_row=1):
    headers = {cell.value: cell.column for cell in worksheet[header_row]}
    delta_column = headers.get("Delta")
    percent_column = headers.get("Percent Change")
    if delta_column is None and percent_column is None:
        return

    green_fill = PatternFill(fill_type="solid", fgColor="DCFCE7")
    red_fill = PatternFill(fill_type="solid", fgColor="FEE2E2")
    blue_fill = PatternFill(fill_type="solid", fgColor="DBEAFE")

    for row in range(header_row + 1, worksheet.max_row + 1):
        if delta_column is not None:
            delta_cell = worksheet.cell(row=row, column=delta_column)
            if isinstance(delta_cell.value, (int, float)):
                delta_cell.fill = green_fill if delta_cell.value > 0 else red_fill if delta_cell.value < 0 else blue_fill
                delta_cell.font = Font(color="000000", bold=False)
        if percent_column is not None:
            percent_cell = worksheet.cell(row=row, column=percent_column)
            if isinstance(percent_cell.value, (int, float)):
                percent_cell.number_format = '0.0"%"'
                percent_cell.font = Font(color="000000", bold=False)


def excel_fill_color(value, metric_name, max_value, max_abs):
    if metric_name == "Current Quantity":
        ratio = 0 if max_value <= 0 else float(value) / max_value
        return blend_hex("#eff6ff", "#93c5fd", ratio)
    if metric_name == "Previous Quantity":
        ratio = 0 if max_value <= 0 else float(value) / max_value
        return blend_hex("#f8fafc", "#cbd5e1", ratio)
    ratio = 0 if max_abs <= 0 else abs(float(value)) / max_abs
    if value > 0:
        return blend_hex("#f0fdf4", "#86efac", ratio)
    if value < 0:
        return blend_hex("#fef2f2", "#fca5a5", ratio)
    return "#e2e8f0"


def ensure_numeric_cells_black(worksheet, start_row=1):
    for row in worksheet.iter_rows(min_row=start_row, max_row=worksheet.max_row):
        for cell in row:
            if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                cell.font = Font(color="000000", bold=bool(cell.font.bold))


def apply_polish_calendar_highlights(worksheet, date_columns, header_row=1):
    headers = {cell.value: cell.column for cell in worksheet[header_row]}
    saturday_fill = PatternFill(fill_type="solid", fgColor="DBEAFE")
    sunday_fill = PatternFill(fill_type="solid", fgColor="FEE2E2")
    holiday_fill = PatternFill(fill_type="solid", fgColor="FEF3C7")

    for column_name in date_columns:
        column_index = headers.get(column_name)
        if column_index is None:
            continue
        for row in range(header_row + 1, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=column_index)
            if cell.value in (None, ""):
                continue
            try:
                day_info = classify_polish_day(cell.value)
            except Exception:
                continue
            if day_info["Is Holiday"]:
                cell.fill = holiday_fill
            elif day_info["Day Type"] == "Saturday":
                cell.fill = saturday_fill
            elif day_info["Day Type"] == "Sunday":
                cell.fill = sunday_fill
            cell.font = Font(color="000000", bold=bool(cell.font.bold))


def style_matrix_sheet(worksheet, metric_name, header_row=1, start_col=2):
    label_fill = PatternFill(fill_type="solid", fgColor="E2E8F0")
    label_font = Font(color="0F172A", bold=True)
    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    values = []
    for row in range(header_row + 1, worksheet.max_row + 1):
        for col in range(start_col, worksheet.max_column + 1):
            cell_value = worksheet.cell(row=row, column=col).value
            if isinstance(cell_value, (int, float)):
                values.append(float(cell_value))

    max_value = max(values) if values else 0
    max_abs = max((abs(value) for value in values), default=0)

    for row in range(header_row + 1, worksheet.max_row + 1):
        label_cell = worksheet.cell(row=row, column=1)
        label_cell.fill = label_fill
        label_cell.font = label_font
        label_cell.border = thin_border
        for col in range(start_col, worksheet.max_column + 1):
            cell = worksheet.cell(row=row, column=col)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(cell.value, (int, float)):
                bg = excel_fill_color(cell.value, metric_name, max_value, max_abs).replace("#", "")
                cell.fill = PatternFill(fill_type="solid", fgColor=bg)
                cell.font = Font(color="000000", bold=False)
                if metric_name == "Percent Change":
                    cell.number_format = '0.0"%"'
                else:
                    cell.number_format = '#,##0'


def highlight_calendar_rows(worksheet, header_row=1):
    date_column = None
    for cell in worksheet[header_row]:
        if cell.value == "Date":
            date_column = cell.column
            break
    if date_column is None:
        return

    saturday_fill = PatternFill(fill_type="solid", fgColor="DBEAFE")
    sunday_fill = PatternFill(fill_type="solid", fgColor="FEE2E2")
    holiday_fill = PatternFill(fill_type="solid", fgColor="FEF3C7")

    for row in range(header_row + 1, worksheet.max_row + 1):
        date_value = worksheet.cell(row=row, column=date_column).value
        if not date_value:
            continue
        try:
            day_info = classify_polish_day(date_value)
        except Exception:
            continue

        row_fill = None
        if day_info["Is Holiday"]:
            row_fill = holiday_fill
        elif day_info["Day Type"] == "Saturday":
            row_fill = saturday_fill
        elif day_info["Day Type"] == "Sunday":
            row_fill = sunday_fill

        if row_fill is None:
            continue

        for col in range(1, worksheet.max_column + 1):
            worksheet.cell(row=row, column=col).fill = row_fill


def highlight_weekly_rows(worksheet, header_row=1):
    headers = {cell.value: cell.column for cell in worksheet[header_row]}
    status_column = headers.get("Week Status")
    alert_column = headers.get("Any Weekly Alert")
    reference_column = headers.get("Is Reference Week")
    if status_column is None and alert_column is None and reference_column is None:
        return

    partial_fill = PatternFill(fill_type="solid", fgColor="FEF3C7")
    open_fill = PatternFill(fill_type="solid", fgColor="FCE7F3")
    reference_fill = PatternFill(fill_type="solid", fgColor="DBEAFE")
    alert_fill = PatternFill(fill_type="solid", fgColor="FDECEC")

    for row in range(header_row + 1, worksheet.max_row + 1):
        status_value = worksheet.cell(row=row, column=status_column).value if status_column else None
        is_alert = worksheet.cell(row=row, column=alert_column).value if alert_column else False
        is_reference = worksheet.cell(row=row, column=reference_column).value if reference_column else False

        row_fill = None
        if bool(is_reference):
            row_fill = reference_fill
        elif bool(is_alert):
            row_fill = alert_fill
        elif status_value == "Partial range":
            row_fill = partial_fill
        elif status_value == "Open week":
            row_fill = open_fill

        if row_fill is None:
            continue

        for col in range(1, worksheet.max_column + 1):
            worksheet.cell(row=row, column=col).fill = row_fill


def write_summary_sheet(
    worksheet,
    prev_meta,
    curr_meta,
    detail_df,
    product_summary,
    weekly_summary,
    date_basis,
    selected_start_date,
    selected_end_date,
    key_findings,
):
    insert_logo(worksheet, "G1")
    worksheet.merge_cells("A1:F1")
    worksheet["A1"] = BRAND_NAME
    worksheet["A1"].font = Font(size=14, bold=True, color="FFFFFF")
    worksheet["A1"].fill = PatternFill(fill_type="solid", fgColor="0F172A")
    worksheet["A1"].alignment = Alignment(horizontal="center", vertical="center")

    worksheet.merge_cells("A2:F2")
    worksheet["A2"] = "Release Change Executive Summary"
    worksheet["A2"].font = Font(size=16, bold=True, color="0F172A")
    worksheet["A2"].alignment = Alignment(horizontal="left", vertical="center")

    worksheet["A4"] = "PO Number"
    worksheet["B4"] = curr_meta["po_number"]
    worksheet["A5"] = "Previous Release"
    worksheet["B5"] = format_release_summary(prev_meta)
    worksheet["A6"] = "Current Release"
    worksheet["B6"] = format_release_summary(curr_meta)
    worksheet["A7"] = "Date Basis"
    worksheet["B7"] = date_basis
    worksheet["A8"] = "Selected Period"
    worksheet["B8"] = f"{selected_start_date:%Y-%m-%d} to {selected_end_date:%Y-%m-%d}"
    worksheet["A9"] = "Planner"
    worksheet["B9"] = curr_meta["planner_name"]
    worksheet["A10"] = "Planner Email"
    worksheet["B10"] = curr_meta["planner_email"]

    total_prev = detail_df["Quantity_Prev"].sum()
    total_curr = detail_df["Quantity_Curr"].sum()
    total_delta = detail_df["Delta"].sum()
    alert_count = int(detail_df["Alert"].sum())
    products_changed = int((product_summary["Delta"] != 0).sum())

    worksheet["D4"] = "Previous Qty"
    worksheet["E4"] = total_prev
    worksheet["D5"] = "Current Qty"
    worksheet["E5"] = total_curr
    worksheet["D6"] = "Net Delta"
    worksheet["E6"] = total_delta
    worksheet["D7"] = "Alert Rows"
    worksheet["E7"] = alert_count
    worksheet["D8"] = "Products Changed"
    worksheet["E8"] = products_changed

    reference_week = get_last_completed_reference_week(selected_end_date)
    reference_row, previous_row = get_reference_week_rows(weekly_summary)
    worksheet["D10"] = "Reference Week"
    worksheet["E10"] = reference_row["Week Label"] if reference_row is not None else reference_week.week_label
    worksheet["D11"] = "Reference Current Qty"
    worksheet["E11"] = float(reference_row["Quantity_Curr"]) if reference_row is not None else 0
    worksheet["D12"] = "Reference Release %"
    worksheet["E12"] = (
        format_percent_display(reference_row["Release Percent Label"])
        if reference_row is not None
        else "n/a"
    )
    worksheet["D13"] = "Reference WoW %"
    worksheet["E13"] = (
        format_percent_display(reference_row["WoW Percent Label"])
        if reference_row is not None
        else "n/a"
    )
    worksheet["D14"] = "Working Days PL"
    worksheet["E14"] = int(reference_row["Working_Days_PL"]) if reference_row is not None else 0
    worksheet["D15"] = "Previous Closed Week"
    worksheet["E15"] = previous_row["Week Label"] if previous_row is not None else "n/a"

    worksheet["A13"] = "Key Findings"
    worksheet["A13"].font = Font(size=13, bold=True, color="0F172A")
    for idx, finding in enumerate(key_findings[:4], start=14):
        worksheet[f"A{idx}"] = finding["label"]
        worksheet[f"A{idx}"].font = Font(bold=True, color="2563EB")
        worksheet[f"B{idx}"] = finding["title"]
        worksheet[f"C{idx}"] = finding["copy"]

    worksheet["A20"] = "Top Product Changes"
    worksheet["A20"].font = Font(size=13, bold=True, color="0F172A")
    worksheet["A21"] = "Part Number"
    worksheet["B21"] = "Part Description"
    worksheet["C21"] = "Previous Qty"
    worksheet["D21"] = "Current Qty"
    worksheet["E21"] = "Delta"
    worksheet["F21"] = "Alert Count"
    style_excel_header(worksheet, 21)

    top_rows = (
        product_summary.assign(Abs_Delta=product_summary["Delta"].abs())
        .sort_values("Abs_Delta", ascending=False)
        .drop(columns=["Abs_Delta", "Product Label", "Change Direction"])
        .head(10)
    )
    start_row = 22
    for offset, (_, row) in enumerate(top_rows.iterrows()):
        excel_row = start_row + offset
        worksheet.cell(excel_row, 1, row["Part Number"])
        worksheet.cell(excel_row, 2, row["Part Description"])
        worksheet.cell(excel_row, 3, row["Quantity_Prev"])
        worksheet.cell(excel_row, 4, row["Quantity_Curr"])
        worksheet.cell(excel_row, 5, row["Delta"])
        worksheet.cell(excel_row, 6, int(row["Alert_Count"]))

    decorate_delta_column(worksheet, header_row=21)
    autosize_worksheet(worksheet)


def to_excel_bytes(
    detail_df,
    weekly_summary,
    current_matrix_df,
    delta_matrix_df,
    prev_meta,
    curr_meta,
    product_summary,
    date_basis,
    selected_start_date,
    selected_end_date,
    key_findings,
):
    output = io.BytesIO()
    detail_export = detail_df.copy()
    detail_export["Ship Date"] = detail_export["Ship Date"].dt.strftime("%Y-%m-%d")
    detail_export["Receipt Date"] = detail_export["Receipt Date"].dt.strftime("%Y-%m-%d")
    weekly_export = weekly_summary[
        [
            "Week Label",
            "Week Start",
            "Week End",
            "Week Status",
            "Working_Days_PL",
            "Holidays_PL",
            "Weekend_Days",
            "Products",
            "Quantity_Prev",
            "Quantity_Curr",
            "Delta",
            "Release Percent Label",
            "Previous Week Current Qty",
            "WoW Delta",
            "WoW Percent Label",
            "Avg Current / Working Day",
            "Release Alert",
            "WoW Alert",
            "Any Weekly Alert",
            "Is Reference Week",
        ]
    ].copy()
    weekly_export["Week Start"] = pd.to_datetime(weekly_export["Week Start"]).dt.strftime("%Y-%m-%d")
    weekly_export["Week End"] = pd.to_datetime(weekly_export["Week End"]).dt.strftime("%Y-%m-%d")
    weekly_export["Release Percent Label"] = weekly_export["Release Percent Label"].map(format_percent_display)
    weekly_export["WoW Percent Label"] = weekly_export["WoW Percent Label"].map(format_percent_display)
    weekly_export = weekly_export.rename(
        columns={
            "Working_Days_PL": "Working Days PL",
            "Holidays_PL": "Polish Holidays",
            "Weekend_Days": "Weekend Days",
            "Quantity_Prev": "Previous Release Qty",
            "Quantity_Curr": "Current Release Qty",
            "Release Percent Label": "Release Change %",
            "Previous Week Current Qty": "Previous Week Current Qty",
            "WoW Percent Label": "WoW Change %",
            "Avg Current / Working Day": "Current / Working Day",
        }
    )
    calendar_export = build_calendar_frame(selected_start_date, selected_end_date).copy()
    calendar_export["Date"] = pd.to_datetime(calendar_export["Date"]).dt.strftime("%Y-%m-%d")
    calendar_export["Week Start"] = pd.to_datetime(calendar_export["Week Start"]).dt.strftime("%Y-%m-%d")
    calendar_export["Week End"] = pd.to_datetime(calendar_export["Week End"]).dt.strftime("%Y-%m-%d")
    current_matrix_export = current_matrix_df.reset_index()
    delta_matrix_export = delta_matrix_df.reset_index()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, sheet_name="Executive Summary", index=False)
        detail_export.to_excel(writer, sheet_name="Detailed Data", index=False)
        weekly_export.to_excel(writer, sheet_name="Weekly Summary", index=False)
        calendar_export.to_excel(writer, sheet_name="Calendar PL", index=False)
        current_matrix_export.to_excel(writer, sheet_name="Current Matrix", index=False)
        delta_matrix_export.to_excel(writer, sheet_name="Delta Heatmap", index=False)

        summary_sheet = writer.book["Executive Summary"]
        write_summary_sheet(
            summary_sheet,
            prev_meta,
            curr_meta,
            detail_df,
            product_summary,
            weekly_summary,
            date_basis,
            selected_start_date,
            selected_end_date,
            key_findings,
        )

        detail_sheet = writer.book["Detailed Data"]
        style_excel_header(detail_sheet, 1)
        decorate_delta_column(detail_sheet, header_row=1)
        apply_polish_calendar_highlights(detail_sheet, ["Ship Date", "Receipt Date"], header_row=1)
        detail_sheet.freeze_panes = "A2"
        autosize_worksheet(detail_sheet)
        ensure_numeric_cells_black(detail_sheet, start_row=2)

        weekly_sheet = writer.book["Weekly Summary"]
        style_excel_header(weekly_sheet, 1)
        highlight_weekly_rows(weekly_sheet, header_row=1)
        decorate_delta_column(weekly_sheet, header_row=1)
        weekly_sheet.freeze_panes = "A2"
        autosize_worksheet(weekly_sheet)
        ensure_numeric_cells_black(weekly_sheet, start_row=2)

        calendar_sheet = writer.book["Calendar PL"]
        style_excel_header(calendar_sheet, 1)
        highlight_calendar_rows(calendar_sheet, header_row=1)
        apply_polish_calendar_highlights(calendar_sheet, ["Date"], header_row=1)
        calendar_sheet.freeze_panes = "A2"
        autosize_worksheet(calendar_sheet)
        ensure_numeric_cells_black(calendar_sheet, start_row=2)

        current_matrix_sheet = writer.book["Current Matrix"]
        style_excel_header(current_matrix_sheet, 1)
        current_matrix_sheet.freeze_panes = "B2"
        autosize_worksheet(current_matrix_sheet)
        style_matrix_sheet(current_matrix_sheet, "Current Quantity")
        ensure_numeric_cells_black(current_matrix_sheet, start_row=2)

        delta_heatmap_sheet = writer.book["Delta Heatmap"]
        style_excel_header(delta_heatmap_sheet, 1)
        delta_heatmap_sheet.freeze_panes = "B2"
        autosize_worksheet(delta_heatmap_sheet)
        style_matrix_sheet(delta_heatmap_sheet, "Delta")
        ensure_numeric_cells_black(delta_heatmap_sheet, start_row=2)

        ensure_numeric_cells_black(summary_sheet, start_row=1)

    return output.getvalue()


init_auth_state()

if not st.session_state["authenticated"]:
    render_login_screen()
    st.stop()

app_sidebar, app_main = st.columns([0.27, 0.73], gap="large")

with app_sidebar:
    render_sidebar_user(st)
    prev_file, current_file = render_upload_section()

if prev_file is None or current_file is None:
    with app_sidebar:
        render_welcome_side_panel(prev_file, current_file)
    with app_main:
        render_welcome_state(prev_file, current_file)
    st.stop()

try:
    prev_df, prev_meta = load_release(prev_file.getvalue(), prev_file.name)
    curr_df, curr_meta = load_release(current_file.getvalue(), current_file.name)
    result = compare_releases(prev_df, curr_df)
except Exception as exc:
    with app_sidebar:
        render_welcome_side_panel(prev_file, current_file)
    with app_main:
        render_welcome_state(prev_file, current_file)
        st.error(f"Błąd wczytywania plików: {exc}")
    st.stop()

brand_context = detect_brand_context(prev_meta, curr_meta)
with app_sidebar:
    filter_state = render_analysis_side_panel(
        result,
        brand_context,
        prev_meta=prev_meta,
        curr_meta=curr_meta,
    )

date_basis = filter_state["date_basis"]
selected_start_date = filter_state["selected_start_date"]
selected_end_date = filter_state["selected_end_date"]
selected_products = filter_state["selected_products"]
search_term = filter_state["search_term"]
selected_change_directions = filter_state["selected_change_directions"]
only_alerts = filter_state["only_alerts"]

filtered_df = result.copy()
filtered_df = filtered_df[
    filtered_df[date_basis].dt.date.between(
        selected_start_date, selected_end_date
    )
]

if selected_products:
    filtered_df = filtered_df[filtered_df["Product Label"].isin(selected_products)]
else:
    filtered_df = filtered_df.iloc[0:0]

if search_term.strip():
    query = search_term.strip().lower()
    filtered_df = filtered_df[
        filtered_df["Part Number"].str.lower().str.contains(query, na=False)
        | filtered_df["Part Description"].str.lower().str.contains(query, na=False)
    ]

filtered_df = filtered_df[
    filtered_df["Change Direction"].isin(selected_change_directions)
]

if only_alerts:
    filtered_df = filtered_df[filtered_df["Alert"]]

product_summary = summarize_products(filtered_df)
date_summary = summarize_dates(filtered_df, date_basis)
weekly_summary = build_weekly_summary(
    filtered_df,
    date_basis,
    selected_start_date,
    selected_end_date,
    selected_end_date,
    THRESHOLD,
)
key_findings = build_key_findings(
    filtered_df, product_summary, date_summary, date_basis
)

detail_export_table = build_detail_export_table(filtered_df)
csv_bytes = detail_export_table.to_csv(index=False).encode("utf-8")
current_matrix_for_export = build_matrix(filtered_df, date_basis, "Current Quantity")
delta_matrix_for_export = build_matrix(filtered_df, date_basis, "Delta")
excel_bytes = to_excel_bytes(
    filtered_df,
    weekly_summary,
    current_matrix_for_export,
    delta_matrix_for_export,
    prev_meta,
    curr_meta,
    product_summary,
    date_basis,
    selected_start_date,
    selected_end_date,
    key_findings,
)

with app_sidebar:
    render_export_actions(csv_bytes, excel_bytes)

with app_main:
    render_analysis_main(
        filtered_df,
        product_summary,
        date_summary,
        weekly_summary,
        key_findings,
        prev_meta,
        curr_meta,
        date_basis,
        selected_start_date,
        selected_end_date,
        excel_bytes=excel_bytes,
        csv_bytes=csv_bytes,
    )

st.stop()

app_sidebar, app_main = st.columns([0.28, 0.72], gap="large")

with app_main:
    upload_left, upload_right = st.columns(2, gap="large")
    with upload_left:
        render_upload_card(
            "Krok 1",
            "Poprzedni release / poprzedni plan",
            "Dodaj bazowy plik Excel, do ktorego bedzie porownywany aktualny stan zamowien i wysylek.",
        )
        prev_file = st.file_uploader(
            "Upload Previous Release",
            type=["xlsx"],
            key="previous_release_upload",
            label_visibility="collapsed",
        )
    with upload_right:
        render_upload_card(
            "Krok 2",
            "Aktualny release / aktualny plan",
            "Dodaj nowy plik Excel, aby dashboard automatycznie policzyl delty, alerty i zmiany procentowe.",
        )
        current_file = st.file_uploader(
            "Upload Current Release",
            type=["xlsx"],
            key="current_release_upload",
            label_visibility="collapsed",
        )

if prev_file is None and current_file is None:
    with app_sidebar:
        render_welcome_side_panel(prev_file, current_file)
    with app_main:
        quick_cols = st.columns(3, gap="large")
        with quick_cols[0]:
            render_quick_card(
                "Czytelny dashboard porownawczy",
                "Aplikacja zestawia poprzedni i aktualny release, od razu pokazujac bilans zmian, alerty oraz produkty z najwiekszym ruchem.",
            )
        with quick_cols[1]:
            render_quick_card(
                "Macierz podobna do Excela",
                "Otrzymujesz widok tabelaryczny z datami, zmianami ilosci i filtrowaniem po produkcie, kierunku ruchu oraz zakresie dat.",
            )
        with quick_cols[2]:
            render_quick_card(
                "Raport gotowy do wyslania",
                "Po analizie pobierzesz CSV oraz biznesowy raport Excel z podsumowaniem KPI i kluczowymi zmianami.",
            )
        st.info(
            "Zacznij od dodania dwoch plikow Excel. Po zaladowaniu obu release'ow dashboard uruchomi pelna analize porownawcza."
        )
    st.stop()

if prev_file is None or current_file is None:
    with app_sidebar:
        render_welcome_side_panel(prev_file, current_file)
    with app_main:
        missing_label = "poprzedni" if prev_file is None else "aktualny"
        loaded_label = "aktualny" if prev_file is None else "poprzedni"
        st.info(
            f"Plik {loaded_label} jest juz dodany. Dodaj jeszcze plik {missing_label}, aby uruchomic analize i wygenerowac dashboard."
        )
    st.stop()

try:
    prev_df, prev_meta = load_release(prev_file.getvalue(), prev_file.name)
    curr_df, curr_meta = load_release(current_file.getvalue(), current_file.name)
    result = compare_releases(prev_df, curr_df)
except Exception as exc:
    with app_sidebar:
        render_welcome_side_panel(prev_file, current_file)
    with app_main:
        st.error(f"Blad: {exc}")
    st.stop()

brand_context = detect_brand_context(prev_meta, curr_meta)
with app_sidebar:
    filter_state = render_analysis_side_panel(result, brand_context)

date_basis = filter_state["date_basis"]
selected_start_date = filter_state["selected_start_date"]
selected_end_date = filter_state["selected_end_date"]
selected_products = filter_state["selected_products"]
search_term = filter_state["search_term"]
selected_change_directions = filter_state["selected_change_directions"]
only_alerts = filter_state["only_alerts"]

filtered_df = result.copy()
filtered_df = filtered_df[
    filtered_df[date_basis].dt.date.between(
        selected_start_date, selected_end_date
    )
]

if selected_products:
    filtered_df = filtered_df[filtered_df["Product Label"].isin(selected_products)]
else:
    filtered_df = filtered_df.iloc[0:0]

if search_term.strip():
    query = search_term.strip().lower()
    filtered_df = filtered_df[
        filtered_df["Part Number"].str.lower().str.contains(query, na=False)
        | filtered_df["Part Description"].str.lower().str.contains(query, na=False)
    ]

filtered_df = filtered_df[
    filtered_df["Change Direction"].isin(selected_change_directions)
]

if only_alerts:
    filtered_df = filtered_df[filtered_df["Alert"]]

product_summary = summarize_products(filtered_df)
date_summary = summarize_dates(filtered_df, date_basis)
weekly_summary = build_weekly_summary(
    filtered_df,
    date_basis,
    selected_start_date,
    selected_end_date,
    selected_end_date,
    THRESHOLD,
)
key_findings = build_key_findings(
    filtered_df, product_summary, date_summary, date_basis
)

with app_main:
    render_analysis_main(
        filtered_df,
        product_summary,
        date_summary,
        weekly_summary,
        key_findings,
        prev_meta,
        curr_meta,
        date_basis,
        selected_start_date,
        selected_end_date,
    )

st.stop()

render_sidebar_user(st)

upload_left, upload_right = st.columns(2, gap="large")
with upload_left:
    render_upload_card(
        "Krok 1",
        "Poprzedni release / poprzedni plan",
        "Dodaj bazowy plik Excel, do którego będzie porównywany aktualny stan zamówień i wysyłek.",
    )
    prev_file = st.file_uploader(
        "Upload Previous Release",
        type=["xlsx"],
        key="previous_release_upload",
        label_visibility="collapsed",
    )
with upload_right:
    render_upload_card(
        "Krok 2",
        "Aktualny release / aktualny plan",
        "Dodaj nowy plik Excel, aby dashboard automatycznie policzył delty, alerty i zmiany procentowe.",
    )
    current_file = st.file_uploader(
        "Upload Current Release",
        type=["xlsx"],
        key="current_release_upload",
        label_visibility="collapsed",
    )

if prev_file is None and current_file is None:
    quick_cols = st.columns(3, gap="large")
    with quick_cols[0]:
        render_quick_card(
            "Czytelny dashboard porównawczy",
            "Aplikacja zestawia poprzedni i aktualny release, od razu pokazując bilans zmian, alerty oraz produkty z największym ruchem.",
        )
    with quick_cols[1]:
        render_quick_card(
            "Macierz podobna do Excela",
            "Otrzymujesz widok tabelaryczny z datami, zmianami ilości i filtrowaniem po produkcie, kierunku ruchu oraz zakresie dat.",
        )
    with quick_cols[2]:
        render_quick_card(
            "Raport gotowy do wysłania",
            "Po analizie pobierzesz CSV oraz biznesowy raport Excel z podsumowaniem KPI i kluczowymi zmianami.",
        )
    st.info("Zacznij od dodania dwóch plików Excel. Po załadowaniu obu release'ów dashboard uruchomi pełną analizę porównawczą.")
elif prev_file is None or current_file is None:
    missing_label = "poprzedni" if prev_file is None else "aktualny"
    loaded_label = "aktualny" if prev_file is None else "poprzedni"
    st.info(
        f"Plik {loaded_label} jest już dodany. Dodaj jeszcze plik {missing_label}, aby uruchomić analizę i wygenerować dashboard."
    )
else:
    try:
        prev_df, prev_meta = load_release(prev_file.getvalue(), prev_file.name)
        curr_df, curr_meta = load_release(current_file.getvalue(), current_file.name)
        result = compare_releases(prev_df, curr_df)
    except Exception as exc:
        st.error(f"Błąd: {exc}")
    else:
        filter_col, _ = st.columns([0.32, 0.68], gap="large")
        with filter_col:
            filter_state = render_filter_controls(result)

        date_basis = filter_state["date_basis"]
        selected_start_date = filter_state["selected_start_date"]
        selected_end_date = filter_state["selected_end_date"]
        full_product_summary = filter_state["full_product_summary"]
        all_products = full_product_summary["Product Label"].tolist()
        selected_products = filter_state["selected_products"]
        search_term = filter_state["search_term"]
        selected_change_directions = filter_state["selected_change_directions"]
        only_alerts = filter_state["only_alerts"]

        filtered_df = result.copy()
        filtered_df = filtered_df[
            filtered_df[date_basis].dt.date.between(
                selected_start_date, selected_end_date
            )
        ]

        if selected_products:
            filtered_df = filtered_df[filtered_df["Product Label"].isin(selected_products)]
        else:
            filtered_df = filtered_df.iloc[0:0]

        if search_term.strip():
            query = search_term.strip().lower()
            filtered_df = filtered_df[
                filtered_df["Part Number"].str.lower().str.contains(query, na=False)
                | filtered_df["Part Description"].str.lower().str.contains(query, na=False)
            ]

        filtered_df = filtered_df[
            filtered_df["Change Direction"].isin(selected_change_directions)
        ]

        if only_alerts:
            filtered_df = filtered_df[filtered_df["Alert"]]

        product_summary = summarize_products(filtered_df)
        date_summary = summarize_dates(filtered_df, date_basis)
        weekly_summary = build_weekly_summary(
            filtered_df,
            date_basis,
            selected_start_date,
            selected_end_date,
            selected_end_date,
            THRESHOLD,
        )
        key_findings = build_key_findings(
            filtered_df, product_summary, date_summary, date_basis
        )

        st.success("Analiza porównawcza jest gotowa.")

        hero_left, hero_right = st.columns([1.8, 1], gap="large")
        with hero_left:
            hero_logo_html = (
                f'<img class="hero-logo" src="{logo_data_uri()}" alt="{BRAND_NAME} logo" />'
                if logo_available()
                else f'<div class="brand-badge">{BRAND_NAME}</div>'
            )
            st.markdown(
                f"""
                <div class="hero-card">
                    {hero_logo_html}
                    <div class="hero-kicker">Release Intelligence</div>
                    <div class="hero-title">Raport zmian dla PO {curr_meta['po_number']}</div>
                    <p class="hero-copy">
                        Porównaj wersje release'ów, śledź ruch dzień po dniu i szybko
                        wychwyć produkty, które zwiększyły lub zmniejszyły wolumen w wybranym oknie.
                    </p>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with hero_right:
            st.markdown(
                f"""
                <div class="hero-card">
                    <div class="hero-kicker">Aktywne okno analizy</div>
                    <div class="hero-title">{selected_start_date.strftime('%Y-%m-%d')}</div>
                    <p class="hero-copy">
                        do {selected_end_date.strftime('%Y-%m-%d')} na osi <strong>{get_date_label(date_basis)}</strong>
                    </p>
                    <div class="hero-stat-grid">
                        <div class="hero-stat">
                            <div class="hero-stat-label">Poprzedni release</div>
                            <div class="hero-stat-value">{format_release_label(prev_meta)}</div>
                        </div>
                        <div class="hero-stat">
                            <div class="hero-stat-label">Aktualny release</div>
                            <div class="hero-stat-value">{format_release_label(curr_meta)}</div>
                        </div>
                        <div class="hero-stat">
                            <div class="hero-stat-label">Poprzedni plik</div>
                            <div class="hero-stat-value">{prev_meta['file_name']}</div>
                        </div>
                        <div class="hero-stat">
                            <div class="hero-stat-label">Aktualny plik</div>
                            <div class="hero-stat-value">{curr_meta['file_name']}</div>
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        header_left, header_right = st.columns([1.5, 1], gap="large")
        with header_left:
            render_meta_card(
                "Kontekst release'u",
                [
                    f"<strong>Numer PO:</strong> {curr_meta['po_number']}",
                    (
                        f"<strong>Poprzedni release:</strong> {format_release_summary(prev_meta)}"
                    ),
                    (
                        f"<strong>Aktualny release:</strong> {format_release_summary(curr_meta)}"
                    ),
                ],
            )
        with header_right:
            render_meta_card(
                "Planista",
                [
                    f"<strong>Planista:</strong> {curr_meta['planner_name']}",
                    f"<strong>Email:</strong> {curr_meta['planner_email']}",
                    (
                        f"<strong>Produkty w zakresie:</strong> "
                        f"{product_summary['Part Number'].nunique()}"
                    ),
                ],
            )

        if filtered_df.empty:
            st.warning(
                "Po zastosowaniu filtrów nie ma danych do pokazania. "
                "Poszerz zakres dat albo przywróć produkty w filtrach bocznych."
            )
        else:
            total_prev = filtered_df["Quantity_Prev"].sum()
            total_curr = filtered_df["Quantity_Curr"].sum()
            total_delta = filtered_df["Delta"].sum()
            alert_count = int(filtered_df["Alert"].sum())
            products_changed = int((product_summary["Delta"] != 0).sum())

            render_status_pills(total_delta, alert_count, products_changed)
            metric_cols = st.columns(5)
            metric_cols[0].metric("Poprzednia ilość", f"{total_prev:,.0f}")
            metric_cols[1].metric(
                "Aktualna ilość",
                f"{total_curr:,.0f}",
                delta=f"{total_curr - total_prev:+,.0f}",
            )
            metric_cols[2].metric("Bilans zmian", f"{total_delta:+,.0f}")
            metric_cols[3].metric(
                "Liczba alertów",
                f"{alert_count:,}",
                delta=f"{(alert_count / len(filtered_df)):.1%}",
                delta_color="inverse",
            )
            metric_cols[4].metric("Zmienne produkty", f"{products_changed:,}")

            reference_week = get_last_completed_reference_week(selected_end_date)
            reference_row, previous_week_row = get_reference_week_rows(weekly_summary)
            reference_week_label = (
                reference_row["Week Label"] if reference_row is not None else reference_week.week_label
            )
            reference_range_label = (
                format_week_range(reference_row["Week Start"], reference_row["Week End"])
                if reference_row is not None
                else format_week_range(reference_week.week_start, reference_week.week_end)
            )
            reference_release_delta = (
                format_signed_int(reference_row["Delta"]) if reference_row is not None else "+0"
            )
            reference_release_pct = (
                format_percent_display(reference_row["Release Percent Label"])
                if reference_row is not None
                else "n/a"
            )
            reference_wow_delta = (
                format_signed_int(reference_row["WoW Delta"]) if reference_row is not None else "+0"
            )
            reference_wow_pct = (
                format_percent_display(reference_row["WoW Percent Label"])
                if reference_row is not None
                else "n/a"
            )
            reference_working_days = (
                int(reference_row["Working_Days_PL"]) if reference_row is not None else 0
            )
            reference_per_day = (
                "n/a"
                if reference_row is None or pd.isna(reference_row["Avg Current / Working Day"])
                else f"{float(reference_row['Avg Current / Working Day']):,.2f} / dzien"
            )
            previous_week_label = (
                previous_week_row["Week Label"] if previous_week_row is not None else "brak"
            )

            st.caption(
                f"Analiza tygodniowa odnosi sie do {reference_week_label} ({reference_range_label}). "
                f"Data referencyjna: {selected_end_date:%Y-%m-%d}. "
                + (
                    "Poniewaz data koncowa wypada w trakcie tygodnia, jako referencje przyjeto ostatni pelny zakonczony tydzien ISO."
                    if selected_end_date.weekday() != 6
                    else "Poniewaz data koncowa wypada w niedziele, ten tydzien zostal uznany za pelny zakonczony tydzien ISO."
                )
            )

            weekly_metric_cols = st.columns(5)
            weekly_metric_cols[0].metric(
                "Referencyjny tydzien ISO",
                reference_week_label,
                delta=reference_range_label,
            )
            weekly_metric_cols[1].metric(
                "Aktualny wolumen tygodnia",
                f"{float(reference_row['Quantity_Curr']):,.0f}" if reference_row is not None else "0",
                delta=reference_release_delta,
            )
            weekly_metric_cols[2].metric(
                "Zmiana vs poprzedni release",
                reference_release_pct,
                delta=f"prev {float(reference_row['Quantity_Prev']):,.0f}" if reference_row is not None else "prev 0",
            )
            weekly_metric_cols[3].metric(
                "Zmiana WoW",
                reference_wow_delta,
                delta=f"{reference_wow_pct} vs {previous_week_label}",
            )
            weekly_metric_cols[4].metric(
                "Dni robocze PL",
                f"{reference_working_days}",
                delta=reference_per_day,
            )

            st.markdown(
                """
                <div class="section-banner">
                    <div class="section-kicker">Executive Summary</div>
                    <div class="section-copy">
                        Najważniejsze sygnały, które warto sprawdzić w pierwszej kolejności.
                        Duży nagłówek pokazuje nazwę produktu, a krótki opis pod spodem wyjaśnia znaczenie zmiany.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.subheader("Kluczowe wnioski")
            finding_cols = st.columns(max(1, min(len(key_findings), 4)), gap="large")
            for idx, finding in enumerate(key_findings):
                with finding_cols[idx]:
                    render_finding_card(
                        finding["label"], finding["title"], finding["copy"]
                    )

            dashboard_tab, weekly_tab, product_tab, matrix_tab, detail_tab = st.tabs(
                ["Dashboard", "Analiza tygodniowa", "Raport produktu", "Macierz release'u", "Dane szczegółowe"]
            )

            with dashboard_tab:
                st.subheader(f"Trend zmian według osi: {get_date_label(date_basis)}")
                render_chart_table_switch(
                    "dashboard_trend",
                    build_quantity_chart(date_summary, get_date_label(date_basis)),
                    date_summary,
                    table_height=360,
                )

                trend_left, trend_right = st.columns([1.45, 1], gap="large")
                with trend_left:
                    render_chart_table_switch(
                        "dashboard_delta",
                        build_delta_chart(date_summary, get_date_label(date_basis)),
                        date_summary,
                        table_height=320,
                    )
                with trend_right:
                    st.subheader("Struktura zmian")
                    render_chart_table_switch(
                        "dashboard_mix",
                        build_change_mix_chart(filtered_df),
                        build_change_mix_source(filtered_df),
                        table_height=240,
                    )

                increase_chart, increase_title = build_product_bar_chart(
                    product_summary, "increase"
                )
                decrease_chart, decrease_title = build_product_bar_chart(
                    product_summary, "decrease"
                )
                dashboard_left, dashboard_right = st.columns(2)

                with dashboard_left:
                    st.subheader(increase_title)
                    if increase_chart is None:
                        st.info("Brak produktów ze wzrostem w aktualnym filtrowaniu.")
                    else:
                        render_chart_table_switch(
                            "dashboard_increase",
                            increase_chart,
                            build_product_bar_source(product_summary, "increase"),
                            table_height=340,
                        )

                with dashboard_right:
                    st.subheader(decrease_title)
                    if decrease_chart is None:
                        st.info("Brak produktów ze spadkiem w aktualnym filtrowaniu.")
                    else:
                        render_chart_table_switch(
                            "dashboard_decrease",
                            decrease_chart,
                            build_product_bar_source(product_summary, "decrease"),
                            table_height=340,
                        )

                st.subheader("Najważniejsze zmiany")
                highlight_table = (
                    product_summary.assign(Abs_Delta=product_summary["Delta"].abs())
                    .sort_values("Abs_Delta", ascending=False)
                    .drop(columns=["Abs_Delta"])
                    .head(10)
                )
                highlight_table["Quantity_Prev"] = highlight_table["Quantity_Prev"].map(
                    lambda value: f"{value:,.0f}"
                )
                highlight_table["Quantity_Curr"] = highlight_table["Quantity_Curr"].map(
                    lambda value: f"{value:,.0f}"
                )
                highlight_table["Delta"] = highlight_table["Delta"].map(format_signed_int)
                highlight_table = highlight_table.rename(
                    columns={
                        "Part Number": "Numer części",
                        "Part Description": "Opis produktu",
                        "Quantity_Prev": "Poprzednia ilość",
                        "Quantity_Curr": "Aktualna ilość",
                        "Delta": "Zmiana ilości",
                        "Alert_Count": "Liczba alertów",
                        "Change Direction": "Kierunek zmiany",
                    }
                )
                st.dataframe(highlight_table, use_container_width=True, height=360)

                st.subheader("Tygodnie ISO")
                weekly_chart = build_weekly_quantity_chart(weekly_summary)
                weekly_preview = prepare_weekly_display_table(weekly_summary).tail(8)
                render_chart_table_switch(
                    "dashboard_weekly",
                    weekly_chart,
                    weekly_preview,
                    chart_empty_message="Brak danych tygodniowych do wykresu.",
                    table_height=320,
                )

            with weekly_tab:
                st.subheader("Analiza tygodniowa oparta na datach")
                weekly_partial = weekly_summary[
                    weekly_summary["Is Partial Range"] | ~weekly_summary["Is Closed Week"]
                ]
                if not weekly_partial.empty:
                    st.info(
                        "W tabeli i wykresach tygodnie oznaczone jako 'Partial range' lub 'Open week' "
                        "obejmują niepełny zakres albo nie były jeszcze zakończone względem daty referencyjnej."
                    )

                weekly_qty_chart = build_weekly_quantity_chart(weekly_summary)
                render_chart_table_switch(
                    "weekly_quantity",
                    weekly_qty_chart,
                    prepare_weekly_display_table(weekly_summary),
                    chart_empty_message="Brak danych tygodniowych do wykresu.",
                    table_height=360,
                )

                weekly_left, weekly_right = st.columns([1.3, 1], gap="large")
                with weekly_left:
                    weekly_delta_chart = build_weekly_delta_chart(weekly_summary)
                    render_chart_table_switch(
                        "weekly_delta",
                        weekly_delta_chart,
                        prepare_weekly_display_table(weekly_summary),
                        chart_empty_message="Brak danych tygodniowych do wykresu delta.",
                        table_height=320,
                    )
                with weekly_right:
                    weekly_focus = pd.DataFrame(
                        [
                            {
                                "Widok": "Referencyjny tydzien",
                                "Tydzien ISO": reference_week_label,
                                "Aktualny release": (
                                    f"{float(reference_row['Quantity_Curr']):,.0f}"
                                    if reference_row is not None
                                    else "0"
                                ),
                                "Poprzedni release": (
                                    f"{float(reference_row['Quantity_Prev']):,.0f}"
                                    if reference_row is not None
                                    else "0"
                                ),
                                "Delta release": reference_release_delta,
                                "Zmiana release %": reference_release_pct,
                                "Delta WoW": reference_wow_delta,
                                "Zmiana WoW %": reference_wow_pct,
                            },
                            {
                                "Widok": "Poprzedni tydzien",
                                "Tydzien ISO": previous_week_label,
                                "Aktualny release": (
                                    f"{float(previous_week_row['Quantity_Curr']):,.0f}"
                                    if previous_week_row is not None
                                    else "0"
                                ),
                                "Poprzedni release": (
                                    f"{float(previous_week_row['Quantity_Prev']):,.0f}"
                                    if previous_week_row is not None
                                    else "0"
                                ),
                                "Delta release": (
                                    format_signed_int(previous_week_row["Delta"])
                                    if previous_week_row is not None
                                    else "+0"
                                ),
                                "Zmiana release %": (
                                    format_percent_display(previous_week_row["Release Percent Label"])
                                    if previous_week_row is not None
                                    else "n/a"
                                ),
                                "Delta WoW": (
                                    format_signed_int(previous_week_row["WoW Delta"])
                                    if previous_week_row is not None
                                    else "+0"
                                ),
                                "Zmiana WoW %": (
                                    format_percent_display(previous_week_row["WoW Percent Label"])
                                    if previous_week_row is not None
                                    else "n/a"
                                ),
                            },
                        ]
                    )
                    st.subheader("Porównanie tygodni")
                    st.dataframe(weekly_focus, use_container_width=True, height=240)

                weekly_table = prepare_weekly_display_table(weekly_summary)
                st.subheader("Tabela tygodniowa")
                st.dataframe(weekly_table, use_container_width=True, height=420)

            with product_tab:
                st.subheader("Analiza wybranego produktu")
                selected_product_label = st.selectbox(
                    "Wybierz produkt",
                    options=product_summary["Product Label"].tolist(),
                )
                product_detail = filtered_df[
                    filtered_df["Product Label"] == selected_product_label
                ].sort_values(date_basis)
                product_date_summary = summarize_dates(product_detail, date_basis)

                product_metrics = st.columns(4)
                product_metrics[0].metric(
                    "Poprzednia ilość", f"{product_detail['Quantity_Prev'].sum():,.0f}"
                )
                product_metrics[1].metric(
                    "Aktualna ilość", f"{product_detail['Quantity_Curr'].sum():,.0f}"
                )
                product_metrics[2].metric(
                    "Bilans zmian", f"{product_detail['Delta'].sum():+,.0f}"
                )
                product_metrics[3].metric(
                    "Liczba alertów", int(product_detail["Alert"].sum())
                )

                render_chart_table_switch(
                    "product_quantity",
                    build_quantity_chart(product_date_summary, get_date_label(date_basis)),
                    product_date_summary,
                    table_height=320,
                )
                render_chart_table_switch(
                    "product_delta",
                    build_delta_chart(product_date_summary, get_date_label(date_basis)),
                    product_date_summary,
                    table_height=320,
                )

                product_weekly_summary = build_weekly_summary(
                    product_detail,
                    date_basis,
                    selected_start_date,
                    selected_end_date,
                    selected_end_date,
                    THRESHOLD,
                )
                st.subheader("Tygodnie ISO dla produktu")
                product_weekly_chart = build_weekly_quantity_chart(product_weekly_summary)
                render_chart_table_switch(
                    "product_weekly",
                    product_weekly_chart,
                    prepare_weekly_display_table(product_weekly_summary),
                    chart_empty_message="Brak danych tygodniowych dla wybranego produktu.",
                    table_height=280,
                )

                product_table = product_detail[available_detail_columns(product_detail)].copy()
                product_table["Ship Date"] = product_table["Ship Date"].dt.strftime("%Y-%m-%d")
                product_table["Receipt Date"] = product_table["Receipt Date"].dt.strftime(
                    "%Y-%m-%d"
                )
                product_table["Change Direction"] = product_table["Change Direction"].map(
                    get_change_label
                )
                product_table["Alert"] = product_table["Alert"].map(
                    lambda value: "Tak" if value else "Nie"
                )
                product_table = product_table.rename(
                    columns={
                        "Part Number": "Numer części",
                        "Part Description": "Opis produktu",
                        "Origin Doc": "Origin Doc",
                        "Item": "Pozycja",
                        "Ship To": "Ship-to",
                        "Customer Material": "Materiał klienta",
                        "Unrestricted Qty": "Ilość unrestr.",
                        "Unloading Point": "Punkt rozładunku",
                        "Ship Date": "Data wysyłki",
                        "Receipt Date": "Data odbioru",
                        "Unit of Measure": "JM",
                        "CumQty": "CumQty",
                        "Quantity_Prev": "Poprzednia ilość",
                        "Quantity_Curr": "Aktualna ilość",
                        "Delta": "Zmiana ilości",
                        "Percent Change": "Zmiana %",
                        "Demand Status": "Status popytu",
                        "Change Direction": "Kierunek zmiany",
                        "Alert": "Alert",
                    }
                )
                st.dataframe(product_table, use_container_width=True, height=360)

            with matrix_tab:
                st.subheader("Macierz podobna do arkusza release'u")
                matrix_metric = st.radio(
                    "Metryka",
                    options=["Current Quantity", "Previous Quantity", "Delta", "Percent Change"],
                    horizontal=True,
                    format_func=get_metric_label,
                )
                matrix = build_matrix(filtered_df, date_basis, matrix_metric)
                matrix_cells = matrix.shape[0] * max(matrix.shape[1], 1)

                if matrix.empty:
                    st.info("Brak danych do macierzy.")
                elif matrix_cells <= MAX_MATRIX_STYLE_CELLS:
                    st.dataframe(
                        style_matrix(matrix, matrix_metric),
                        use_container_width=True,
                        height=520,
                    )
                else:
                    st.info(
                        "Macierz jest zbyt duza do stylowania, dlatego pokazuje ja "
                        "bez dodatkowego formatowania."
                    )
                    st.dataframe(matrix, use_container_width=True, height=520)

            with detail_tab:
                st.subheader("Dane szczegółowe")
                preview_limit = st.selectbox(
                    "Liczba wierszy w podglądzie",
                    options=[100, 250, 500, 1000],
                    index=2,
                )
                detail_table = filtered_df[available_detail_columns(filtered_df)].copy()
                detail_table["Ship Date"] = detail_table["Ship Date"].dt.strftime("%Y-%m-%d")
                detail_table["Receipt Date"] = detail_table["Receipt Date"].dt.strftime(
                    "%Y-%m-%d"
                )
                detail_table["Change Direction"] = detail_table["Change Direction"].map(
                    get_change_label
                )
                detail_table["Alert"] = detail_table["Alert"].map(
                    lambda value: "Tak" if value else "Nie"
                )
                detail_table = detail_table.rename(
                    columns={
                        "PO Number": "Numer PO",
                        "Origin Doc": "Origin Doc",
                        "Item": "Pozycja",
                        "Ship To": "Ship-to",
                        "Part Number": "Numer części",
                        "Part Description": "Opis produktu",
                        "Customer Material": "Materiał klienta",
                        "Unrestricted Qty": "Ilość unrestr.",
                        "Unloading Point": "Punkt rozładunku",
                        "Ship Date": "Data wysyłki",
                        "Receipt Date": "Data odbioru",
                        "Unit of Measure": "JM",
                        "CumQty": "CumQty",
                        "Quantity_Prev": "Poprzednia ilość",
                        "Quantity_Curr": "Aktualna ilość",
                        "Delta": "Zmiana ilości",
                        "Percent Change": "Zmiana %",
                        "Demand Status": "Status popytu",
                        "Change Direction": "Kierunek zmiany",
                        "Alert": "Alert",
                    }
                )

                if len(detail_table) > preview_limit:
                    st.info(
                        f"Pokazuje pierwsze {preview_limit} z {len(detail_table)} wierszy. "
                        "Pełny raport jest dostępny do pobrania."
                    )
                st.dataframe(
                    detail_table.head(preview_limit),
                    use_container_width=True,
                    height=420,
                )

                current_matrix_for_export = build_matrix(
                    filtered_df, date_basis, "Current Quantity"
                )
                delta_matrix_for_export = build_matrix(filtered_df, date_basis, "Delta")
                excel_bytes = to_excel_bytes(
                    filtered_df,
                    weekly_summary,
                    current_matrix_for_export,
                    delta_matrix_for_export,
                    prev_meta,
                    curr_meta,
                    product_summary,
                    date_basis,
                    selected_start_date,
                    selected_end_date,
                    key_findings,
                )
                csv_bytes = detail_table.to_csv(index=False).encode("utf-8")

                download_left, download_right = st.columns(2)
                with download_left:
                    st.download_button(
                        "Pobierz filtrowane dane CSV",
                        data=csv_bytes,
                        file_name="pjoter_development_release_change_filtered.csv",
                        mime="text/csv",
                    )
                with download_right:
                    st.download_button(
                        "Pobierz raport Excel",
                        data=excel_bytes,
                        file_name="pjoter_development_release_change_report.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ),
                    )
