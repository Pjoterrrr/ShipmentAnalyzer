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
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


THRESHOLD = 15
MAX_MATRIX_STYLE_CELLS = 50000
BRAND_NAME = "Pjoter Development"
BASE_DIR = Path(__file__).resolve().parent
RUNTIME_ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else BASE_DIR


def resolve_runtime_path(relative_path):
    external_path = RUNTIME_ROOT / relative_path
    internal_path = BASE_DIR / relative_path
    return external_path if external_path.exists() else internal_path


LOGO_PATH = resolve_runtime_path(Path("assets") / "logo.png")
AUTH_USERS_PATH = resolve_runtime_path(Path("config") / "users.json")
REQUIRED_RAW_COLUMNS = [
    "PO Number",
    "PO Line #",
    "Release Version",
    "Release Date",
    "Part Number",
    "Part Description",
    "Ship Date",
    "Receipt Date",
    "Open Quantity",
    "Unit of Measure",
]
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


st.set_page_config(
    page_title="Pjoter Development | Analiza zamówień i wysyłek",
    layout="wide",
    initial_sidebar_state="expanded",
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
        border-right: 1px solid var(--line);
        background:
            linear-gradient(180deg, rgba(6,10,16,0.98), rgba(10,15,24,0.96)),
            radial-gradient(circle at top, rgba(76, 201, 240, 0.08), transparent 26%);
        backdrop-filter: blur(18px);
    }
    section[data-testid="stSidebar"] .stRadio > label,
    section[data-testid="stSidebar"] .stMultiSelect label,
    section[data-testid="stSidebar"] .stTextInput label,
    section[data-testid="stSidebar"] .stDateInput label {
        color: var(--ink);
        font-weight: 700;
        letter-spacing: 0.01em;
    }
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
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
    }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown(
    """
    <div class="section-banner">
        <div class="section-kicker">Premium Release Intelligence</div>
        <div class="section-copy">
            Prześlij poprzedni i aktualny release. Dashboard porówna zamówienia po dacie wysyłki
            lub odbioru, pokaże wzrosty, spadki, alerty oraz gotowy raport do dalszej pracy.
        </div>
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


def load_auth_config():
    if not AUTH_USERS_PATH.exists():
        raise FileNotFoundError(
            f"Brakuje pliku konfiguracyjnego uzytkownikow: {AUTH_USERS_PATH}"
        )
    with AUTH_USERS_PATH.open("r", encoding="utf-8") as file:
        payload = json.load(file)
    return payload.get("users", [])


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


def logo_data_uri():
    if not logo_available():
        return ""
    encoded = base64.b64encode(LOGO_PATH.read_bytes()).decode("utf-8")
    return f"data:image/png;base64,{encoded}"


def render_sidebar_user():
    auth_user = st.session_state.get("auth_user") or {}
    st.sidebar.markdown(
        f"""
        <div class="sidebar-user-card">
            <div class="sidebar-user-label">Aktywna sesja</div>
            <div class="sidebar-user-name">{html.escape(str(auth_user.get('display_name', 'User')))}</div>
            <div class="sidebar-user-role">{html.escape(str(auth_user.get('role', 'User')))} &middot; {html.escape(str(auth_user.get('username', '')))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.sidebar.button("Wyloguj", use_container_width=True):
        logout_user()
        st.rerun()


def apply_chart_theme(chart):
    return (
        chart.configure_view(strokeOpacity=0)
        .configure(background="transparent")
        .configure_axis(
            gridColor="#243244",
            domainColor="#1b2635",
            tickColor="#243244",
            labelColor="#9fb0c9",
            titleColor="#eef4ff",
            labelFontSize=12,
            titleFontSize=13,
            gridDash=[2, 4],
            labelPadding=8,
            titlePadding=14,
        )
        .configure_legend(
            labelColor="#b5c0d4",
            titleColor="#eef4ff",
            labelFontSize=12,
            titleFontSize=13,
            symbolType="circle",
        )
        .configure_title(color="#eef4ff", fontSize=16, fontWeight="bold", anchor="start")
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


@st.cache_data(show_spinner=False)
def load_release(file_bytes, file_name):
    raw_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Raw")
    raw_df.columns = [str(column).replace("? ", "").strip() for column in raw_df.columns]

    missing_columns = [column for column in REQUIRED_RAW_COLUMNS if column not in raw_df.columns]
    if missing_columns:
        raise ValueError(
            "Raw sheet is missing required columns: " + ", ".join(missing_columns)
        )

    overview_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=4)

    raw_df["PO Number"] = raw_df["PO Number"].astype(str).str.strip()
    raw_df["Part Number"] = raw_df["Part Number"].astype(str).str.strip()
    raw_df["Part Description"] = raw_df["Part Description"].astype(str).str.strip()
    raw_df["Unit of Measure"] = raw_df["Unit of Measure"].astype(str).str.strip()
    raw_df["Release Date"] = pd.to_datetime(raw_df["Release Date"], errors="coerce")
    raw_df["Ship Date"] = pd.to_datetime(raw_df["Ship Date"], errors="coerce")
    raw_df["Receipt Date"] = pd.to_datetime(raw_df["Receipt Date"], errors="coerce")
    raw_df["Open Quantity"] = pd.to_numeric(raw_df["Open Quantity"], errors="coerce").fillna(0)
    raw_df = raw_df.dropna(
        subset=["Part Number", "Part Description", "Ship Date", "Receipt Date"]
    ).copy()

    metadata = {
        "file_name": file_name,
        "po_number": first_non_empty(raw_df["PO Number"]),
        "release_version": first_non_empty(raw_df["Release Version"]),
        "release_date": raw_df["Release Date"].dropna().min(),
        "planner_name": (
            first_non_empty(overview_df["Planner Name"])
            if "Planner Name" in overview_df.columns
            else "n/a"
        ),
        "planner_email": (
            first_non_empty(overview_df["Planner Email"])
            if "Planner Email" in overview_df.columns
            else "n/a"
        ),
        "products": raw_df["Part Number"].nunique(),
        "rows": len(raw_df),
    }
    return raw_df, metadata


def compare_releases(prev_df, curr_df):
    keys = ["Part Number", "Part Description", "Ship Date", "Receipt Date"]

    prev_summary = prev_df.groupby(keys, as_index=False).agg(
        Quantity_Prev=("Open Quantity", "sum"),
        UoM_Prev=("Unit of Measure", "first"),
        PO_Prev=("PO Number", "first"),
    )
    curr_summary = curr_df.groupby(keys, as_index=False).agg(
        Quantity_Curr=("Open Quantity", "sum"),
        UoM_Curr=("Unit of Measure", "first"),
        PO_Curr=("PO Number", "first"),
    )

    merged = prev_summary.merge(curr_summary, on=keys, how="outer")
    merged["Quantity_Prev"] = merged["Quantity_Prev"].fillna(0)
    merged["Quantity_Curr"] = merged["Quantity_Curr"].fillna(0)
    merged["Unit of Measure"] = merged["UoM_Prev"].combine_first(merged["UoM_Curr"])
    merged["PO Number"] = merged["PO_Prev"].combine_first(merged["PO_Curr"])
    merged["Delta"] = merged["Quantity_Curr"] - merged["Quantity_Prev"]
    merged["Abs Delta"] = merged["Delta"].abs()
    merged["Percent Change"] = merged.apply(
        lambda row: 100
        if row["Quantity_Prev"] == 0 and row["Quantity_Curr"] > 0
        else (
            0
            if row["Quantity_Prev"] == 0
            else round((row["Delta"] / row["Quantity_Prev"]) * 100, 2)
        ),
        axis=1,
    )
    merged["Alert"] = merged["Percent Change"].abs() >= THRESHOLD
    merged["Change Direction"] = merged["Delta"].apply(
        lambda value: "Increase" if value > 0 else ("Decrease" if value < 0 else "No Change")
    )
    merged["Product Label"] = (
        merged["Part Number"].astype(str) + " | " + merged["Part Description"].astype(str)
    )
    merged = merged.drop(columns=["UoM_Prev", "UoM_Curr", "PO_Prev", "PO_Curr"])
    return merged.sort_values(["Receipt Date", "Ship Date", "Part Description"]).reset_index(
        drop=True
    )


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


def build_quantity_chart(date_summary, x_title):
    chart_data = date_summary.sort_values("Analysis Date").copy()
    latest_point = chart_data.tail(1).copy()
    latest_point["Current Label"] = latest_point["Quantity_Curr"].map(lambda value: f"Aktualnie {value:,.0f}")
    latest_point["Previous Label"] = latest_point["Quantity_Prev"].map(lambda value: f"Poprzednio {value:,.0f}")

    prev_line = (
        alt.Chart(chart_data)
        .mark_line(strokeWidth=2.2, interpolate="monotone", color="#8f9caf", strokeDash=[5, 4])
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
        .mark_area(color="#6b89b6", opacity=0.12, interpolate="monotone")
        .encode(
            x=alt.X("Analysis Date:T", title=x_title, axis=alt.Axis(labelAngle=-24, labelLimit=140)),
            y=alt.Y("Quantity_Curr:Q", title="Ilość otwarta"),
        )
    )
    current_line = (
        alt.Chart(chart_data)
        .mark_line(
            point=alt.OverlayMarkDef(size=72, filled=True, fill="#0f1722", stroke="#7f93b1", strokeWidth=1.8),
            strokeWidth=3.0,
            interpolate="monotone",
            color="#7f93b1",
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
            align="right",
            baseline="bottom",
            dx=-8,
            dy=-12,
            color="#eef4ff",
            fontSize=12,
            fontWeight="bold",
        )
        .encode(x="Analysis Date:T", y="Quantity_Curr:Q", text="Current Label:N")
    )
    previous_label = (
        alt.Chart(latest_point)
        .mark_text(
            align="right",
            baseline="top",
            dx=-8,
            dy=12,
            color="#93a4bd",
            fontSize=11,
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


def build_product_bar_chart(product_summary, chart_type):
    if chart_type == "increase":
        source = (
            product_summary[product_summary["Delta"] > 0]
            .nlargest(10, "Delta")[["Part Description", "Delta"]]
        )
        color = "#5f8b75"
        title = "Największe wzrosty"
    else:
        source = (
            product_summary[product_summary["Delta"] < 0]
            .nsmallest(10, "Delta")[["Part Description", "Delta"]]
        )
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


def build_change_mix_chart(dataframe):
    mix = (
        dataframe.groupby("Change Direction", as_index=False)
        .agg(Rows=("Change Direction", "size"), Total_Delta=("Delta", "sum"))
    )
    mix["Direction Label"] = mix["Change Direction"].map(get_change_label)
    mix["Share"] = mix["Rows"] / max(int(mix["Rows"].sum()), 1)
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
        if percent_column is not None:
            percent_cell = worksheet.cell(row=row, column=percent_column)
            if isinstance(percent_cell.value, (int, float)):
                percent_cell.number_format = '0.0"%"'


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
                fill_css = style_value(cell.value, metric_name, max_value, max_abs)
                bg = fill_css.split("background-color: ")[1].split(";")[0].replace("#", "")
                fg = fill_css.split("color: ")[1].split(";")[0].replace("#", "")
                cell.fill = PatternFill(fill_type="solid", fgColor=bg)
                cell.font = Font(color=fg, bold=False)
                if metric_name == "Percent Change":
                    cell.number_format = '0.0"%"'
                else:
                    cell.number_format = '#,##0'


def write_summary_sheet(
    worksheet,
    prev_meta,
    curr_meta,
    detail_df,
    product_summary,
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
    worksheet["B5"] = f"v{prev_meta['release_version']} / {format_date(prev_meta['release_date'])}"
    worksheet["A6"] = "Current Release"
    worksheet["B6"] = f"v{curr_meta['release_version']} / {format_date(curr_meta['release_date'])}"
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
    current_matrix_export = current_matrix_df.reset_index()
    delta_matrix_export = delta_matrix_df.reset_index()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, sheet_name="Executive Summary", index=False)
        detail_export.to_excel(writer, sheet_name="Detailed Data", index=False)
        current_matrix_export.to_excel(writer, sheet_name="Current Matrix", index=False)
        delta_matrix_export.to_excel(writer, sheet_name="Delta Heatmap", index=False)

        summary_sheet = writer.book["Executive Summary"]
        write_summary_sheet(
            summary_sheet,
            prev_meta,
            curr_meta,
            detail_df,
            product_summary,
            date_basis,
            selected_start_date,
            selected_end_date,
            key_findings,
        )

        detail_sheet = writer.book["Detailed Data"]
        style_excel_header(detail_sheet, 1)
        decorate_delta_column(detail_sheet, header_row=1)
        detail_sheet.freeze_panes = "A2"
        autosize_worksheet(detail_sheet)

        current_matrix_sheet = writer.book["Current Matrix"]
        style_excel_header(current_matrix_sheet, 1)
        current_matrix_sheet.freeze_panes = "B2"
        autosize_worksheet(current_matrix_sheet)
        style_matrix_sheet(current_matrix_sheet, "Current Quantity")

        delta_heatmap_sheet = writer.book["Delta Heatmap"]
        style_excel_header(delta_heatmap_sheet, 1)
        delta_heatmap_sheet.freeze_panes = "B2"
        autosize_worksheet(delta_heatmap_sheet)
        style_matrix_sheet(delta_heatmap_sheet, "Delta")

    return output.getvalue()


init_auth_state()

if not st.session_state["authenticated"]:
    render_login_screen()
    st.stop()

render_sidebar_user()

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
        st.sidebar.header("Filtry")
        if logo_available():
            st.sidebar.image(str(LOGO_PATH), use_container_width=True)
        date_basis = st.sidebar.radio(
            "Oś dat",
            DATE_OPTIONS,
            index=0,
            format_func=get_date_label,
        )

        full_product_summary = summarize_products(result)
        all_products = full_product_summary["Product Label"].tolist()
        selected_products = st.sidebar.multiselect(
            "Produkty",
            options=all_products,
            default=all_products,
        )
        search_term = st.sidebar.text_input("Szukaj po numerze lub opisie")
        selected_change_directions = st.sidebar.multiselect(
            "Kierunek zmiany",
            options=["Increase", "Decrease", "No Change"],
            default=["Increase", "Decrease", "No Change"],
            format_func=get_change_label,
        )
        only_alerts = st.sidebar.checkbox(f"Tylko alerty >= {THRESHOLD}%")

        available_dates = result[date_basis].dropna().sort_values()
        min_date = available_dates.min().date()
        max_date = available_dates.max().date()
        selected_date_input = st.date_input(
            f"Kalendarz: {get_date_label(date_basis)}",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            help="Wybierz jeden dzień albo zakres dat do analizy.",
        )
        selected_start_date, selected_end_date = normalize_date_selection(
            selected_date_input, min_date, max_date
        )
        if selected_start_date > selected_end_date:
            selected_start_date, selected_end_date = (
                selected_end_date,
                selected_start_date,
            )

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
                            <div class="hero-stat-value">v{prev_meta['release_version']}</div>
                        </div>
                        <div class="hero-stat">
                            <div class="hero-stat-label">Aktualny release</div>
                            <div class="hero-stat-value">v{curr_meta['release_version']}</div>
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
                        f"<strong>Poprzedni release:</strong> v{prev_meta['release_version']} "
                        f"({format_date(prev_meta['release_date'])})"
                    ),
                    (
                        f"<strong>Aktualny release:</strong> v{curr_meta['release_version']} "
                        f"({format_date(curr_meta['release_date'])})"
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

            dashboard_tab, product_tab, matrix_tab, detail_tab = st.tabs(
                ["Dashboard", "Raport produktu", "Macierz release'u", "Dane szczegółowe"]
            )

            with dashboard_tab:
                st.subheader(f"Trend zmian według osi: {get_date_label(date_basis)}")
                st.altair_chart(
                    build_quantity_chart(date_summary, get_date_label(date_basis)),
                    use_container_width=True,
                )

                trend_left, trend_right = st.columns([1.45, 1], gap="large")
                with trend_left:
                    st.altair_chart(
                        build_delta_chart(date_summary, get_date_label(date_basis)),
                        use_container_width=True,
                    )
                with trend_right:
                    st.subheader("Struktura zmian")
                    st.altair_chart(
                        build_change_mix_chart(filtered_df), use_container_width=True
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
                        st.altair_chart(increase_chart, use_container_width=True)

                with dashboard_right:
                    st.subheader(decrease_title)
                    if decrease_chart is None:
                        st.info("Brak produktów ze spadkiem w aktualnym filtrowaniu.")
                    else:
                        st.altair_chart(decrease_chart, use_container_width=True)

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

                st.altair_chart(
                    build_quantity_chart(product_date_summary, get_date_label(date_basis)), use_container_width=True
                )
                st.altair_chart(
                    build_delta_chart(product_date_summary, get_date_label(date_basis)), use_container_width=True
                )

                product_table = product_detail[
                    [
                        "Part Number",
                        "Part Description",
                        "Ship Date",
                        "Receipt Date",
                        "Quantity_Prev",
                        "Quantity_Curr",
                        "Delta",
                        "Percent Change",
                        "Change Direction",
                        "Alert",
                    ]
                ].copy()
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
                        "Ship Date": "Data wysyłki",
                        "Receipt Date": "Data odbioru",
                        "Quantity_Prev": "Poprzednia ilość",
                        "Quantity_Curr": "Aktualna ilość",
                        "Delta": "Zmiana ilości",
                        "Percent Change": "Zmiana %",
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
                detail_table = filtered_df[
                    [
                        "PO Number",
                        "Part Number",
                        "Part Description",
                        "Ship Date",
                        "Receipt Date",
                        "Unit of Measure",
                        "Quantity_Prev",
                        "Quantity_Curr",
                        "Delta",
                        "Percent Change",
                        "Change Direction",
                        "Alert",
                    ]
                ].copy()
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
                        "Part Number": "Numer części",
                        "Part Description": "Opis produktu",
                        "Ship Date": "Data wysyłki",
                        "Receipt Date": "Data odbioru",
                        "Unit of Measure": "JM",
                        "Quantity_Prev": "Poprzednia ilość",
                        "Quantity_Curr": "Aktualna ilość",
                        "Delta": "Zmiana ilości",
                        "Percent Change": "Zmiana %",
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
