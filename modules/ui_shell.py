from __future__ import annotations

import html
from dataclasses import dataclass

import streamlit as st


APP_TITLE = "Aplikacja Analityczna"


@dataclass(frozen=True)
class PrimaryTile:
    key: str
    title: str
    icon: str
    copy: str
    action_label: str


PRIMARY_TILES = (
    PrimaryTile(
        key="dashboard",
        title="Dashboard",
        icon="dashboard",
        copy="KPI, alerty i najwazniejsze sygnaly po wejsciu do osobnego widoku.",
        action_label="Otworz Dashboard",
    ),
    PrimaryTile(
        key="filters",
        title="Filtry",
        icon="filters",
        copy="Jedyny panel rozwijany. Kontroluje zakres, produkty i kierunek zmian.",
        action_label="Pokaz Filtry",
    ),
    PrimaryTile(
        key="files",
        title="Analiza plikow",
        icon="files",
        copy="Upload, status, planner, eksport i wyniki analizy w uporzadkowanym workspace.",
        action_label="Otworz Analize plikow",
    ),
    PrimaryTile(
        key="charts",
        title="Wykresy",
        icon="charts",
        copy="Wizualizacje i raporty prezentowane osobno, bez przeladowania ekranu glownego.",
        action_label="Otworz Wykresy",
    ),
)


def inject_styles():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@500;600;700;800&family=IBM+Plex+Sans:wght@400;500;600&display=swap');

        :root {
            --app-bg: #f5f7fb;
            --app-bg-soft: #eef3f9;
            --surface: rgba(255, 255, 255, 0.94);
            --surface-strong: #ffffff;
            --surface-muted: #f8fbff;
            --line: rgba(15, 23, 42, 0.08);
            --line-strong: rgba(15, 23, 42, 0.14);
            --text: #172033;
            --text-soft: #5b667a;
            --text-faint: #7b8699;
            --accent: #225f9b;
            --accent-soft: #eef5ff;
            --accent-strong: #163d66;
            --success: #1f8f64;
            --danger: #c85a54;
            --warning: #cf8a2e;
            --shadow-lg: 0 28px 72px rgba(18, 38, 63, 0.10);
            --shadow-md: 0 18px 42px rgba(18, 38, 63, 0.08);
            --shadow-sm: 0 10px 24px rgba(18, 38, 63, 0.06);
            --radius-xl: 28px;
            --radius-lg: 22px;
            --radius-md: 18px;
            --radius-sm: 14px;
        }

        html, body, [class*="css"] {
            font-family: "IBM Plex Sans", "Segoe UI", sans-serif !important;
            color: var(--text);
        }

        h1, h2, h3, h4, h5, h6 {
            font-family: "Manrope", "Segoe UI", sans-serif !important;
            color: var(--text);
            letter-spacing: -0.02em;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(34, 95, 155, 0.10), transparent 24%),
                radial-gradient(circle at top right, rgba(119, 151, 191, 0.10), transparent 22%),
                linear-gradient(180deg, #f8fbff 0%, var(--app-bg) 52%, #eef2f8 100%) !important;
            color: var(--text) !important;
        }

        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2.6rem;
            max-width: 1480px;
        }

        .stMarkdown,
        .stCaption,
        p,
        label,
        span {
            color: var(--text);
        }

        [data-testid="stSidebar"],
        [data-testid="collapsedControl"],
        [data-testid="stSidebarCollapseButton"],
        button[aria-label="Close sidebar"],
        button[aria-label="Open sidebar"] {
            display: none !important;
        }

        [data-testid="stExpander"] {
            background: var(--surface);
            border: 1px solid var(--line);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-sm);
            overflow: hidden;
            margin: 0 0 1.35rem 0;
        }

        [data-testid="stExpander"] details {
            border-radius: var(--radius-lg);
        }

        [data-testid="stExpander"] summary {
            background: linear-gradient(180deg, #ffffff, #f7faff);
            padding: 0.9rem 1.05rem;
        }

        [data-testid="stExpander"] summary p {
            font-family: "Manrope", "Segoe UI", sans-serif;
            font-size: 0.98rem;
            font-weight: 700;
            color: var(--text) !important;
        }

        .stAlert {
            border-radius: var(--radius-md);
            border: 1px solid var(--line);
            box-shadow: var(--shadow-sm);
        }

        .stTextInput input,
        .stDateInput input,
        .stSelectbox div[data-baseweb="select"] > div,
        .stMultiSelect div[data-baseweb="select"] > div,
        .stNumberInput input,
        .stTextArea textarea {
            border-radius: 14px !important;
            border-color: rgba(15, 23, 42, 0.12) !important;
            background: rgba(255, 255, 255, 0.95) !important;
            color: var(--text) !important;
        }

        .stRadio [role="radiogroup"],
        .stSegmentedControl {
            background: rgba(233, 240, 248, 0.86);
            border-radius: 16px;
            padding: 0.2rem;
        }

        .stButton > button,
        .stDownloadButton > button,
        .stFormSubmitButton > button {
            border-radius: 14px !important;
            border: 1px solid rgba(34, 95, 155, 0.12) !important;
            background: linear-gradient(180deg, #ffffff 0%, #f6faff 100%) !important;
            color: var(--text) !important;
            box-shadow: 0 10px 24px rgba(18, 38, 63, 0.05);
            min-height: 2.8rem;
            font-weight: 600 !important;
            transition: transform 120ms ease, box-shadow 120ms ease, border-color 120ms ease;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        .stFormSubmitButton > button:hover {
            border-color: rgba(34, 95, 155, 0.28) !important;
            box-shadow: 0 16px 30px rgba(18, 38, 63, 0.08);
            transform: translateY(-1px);
        }

        .stButton > button:focus,
        .stDownloadButton > button:focus,
        .stFormSubmitButton > button:focus {
            outline: none !important;
            border-color: rgba(34, 95, 155, 0.45) !important;
            box-shadow: 0 0 0 3px rgba(34, 95, 155, 0.12) !important;
        }

        div[class*="st-key-home_tile_"] {
            background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(247,250,255,0.96));
            border: 1px solid var(--line);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-sm);
            padding: 1.2rem 1.2rem 1rem 1.2rem;
            min-height: 100%;
            transition: transform 160ms ease, box-shadow 160ms ease, border-color 160ms ease;
        }

        div[class*="st-key-home_tile_"]:hover {
            transform: translateY(-3px);
            box-shadow: var(--shadow-md);
            border-color: rgba(34, 95, 155, 0.18);
        }

        div[class*="st-key-home_tile_"] .stButton > button {
            width: 100%;
            margin-top: 0.55rem;
        }

        .aa-hero {
            max-width: 820px;
            margin: 0 auto 2rem auto;
            text-align: center;
            padding: 2.2rem 1rem 0.4rem 1rem;
        }

        .aa-hero__logo {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-height: 148px;
            margin-bottom: 1rem;
        }

        .aa-hero__logo img {
            max-width: min(440px, 72vw);
            width: auto;
            max-height: 132px;
            object-fit: contain;
            filter: drop-shadow(0 18px 30px rgba(17, 38, 64, 0.10));
        }

        .aa-hero__title {
            font-family: "Manrope", "Segoe UI", sans-serif;
            font-size: clamp(2.2rem, 4.8vw, 3.6rem);
            font-weight: 800;
            color: var(--text);
            letter-spacing: -0.04em;
            margin: 0;
        }

        .aa-hero__copy {
            margin: 0.9rem auto 0 auto;
            max-width: 660px;
            color: var(--text-soft);
            font-size: 1rem;
            line-height: 1.65;
        }

        .aa-hero__fallback {
            width: 110px;
            height: 110px;
            border-radius: 28px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            background: linear-gradient(135deg, #225f9b, #6f8ead);
            color: #ffffff;
            font-family: "Manrope", "Segoe UI", sans-serif;
            font-size: 2.1rem;
            font-weight: 800;
            letter-spacing: -0.04em;
            box-shadow: var(--shadow-md);
        }

        .aa-shell {
            display: flex;
            justify-content: space-between;
            gap: 1.25rem;
            padding: 1.05rem 1.15rem;
            border-radius: var(--radius-xl);
            border: 1px solid var(--line);
            background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(245,248,253,0.96));
            box-shadow: var(--shadow-sm);
            margin-bottom: 1.35rem;
            align-items: center;
            flex-wrap: wrap;
        }

        .aa-shell__brand {
            display: flex;
            align-items: center;
            gap: 1rem;
            min-width: 0;
            flex: 1 1 540px;
        }

        .aa-shell__logo {
            width: 68px;
            height: 68px;
            border-radius: 20px;
            background: linear-gradient(180deg, #ffffff, #f4f8fc);
            border: 1px solid rgba(15, 23, 42, 0.08);
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.7);
        }

        .aa-shell__logo img {
            width: 100%;
            height: 100%;
            object-fit: contain;
        }

        .aa-shell__eyebrow {
            font-size: 0.78rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            color: var(--accent);
            margin-bottom: 0.15rem;
        }

        .aa-shell__title {
            font-family: "Manrope", "Segoe UI", sans-serif;
            font-size: clamp(1.45rem, 2.5vw, 2rem);
            font-weight: 800;
            letter-spacing: -0.03em;
            color: var(--text);
            margin: 0;
        }

        .aa-shell__copy {
            margin-top: 0.3rem;
            color: var(--text-soft);
            line-height: 1.55;
            max-width: 740px;
        }

        .aa-context-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(168px, 1fr));
            gap: 0.9rem;
            margin: 0 0 1.35rem 0;
        }

        .aa-context-card {
            border-radius: var(--radius-md);
            border: 1px solid var(--line);
            background: rgba(255, 255, 255, 0.88);
            padding: 0.95rem 1rem;
            box-shadow: var(--shadow-sm);
        }

        .aa-context-card__label {
            font-size: 0.76rem;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-weight: 700;
            color: var(--text-faint);
            margin-bottom: 0.35rem;
        }

        .aa-context-card__value {
            font-size: 0.98rem;
            font-weight: 700;
            color: var(--text);
            line-height: 1.45;
            word-break: break-word;
        }

        .aa-tile__icon {
            width: 58px;
            height: 58px;
            border-radius: 18px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 1rem;
            background: linear-gradient(180deg, #eef5ff, #f7fbff);
            border: 1px solid rgba(34, 95, 155, 0.10);
            color: var(--accent-strong);
        }

        .aa-tile__title {
            font-family: "Manrope", "Segoe UI", sans-serif;
            font-size: 1.22rem;
            font-weight: 800;
            color: var(--text);
            margin-bottom: 0.5rem;
            letter-spacing: -0.02em;
        }

        .aa-tile__copy {
            color: var(--text-soft);
            line-height: 1.65;
            min-height: 4.9rem;
        }

        .aa-panel-intro {
            border-radius: var(--radius-lg);
            border: 1px solid var(--line);
            background: linear-gradient(180deg, #ffffff, #f8fbff);
            padding: 1.2rem 1.25rem;
            box-shadow: var(--shadow-sm);
            margin-bottom: 1rem;
        }

        .aa-panel-intro__kicker {
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-weight: 700;
            color: var(--accent);
            margin-bottom: 0.35rem;
        }

        .aa-panel-intro__title {
            font-family: "Manrope", "Segoe UI", sans-serif;
            font-size: 1.32rem;
            font-weight: 800;
            color: var(--text);
            margin-bottom: 0.35rem;
            letter-spacing: -0.03em;
        }

        .aa-panel-intro__copy {
            color: var(--text-soft);
            line-height: 1.65;
        }

        .section-head,
        .meta-card,
        .finding-card,
        .upload-card,
        .quick-card,
        .report-meta-card,
        .filter-panel-shell,
        .login-brand-card,
        .login-form-card,
        .kpi-card,
        .alert-card,
        .sidebar-user-card,
        .app-header,
        .empty-state-shell {
            background: var(--surface) !important;
            border: 1px solid var(--line) !important;
            color: var(--text) !important;
            box-shadow: var(--shadow-sm) !important;
        }

        .section-head,
        .meta-card,
        .finding-card,
        .upload-card,
        .quick-card,
        .report-meta-card,
        .kpi-card,
        .alert-card,
        .filter-panel-shell,
        .login-brand-card,
        .login-form-card,
        .empty-state-shell {
            border-radius: var(--radius-lg) !important;
        }

        .section-kicker,
        .meta-label,
        .upload-step,
        .report-meta-label,
        .app-header__eyebrow,
        .filter-panel-kicker,
        .login-kicker {
            color: var(--accent) !important;
        }

        .section-title,
        .meta-value,
        .finding-title,
        .upload-title,
        .quick-title,
        .report-meta-value,
        .app-header__title,
        .filter-panel-title,
        .login-title,
        .login-form-heading,
        .kpi-value {
            color: var(--text) !important;
        }

        .section-copy,
        .finding-copy,
        .upload-copy,
        .quick-copy,
        .app-header__subtitle,
        .filter-panel-copy,
        .login-copy,
        .login-form-copy,
        .kpi-copy,
        .app-header-caption,
        .upload-status-caption,
        .upload-status-meta {
            color: var(--text-soft) !important;
        }

        .report-metadata-grid,
        .upload-status-grid {
            gap: 0.9rem;
        }

        .upload-status-card {
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid var(--line);
            border-radius: var(--radius-md);
            box-shadow: var(--shadow-sm);
            padding: 1rem;
        }

        .upload-status-label {
            color: var(--text-faint);
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-size: 0.74rem;
            font-weight: 700;
            margin-bottom: 0.4rem;
        }

        .upload-status-name {
            color: var(--text);
            font-size: 1rem;
            font-weight: 700;
            margin-bottom: 0.3rem;
            word-break: break-word;
        }

        .context-chip,
        .pill {
            background: #eef4fb !important;
            border: 1px solid rgba(34, 95, 155, 0.10) !important;
            color: var(--text) !important;
        }

        .pill-positive {
            color: var(--success) !important;
            background: #edf9f4 !important;
        }

        .pill-negative {
            color: var(--danger) !important;
            background: #fff1f0 !important;
        }

        .kpi-card--positive {
            background: linear-gradient(180deg, #f4fbf7, #ecf9f2) !important;
        }

        .kpi-card--negative {
            background: linear-gradient(180deg, #fff7f6, #fff1ef) !important;
        }

        .kpi-card--neutral {
            background: linear-gradient(180deg, #ffffff, #f7faff) !important;
        }

        .brand-wordmark,
        .brand-badge {
            color: var(--text) !important;
            background: var(--surface-muted) !important;
            border: 1px solid var(--line) !important;
        }

        [data-testid="stDataFrame"],
        [data-testid="stMetric"],
        .stDataFrame,
        .stMetric {
            border-radius: var(--radius-md);
        }

        @media (max-width: 960px) {
            .block-container {
                padding-top: 1rem;
                padding-left: 1rem;
                padding-right: 1rem;
            }

            .aa-hero {
                padding-top: 1.2rem;
            }

            .aa-shell {
                padding: 1rem;
            }

            .aa-tile__copy {
                min-height: auto;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _icon_svg(icon_name: str) -> str:
    icons = {
        "dashboard": """
            <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <rect x="3" y="4" width="8" height="7" rx="2"></rect>
                <rect x="13" y="4" width="8" height="11" rx="2"></rect>
                <rect x="3" y="13" width="8" height="7" rx="2"></rect>
                <path d="M13 19h8"></path>
            </svg>
        """,
        "filters": """
            <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M4 6h16"></path>
                <path d="M7 12h10"></path>
                <path d="M10 18h4"></path>
            </svg>
        """,
        "files": """
            <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"></path>
                <path d="M14 3v5h5"></path>
                <path d="M9 13h6"></path>
                <path d="M9 17h6"></path>
            </svg>
        """,
        "charts": """
            <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M4 19h16"></path>
                <path d="M7 16l3-4 3 2 4-6"></path>
                <circle cx="7" cy="16" r="1.2"></circle>
                <circle cx="10" cy="12" r="1.2"></circle>
                <circle cx="13" cy="14" r="1.2"></circle>
                <circle cx="17" cy="8" r="1.2"></circle>
            </svg>
        """,
    }
    return icons.get(icon_name, icons["dashboard"])


def build_logo_markup(logo_data_uri: str | None, *, compact: bool = False) -> str:
    if logo_data_uri:
        return (
            f'<img src="{logo_data_uri}" alt="{APP_TITLE} logo" '
            f'class="{"aa-shell__logo-img" if compact else "aa-hero__logo-img"}" />'
        )
    fallback_label = "AA" if compact else "AA"
    return f'<div class="aa-hero__fallback">{fallback_label}</div>'


def render_home_hero(logo_markup: str):
    st.markdown(
        f"""
        <div class="aa-hero">
            <div class="aa-hero__logo">{logo_markup}</div>
            <div class="aa-hero__title">{APP_TITLE}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_workspace_shell(logo_markup: str, eyebrow: str, title: str, copy: str):
    st.markdown(
        f"""
        <div class="aa-shell">
            <div class="aa-shell__brand">
                <div class="aa-shell__logo">{logo_markup}</div>
                <div>
                    <div class="aa-shell__eyebrow">{html.escape(str(eyebrow))}</div>
                    <div class="aa-shell__title">{html.escape(str(title))}</div>
                    <div class="aa-shell__copy">{html.escape(str(copy))}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_context_cards(items: list[dict[str, str]]):
    cards_html = "".join(
        (
            '<div class="aa-context-card">'
            f'<div class="aa-context-card__label">{html.escape(str(item.get("label", "")))}</div>'
            f'<div class="aa-context-card__value">{html.escape(str(item.get("value", "")))}</div>'
            "</div>"
        )
        for item in items
        if item.get("value")
    )
    if cards_html:
        st.markdown(
            f'<div class="aa-context-grid">{cards_html}</div>',
            unsafe_allow_html=True,
        )


def render_panel_intro(kicker: str, title: str, copy: str):
    st.markdown(
        f"""
        <div class="aa-panel-intro">
            <div class="aa-panel-intro__kicker">{html.escape(str(kicker))}</div>
            <div class="aa-panel-intro__title">{html.escape(str(title))}</div>
            <div class="aa-panel-intro__copy">{html.escape(str(copy))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_tile_card(tile: PrimaryTile):
    st.markdown(
        f"""
        <div class="aa-tile">
            <div class="aa-tile__icon">{_icon_svg(tile.icon)}</div>
            <div class="aa-tile__title">{html.escape(tile.title)}</div>
            <div class="aa-tile__copy">{html.escape(tile.copy)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
