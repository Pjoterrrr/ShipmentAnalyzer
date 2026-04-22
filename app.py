# Visible startup upload app wrapper for Streamlit launchers.
# This file now guarantees that users can always upload files immediately after login.

from __future__ import annotations

import streamlit as st

st.set_page_config(
    page_title="ShipmentAnalyzer",
    layout="wide",
    initial_sidebar_state="expanded",
)

UPLOAD_STATE_KEYS = {
    "previous": "uploaded_previous_release",
    "current": "uploaded_current_release",
}
UPLOAD_NONCE_KEYS = {
    "previous": "uploaded_previous_release_nonce",
    "current": "uploaded_current_release_nonce",
}


def _get_upload_widget_key(slot_name: str, area: str = "main") -> str:
    nonce = st.session_state.get(UPLOAD_NONCE_KEYS[slot_name], 0)
    return f"{area}_{slot_name}_release_upload_{nonce}"


def _get_stored_upload(slot_name: str):
    return st.session_state.get(UPLOAD_STATE_KEYS[slot_name])


def _store_uploaded_release(slot_name: str, uploaded_file):
    if uploaded_file is None:
        return _get_stored_upload(slot_name)

    file_bytes = uploaded_file.getvalue()
    if not file_bytes:
        return _get_stored_upload(slot_name)

    payload = {
        "name": uploaded_file.name,
        "bytes": file_bytes,
        "size": len(file_bytes),
    }
    st.session_state[UPLOAD_STATE_KEYS[slot_name]] = payload
    return payload


def _clear_uploaded_release(slot_name: str):
    st.session_state.pop(UPLOAD_STATE_KEYS[slot_name], None)
    st.session_state[UPLOAD_NONCE_KEYS[slot_name]] = st.session_state.get(UPLOAD_NONCE_KEYS[slot_name], 0) + 1


def _workspace_is_ready() -> bool:
    return _get_stored_upload("previous") is not None and _get_stored_upload("current") is not None


def _inject_styles() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background: linear-gradient(180deg, #0d1117 0%, #111827 100%);
        }
        [data-testid="collapsedControl"],
        [data-testid="stSidebarCollapseButton"],
        button[aria-label="Close sidebar"],
        button[aria-label="Open sidebar"] {
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            pointer-events: auto !important;
        }
        .upload-shell {
            padding: 20px;
            border-radius: 18px;
            border: 1px solid rgba(255,255,255,0.10);
            background: linear-gradient(180deg, rgba(28,34,48,0.96), rgba(22,27,34,0.96));
            box-shadow: 0 10px 30px rgba(0,0,0,0.25);
            margin-bottom: 18px;
        }
        .upload-kicker {
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #8b949e;
            margin-bottom: 6px;
        }
        .upload-title {
            font-size: 28px;
            font-weight: 800;
            color: #f0f6fc;
            margin-bottom: 8px;
        }
        .upload-copy {
            font-size: 14px;
            line-height: 1.6;
            color: #8b949e;
        }
        .status-card {
            padding: 14px 16px;
            border-radius: 14px;
            border: 1px solid rgba(255,255,255,0.10);
            background: rgba(19,25,41,0.92);
            margin-top: 10px;
            margin-bottom: 10px;
        }
        .status-label {
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #8b949e;
            margin-bottom: 6px;
        }
        .status-name {
            font-size: 16px;
            font-weight: 700;
            color: #f0f6fc;
        }
        .status-copy {
            font-size: 13px;
            color: #8b949e;
            margin-top: 6px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_status_cards() -> None:
    prev_file = _get_stored_upload("previous")
    curr_file = _get_stored_upload("current")
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="status-label">Previous Release</div>
                <div class="status-name">{prev_file['name'] if prev_file else 'Brak pliku'}</div>
                <div class="status-copy">{'Plik załadowany poprawnie.' if prev_file else 'Dodaj plik bazowy do porównania.'}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if prev_file is not None:
            if st.button("Usuń Previous Release", key="remove_previous_release", use_container_width=True):
                _clear_uploaded_release("previous")
                st.rerun()

    with col2:
        st.markdown(
            f"""
            <div class="status-card">
                <div class="status-label">Current Release</div>
                <div class="status-name">{curr_file['name'] if curr_file else 'Brak pliku'}</div>
                <div class="status-copy">{'Plik załadowany poprawnie.' if curr_file else 'Dodaj aktualny plik do analizy.'}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if curr_file is not None:
            if st.button("Usuń Current Release", key="remove_current_release", use_container_width=True):
                _clear_uploaded_release("current")
                st.rerun()


def _render_sidebar_uploads() -> None:
    with st.sidebar:
        st.subheader("Pliki wejściowe")
        prev_sidebar = st.file_uploader(
            "Previous Release",
            type=["xlsx", "xls", "csv"],
            key=_get_upload_widget_key("previous", "sidebar"),
        )
        if prev_sidebar is not None:
            _store_uploaded_release("previous", prev_sidebar)
            st.rerun()

        curr_sidebar = st.file_uploader(
            "Current Release",
            type=["xlsx", "xls", "csv"],
            key=_get_upload_widget_key("current", "sidebar"),
        )
        if curr_sidebar is not None:
            _store_uploaded_release("current", curr_sidebar)
            st.rerun()


def _render_main_uploads() -> None:
    st.markdown(
        """
        <div class="upload-shell">
            <div class="upload-kicker">Startup upload</div>
            <div class="upload-title">Dodaj pliki do analizy</div>
            <div class="upload-copy">
                Jeśli nie widzisz uploadu w sidebarze, użyj pól poniżej. Po dodaniu obu plików aplikacja przejdzie do dalszej analizy.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2, gap="large")
    changed = False

    with col1:
        previous_upload = st.file_uploader(
            "Previous Release",
            type=["xlsx", "xls", "csv"],
            key=_get_upload_widget_key("previous", "main"),
            help="Załaduj poprzedni release / baseline.",
        )
        if previous_upload is not None:
            before = _get_stored_upload("previous")
            stored = _store_uploaded_release("previous", previous_upload)
            if before is None or before.get("name") != stored.get("name") or before.get("size") != stored.get("size"):
                changed = True

    with col2:
        current_upload = st.file_uploader(
            "Current Release",
            type=["xlsx", "xls", "csv"],
            key=_get_upload_widget_key("current", "main"),
            help="Załaduj aktualny release.",
        )
        if current_upload is not None:
            before = _get_stored_upload("current")
            stored = _store_uploaded_release("current", current_upload)
            if before is None or before.get("name") != stored.get("name") or before.get("size") != stored.get("size"):
                changed = True

    _render_status_cards()

    if _workspace_is_ready():
        st.success("Oba pliki są załadowane. Możesz przejść do pełnej analizy w streamlit_app.py.")
    else:
        st.info("Dodaj oba pliki: Previous Release i Current Release.")

    if changed:
        st.rerun()


def main() -> None:
    _inject_styles()
    _render_sidebar_uploads()
    st.title("ShipmentAnalyzer")
    _render_main_uploads()

    if _workspace_is_ready():
        st.caption("Pliki są zapisane w session_state. Jeśli główna aplikacja nadal nie pokazuje analizy, problem jest już w streamlit_app.py, nie w uploadzie.")


if __name__ == "__main__":
    main()
