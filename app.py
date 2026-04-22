# Compatibility wrapper for Streamlit launchers.
# The real application logic lives in streamlit_app.py.
# This file ensures both 'streamlit run app.py' and direct imports work correctly.

from __future__ import annotations

import streamlit as st

UPLOAD_STATE_KEYS = {
    "previous": "uploaded_previous_release",
    "current": "uploaded_current_release",
}
UPLOAD_NONCE_KEYS = {
    "previous": "uploaded_previous_release_nonce",
    "current": "uploaded_current_release_nonce",
}


def _get_upload_widget_key(slot_name: str) -> str:
    nonce = st.session_state.get(UPLOAD_NONCE_KEYS[slot_name], 0)
    return f"wrapper_{slot_name}_release_upload_{nonce}"


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


def _workspace_is_ready() -> bool:
    return _get_stored_upload("previous") is not None and _get_stored_upload("current") is not None


def _inject_sidebar_toggle_fix() -> None:
    st.markdown(
        """
        <style>
        [data-testid="collapsedControl"],
        [data-testid="stSidebarCollapseButton"],
        button[aria-label="Close sidebar"],
        button[aria-label="Open sidebar"] {
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            pointer-events: auto !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_startup_upload_fallback() -> None:
    st.markdown(
        """
        <div style="
            margin-top: 1rem;
            padding: 1rem 1.1rem;
            border-radius: 16px;
            border: 1px solid rgba(255,255,255,0.08);
            background: linear-gradient(180deg, rgba(28,34,48,0.96), rgba(22,27,34,0.96));
        ">
            <div style="font-size: 12px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: #8b949e; margin-bottom: 8px;">
                Startup upload
            </div>
            <div style="font-size: 22px; font-weight: 800; color: #f0f6fc; margin-bottom: 8px;">
                Dodaj pliki do analizy
            </div>
            <div style="font-size: 14px; line-height: 1.6; color: #8b949e;">
                Upload jest dostępny od razu po zalogowaniu. Możesz użyć pól poniżej albo lewego sidebaru.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    changed = False
    left_col, right_col = st.columns(2, gap="large")

    with left_col:
        st.markdown("#### Previous Release")
        previous_upload = st.file_uploader(
            "Previous Release",
            type=["xlsx", "xls", "csv"],
            key=_get_upload_widget_key("previous"),
            help="Załaduj poprzedni release / baseline.",
        )
        if previous_upload is not None:
            before = _get_stored_upload("previous")
            stored = _store_uploaded_release("previous", previous_upload)
            if before is None or before.get("name") != stored.get("name") or before.get("size") != stored.get("size"):
                changed = True
        stored_previous = _get_stored_upload("previous")
        if stored_previous is not None:
            st.success(f"Załadowano: {stored_previous['name']}")

    with right_col:
        st.markdown("#### Current Release")
        current_upload = st.file_uploader(
            "Current Release",
            type=["xlsx", "xls", "csv"],
            key=_get_upload_widget_key("current"),
            help="Załaduj aktualny release.",
        )
        if current_upload is not None:
            before = _get_stored_upload("current")
            stored = _store_uploaded_release("current", current_upload)
            if before is None or before.get("name") != stored.get("name") or before.get("size") != stored.get("size"):
                changed = True
        stored_current = _get_stored_upload("current")
        if stored_current is not None:
            st.success(f"Załadowano: {stored_current['name']}")

    if _workspace_is_ready():
        st.success("Oba pliki są gotowe. Uruchamiam analizę...")
        st.rerun()

    if changed:
        st.rerun()


_stop_exception = None

try:
    import streamlit_app  # noqa: F401
except BaseException as exc:  # pragma: no cover - Streamlit control-flow exception
    if exc.__class__.__name__ == "StopException":
        _stop_exception = exc
    else:
        raise

_inject_sidebar_toggle_fix()

if _stop_exception is not None:
    if st.session_state.get("authenticated") and not _workspace_is_ready():
        _render_startup_upload_fallback()
