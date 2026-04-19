#!/usr/bin/env python

try:
    print("Attempting to import streamlit...")
    import streamlit

    print(f"[ok] streamlit version: {streamlit.__version__}")

    print("Attempting to import pandas...")
    import pandas

    print(f"[ok] pandas version: {pandas.__version__}")

    print("Attempting to import openpyxl...")
    import openpyxl

    print(f"[ok] openpyxl version: {openpyxl.__version__}")

    print("Attempting to import altair...")
    import altair

    print(f"[ok] altair version: {altair.__version__}")

    print("Attempting to compile app entrypoints...")
    import py_compile

    py_compile.compile("analytics_calendar.py", doraise=True)
    py_compile.compile("streamlit_app.py", doraise=True)
    py_compile.compile("app.py", doraise=True)
    print("[ok] source files compiled successfully")

except Exception as exc:
    import traceback

    print(f"\n[error] {type(exc).__name__}: {exc}")
    print("\nFull traceback:")
    traceback.print_exc()
