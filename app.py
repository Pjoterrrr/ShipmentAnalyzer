# Compatibility wrapper for Streamlit launchers.
# The real application logic lives in streamlit_app.py.
# This file ensures both 'streamlit run app.py' and direct imports work correctly.

if __name__ == "__main__":
    import streamlit_app
