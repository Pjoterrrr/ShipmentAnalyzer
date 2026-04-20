# Shipment Analyzer

Streamlit application for comparing two Excel snapshots of demand and shipment data, highlighting changes, aggregating weekly quantities, and exporting a business-ready Excel report.

The application supports both input formats:

- legacy wide format with release data in a `Raw` sheet
- newer VL10E-style block format, even when a `Raw` sheet is not present

## Quick Start

### Requirements

- Python 3.11+
- `pip`

### Run locally

```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

The app will be available at [http://localhost:8501](http://localhost:8501).

## Project Structure

```text
.
|-- .streamlit/
|   `-- config.toml
|-- assets/
|   |-- icon.ico
|   `-- logo.png
|-- config/
|   `-- users.json
|-- tests/
|   |-- test_analytics_calendar.py
|   |-- test_excel_export.py
|   `-- test_release_loader.py
|-- analytics_calendar.py
|-- app.py
|-- release_loader.py
|-- requirements.txt
|-- run_web_app.bat
|-- run_web_app_lan.bat
`-- streamlit_app.py
```

## Main Entry Point

The main Streamlit entry point is:

```text
streamlit_app.py
```

Start the app with:

```powershell
streamlit run streamlit_app.py
```

`app.py` remains in the repository only as a lightweight compatibility wrapper.

## Features

- compare previous and current Excel files
- support legacy release format and VL10E-style format
- weekly aggregation by ISO week
- percentage change and deviation highlighting
- charts and dashboard views
- CSV download and Excel export
- local login based on `config/users.json`

## Configuration

### Streamlit config

The repository already includes `.streamlit/config.toml` for local and cloud-friendly Streamlit settings.

### Login config

Local authentication is stored in `config/users.json`.

Before publishing a public repository or deploying for shared use:

1. review the file content
2. replace any temporary or local-only user entries
3. avoid committing real personal credentials

The repository should contain only hashed passwords, never plaintext passwords or tokens.

## Publish to GitHub

1. Create an empty repository on GitHub.
2. From the project root, run:

```powershell
git init
git add .
git commit -m "Prepare Streamlit app for GitHub and deployment"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repository>.git
git push -u origin main
```

## Deployment on Streamlit Community Cloud

### What to configure

- Repository: your GitHub repository, for example `your-username/shipment-analyzer`
- Branch: `main`
- Main file path: `streamlit_app.py`

### Steps

1. Push the repository to GitHub.
2. Open [Streamlit Community Cloud](https://share.streamlit.io/).
3. Click `New app`.
4. Select your GitHub repository.
5. Set branch to `main`.
6. Set the main file path to `streamlit_app.py`.
7. Deploy the app.

If you update the repository later, Streamlit can redeploy directly from GitHub.

## Local Windows Launchers

For convenience, the repository also includes:

- `run_web_app.bat` for local host-only access
- `run_web_app_lan.bat` for LAN access

Both launchers install dependencies into a local virtual environment and start:

```powershell
streamlit run streamlit_app.py
```

## Notes

- No additional `packages.txt` is required for the current dependency set.
- The application uses only repository-relative paths for assets and configuration files.
- Generated artifacts such as virtual environments, caches, builds, exports, and temporary files should stay out of Git.
