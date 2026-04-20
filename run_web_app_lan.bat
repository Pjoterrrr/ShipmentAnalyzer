@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
    echo Python nie jest zainstalowany albo nie ma go w PATH.
    echo Zainstaluj Python 3.11+ i uruchom ten plik ponownie.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Tworzenie lokalnego srodowiska .venv...
    py -3 -m venv .venv
    if errorlevel 1 (
        echo Nie udalo sie utworzyc srodowiska virtualnego.
        pause
        exit /b 1
    )
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Instalacja zaleznosci nie powiodla sie.
    pause
    exit /b 1
)

python -m streamlit run streamlit_app.py --server.address 0.0.0.0 --server.port 8501
pause
