@echo off
setlocal
cd /d "%~dp0"

set "TEMP_DIST=%LOCALAPPDATA%\Temp\ShipmentAnalyzerDist"
set "TEMP_BUILD=%LOCALAPPDATA%\Temp\ShipmentAnalyzerBuild"
set "FINAL_DIST=dist\ShipmentAnalyzerWeb_portable"

if exist "%TEMP_BUILD%" rmdir /s /q "%TEMP_BUILD%"
if exist "%TEMP_DIST%" rmdir /s /q "%TEMP_DIST%"
if exist "%FINAL_DIST%" rmdir /s /q "%FINAL_DIST%"

python -m PyInstaller ShipmentAnalyzerWeb.spec --distpath "%TEMP_DIST%" --workpath "%TEMP_BUILD%"
if errorlevel 1 (
    echo Budowanie EXE nie powiodlo sie.
    pause
    exit /b 1
)

if not exist dist mkdir dist
robocopy "%TEMP_DIST%\ShipmentAnalyzerWeb" "%FINAL_DIST%" /E >nul
if not exist "%FINAL_DIST%\config" mkdir "%FINAL_DIST%\config"
if not exist "%FINAL_DIST%\assets" mkdir "%FINAL_DIST%\assets"
copy /Y "config\users.json" "%FINAL_DIST%\config\users.json" >nul
copy /Y "assets\logo.png" "%FINAL_DIST%\assets\logo.png" >nul
copy /Y "assets\icon.ico" "%FINAL_DIST%\assets\icon.ico" >nul

echo Gotowe. Paczka znajduje sie w %FINAL_DIST%\
pause
