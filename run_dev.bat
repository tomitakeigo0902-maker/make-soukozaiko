@echo off
chcp 65001 >nul
title Soukozaiko - inventory app (dev mode)

python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERROR] Python was not found on this PC.
    echo.
    echo Please install Python 3.11 or newer from:
    echo     https://www.python.org/downloads/
    echo During installation, CHECK the box "Add python.exe to PATH".
    echo Then run this file again.
    echo.
    pause
    exit /b 1
)

echo Installing required libraries (the first run may take a while)...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install the required libraries.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Starting the app. Open this URL in your web browser:
echo      http://localhost:8000/
echo  Press Ctrl+C in this window to stop the app.
echo ============================================================
echo.
python main.py
pause
