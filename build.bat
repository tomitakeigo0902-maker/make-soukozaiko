@echo off
chcp 65001 >nul
title Soukozaiko - build exe

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

echo [1/2] Installing required libraries...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install the required libraries.
    pause
    exit /b 1
)

echo.
echo [2/2] Building the exe...
pyinstaller --noconfirm --onefile --name soukozaiko --add-data "static;static" main.py
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Done!  Output file:  dist\soukozaiko.exe
echo.
echo  Copy this single exe to your office server and
echo  double-click it to run. No Python needed on the server.
echo  Data is saved in inventory.db next to the exe.
echo ============================================================
pause
