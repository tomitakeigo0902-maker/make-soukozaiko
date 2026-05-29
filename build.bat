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
echo  HOW TO USE (server PC):
echo    1. Copy dist\soukozaiko.exe to ONE office PC (the "server").
echo       No Python needed on that PC.
echo    2. Double-click soukozaiko.exe to start. A black window stays
echo       open and the browser opens automatically.
echo       (On first run, click "Yes" if Windows asks to allow the
echo        firewall rule, so other PCs can connect.)
echo    3. Data is saved in inventory.db next to the exe. Back it up by
echo       copying that file.
echo.
echo  HOW TO USE (other PCs):
echo    - Do NOT open the exe. Just open a web browser and go to the
echo      URL shown in the server's black window, e.g.
echo          http://192.168.x.x:8000/
echo    - Or copy the file "倉庫在庫管理を開く.url" (created next to the
echo      exe at startup) onto each PC's desktop and double-click it.
echo ============================================================
pause
