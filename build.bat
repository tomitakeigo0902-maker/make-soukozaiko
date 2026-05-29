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
echo [2/3] Building the exe...
pyinstaller --noconfirm --onefile --name soukozaiko --add-data "static;static" main.py
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo [3/3] Copying helper files next to the exe...
copy /Y "サーバー自動起動を設定.bat" "dist\" >nul 2>&1
copy /Y "サーバー自動起動を解除.bat" "dist\" >nul 2>&1

echo.
echo ============================================================
echo  Done!  Output folder:  dist\
echo    - soukozaiko.exe              (the app)
echo    - サーバー自動起動を設定.bat   (set up auto-start, run once)
echo    - サーバー自動起動を解除.bat   (undo auto-start)
echo.
echo  HOW TO USE (server PC):
echo    1. Copy the whole dist\ folder to ONE office PC (the "server").
echo       No Python needed on that PC. Avoid Program Files; e.g. use
echo       C:\soukozaiko\ so the data file can be written easily.
echo    2. To start it ONCE: double-click soukozaiko.exe.
echo       (On first run, click "Yes" if Windows asks to allow the
echo        firewall rule, so other PCs can connect.)
echo    3. RECOMMENDED for an always-on server: double-click
echo       "サーバー自動起動を設定.bat" once. After that the server
echo       starts automatically every time Windows boots, with no
echo       need to launch the exe by hand.
echo    4. Data is saved in inventory.db next to the exe. Back it up by
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
