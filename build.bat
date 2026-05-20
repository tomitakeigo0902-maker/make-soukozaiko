@echo off
chcp 65001 > nul
echo ============================================
echo   倉庫在庫管理アプリ  exe ビルド
echo ============================================
echo.

echo [1/2] 必要なライブラリをインストールします...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ライブラリのインストールに失敗しました。
    echo Python がインストールされているか確認してください。
    pause
    exit /b 1
)

echo.
echo [2/2] exe をビルドします...
pyinstaller --noconfirm --onefile --name soukozaiko --add-data "static;static" main.py
if errorlevel 1 (
    echo.
    echo ビルドに失敗しました。
    pause
    exit /b 1
)

echo.
echo ============================================
echo   完成しました:  dist\soukozaiko.exe
echo.
echo   この exe 1 ファイルを社内サーバーにコピーし、
echo   ダブルクリックで起動してください。
echo   データは exe と同じフォルダの inventory.db に
echo   保存されます。
echo ============================================
pause
