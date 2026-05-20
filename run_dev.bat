@echo off
chcp 65001 > nul
echo 倉庫在庫管理アプリ を開発モードで起動します...
echo （初回はライブラリのインストールに少し時間がかかります）
echo.
python -m pip install -r requirements.txt
python main.py
pause
