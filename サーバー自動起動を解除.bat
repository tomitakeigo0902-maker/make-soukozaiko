@echo off
chcp 65001 >nul
title 倉庫在庫管理 - サーバー自動起動の解除

rem === Windows 起動時の自動起動をやめ、動いているサーバーを止めます ===

rem --- 管理者権限が無ければ昇格して実行し直す ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 管理者権限が必要です。確認画面が出たら「はい」を押してください...
    powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

set "TASK=SoukozaikoServer"

echo.
echo 自動起動を解除します...

rem 動いているサーバーを停止（タスク経由・プロセス直接の両方）
schtasks /End /TN "%TASK%" >nul 2>&1
taskkill /IM soukozaiko.exe /F >nul 2>&1

rem 登録を削除
schtasks /Delete /TN "%TASK%" /F >nul 2>&1

echo.
echo ============================================================
echo  自動起動を解除しました。
echo.
echo  ・PC を起動してもサーバーは自動では立ち上がりません。
echo  ・また使いたいときは soukozaiko.exe をダブルクリックするか、
echo    「サーバー自動起動を設定.bat」で再登録してください。
echo ============================================================
echo.
pause
