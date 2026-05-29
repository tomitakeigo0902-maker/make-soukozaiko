@echo off
chcp 65001 >nul
title 倉庫在庫管理 - サーバー自動起動の設定

rem === このバッチは soukozaiko.exe と同じフォルダに置いて実行してください ===
rem  Windows の起動と同時に soukozaiko.exe が自動で立ち上がるよう登録します。
rem  一度だけ実行すれば、以降は PC を起動するたびに自動でサーバーが動きます。

rem --- 管理者権限が無ければ昇格して実行し直す ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 管理者権限が必要です。確認画面が出たら「はい」を押してください...
    powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

set "EXE=%~dp0soukozaiko.exe"
set "TASK=SoukozaikoServer"

if not exist "%EXE%" (
    echo.
    echo [エラー] 同じフォルダに soukozaiko.exe が見つかりません。
    echo          このバッチを soukozaiko.exe と同じフォルダに置いてから
    echo          もう一度実行してください。
    echo.
    pause
    exit /b 1
)

echo.
echo 自動起動を登録します...
echo   実行ファイル: %EXE%

rem ONSTART = Windows 起動時。SYSTEM 権限で動かすのでログインしていなくても起動。
rem /RL HIGHEST = 管理者権限で実行（ファイアウォール設定もそのまま通る）。
schtasks /Create /F /TN "%TASK%" /SC ONSTART /RU SYSTEM /RL HIGHEST ^
    /TR "\"%EXE%\""

if errorlevel 1 (
    echo.
    echo [エラー] 自動起動の登録に失敗しました。
    pause
    exit /b 1
)

rem 自動起動には画面が無いので、ブラウザは自動で開かないようにする
setx SOUKOZAIKO_NO_BROWSER 1 /M >nul 2>&1

echo.
echo 今すぐサーバーを起動します...
schtasks /Run /TN "%TASK%" >nul 2>&1

echo.
echo ============================================================
echo  設定が完了しました。
echo.
echo  ・これ以降、この PC を起動すると自動でサーバーが立ち上がります。
echo  ・もう exe を手動で起動する必要はありません。
echo  ・他の PC はブラウザで  http://（このPCのIP）:8000/  を開くだけ。
echo.
echo  自動起動をやめたいときは「サーバー自動起動を解除.bat」を実行してください。
echo ============================================================
echo.
pause
