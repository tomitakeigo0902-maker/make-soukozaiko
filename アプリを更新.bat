@echo off
title Soukozaiko - update app

rem === 最新版の soukozaiko.exe を GitHub から取得して差し替えます ===
rem  このバッチは soukozaiko.exe と同じフォルダに置いて実行してください。
rem  サーバーを止めて exe を入れ替え、自動で再起動します。

rem --- 管理者権限が無ければ昇格して実行し直す ---
net session > nul 2>&1
if %errorlevel% neq 0 (
    echo 管理者権限が必要です。確認画面が出たら「はい」を押してください...
    powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

set "DIR=%~dp0"
set "EXE=%DIR%soukozaiko.exe"
set "TMP=%DIR%soukozaiko_new.exe"
set "BAK=%DIR%soukozaiko_backup.exe"
set "TASK=SoukozaikoServer"
set "URL=https://github.com/tomitakeigo0902-maker/make-soukozaiko/releases/download/latest/soukozaiko.exe"

echo.
echo 最新版を GitHub からダウンロードします...
echo   %URL%
echo.
powershell -NoProfile -Command "$ProgressPreference='SilentlyContinue'; try { Invoke-WebRequest -Uri '%URL%' -OutFile '%TMP%' -UseBasicParsing } catch { exit 1 }"
if errorlevel 1 (
    echo.
    echo [エラー] ダウンロードに失敗しました。
    echo          インターネット接続と GitHub へのアクセスを確認してください。
    echo          社内プロキシ環境では繋がらない場合があります。
    echo.
    pause
    exit /b 1
)

if not exist "%TMP%" (
    echo.
    echo [エラー] ダウンロードできましたが、ファイルが見つかりません。
    pause
    exit /b 1
)

echo サーバーを停止します...
schtasks /End /TN "%TASK%" > nul 2>&1
taskkill /IM soukozaiko.exe /F > nul 2>&1
rem ファイル解放を待つ
ping -n 3 127.0.0.1 > nul

if exist "%EXE%" (
    echo 旧バージョンを退避します（soukozaiko_backup.exe）...
    if exist "%BAK%" del /F /Q "%BAK%" > nul 2>&1
    move /Y "%EXE%" "%BAK%" > nul
)

echo 新しい exe に差し替えます...
move /Y "%TMP%" "%EXE%" > nul
if errorlevel 1 (
    echo.
    echo [エラー] ファイルの差し替えに失敗しました。サーバーがまだ動いている可能性があります。
    if exist "%BAK%" move /Y "%BAK%" "%EXE%" > nul
    pause
    exit /b 1
)

echo サーバーを再起動します...
schtasks /Run /TN "%TASK%" > nul 2>&1
if errorlevel 1 (
    echo タスクが見つからないため、exe を直接起動します。
    start "" "%EXE%"
)

echo.
echo ============================================================
echo  更新が完了しました。
echo.
echo  ・旧バージョンは soukozaiko_backup.exe として残してあります。
echo    問題があればサーバーを止めてこれを soukozaiko.exe にリネーム
echo    し直すと、ひとつ前に戻せます。
echo  ・在庫データ（inventory.db）はそのまま引き継がれます。
echo ============================================================
echo.
pause
