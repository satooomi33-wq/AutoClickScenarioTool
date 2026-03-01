@echo off
REM フルパス不要。ps1 と同じフォルダに置いた場合:
set SCRIPT_DIR=%~dp0
REM 引数1 = ポート (例 COM3)、 引数2 = コマンド (例 KEY:A:50)
if "%1"=="" (
  set PORT=COM3
) else (
  set PORT=%1
)
if "%2"=="" (
  set CMD=KEY:A:50
) else (
  set CMD=%2
)
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%send_key_to_notepad.ps1" -Port "%PORT%" -Cmd "%CMD%"
pause
