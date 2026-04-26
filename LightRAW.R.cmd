@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT_DIR%LightRAW.R.ps1" %*
if errorlevel 1 (
  echo.
  echo Viewer exited with an error. Press any key to close this window.
  pause >nul
)
