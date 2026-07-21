@echo off
setlocal EnableExtensions
chcp 65001 >nul
pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
  echo Cannot enter the NAV automation runtime directory.
  pause
  exit /b 2
)
title NAV email automation - authorization code
if not exist "app\.venv\Scripts\python.exe" (
  echo Runtime Python was not found. Ask the AI to check this deployment.
  pause
  popd
  exit /b 2
)
echo This window only accepts the email authorization code.
echo The code will not be shown in chat or logs.
echo.
"app\.venv\Scripts\python.exe" -X utf8 "app\navctl.py" secret set
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo.
  echo The authorization code was not saved. Keep this window open and tell the AI.
  pause
  popd
  exit /b %RC%
)
echo.
echo The authorization code is encrypted and saved. The AI will verify it.
echo Press any key to close this window.
pause >nul
popd
exit /b 0
