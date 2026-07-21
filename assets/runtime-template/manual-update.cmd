@echo off
setlocal EnableExtensions
chcp 65001 >nul
pushd "%~dp0" >nul 2>&1
if errorlevel 1 exit /b 2
title NAV email automation - manual update
if not exist "app\.venv\Scripts\python.exe" (
  echo Runtime Python was not found. Ask the AI to check this deployment.
  pause
  popd
  exit /b 2
)
"app\.venv\Scripts\python.exe" -X utf8 "app\navctl.py" scheduled-update
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
  echo Manual update finished. Details are available in the logs folder.
) else (
  echo Manual update stopped safely. Ask the AI to inspect the logs folder.
)
pause
popd
exit /b %RC%
