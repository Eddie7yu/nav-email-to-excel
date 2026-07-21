@echo off
setlocal EnableExtensions
chcp 65001 >nul
pushd "%~dp0" >nul 2>&1
if errorlevel 1 exit /b 2
title NAV email automation - status
if not exist "app\.venv\Scripts\python.exe" (
  echo Runtime Python was not found. Ask the AI to check this deployment.
  pause
  popd
  exit /b 2
)
"app\.venv\Scripts\python.exe" -X utf8 "app\navctl.py" schedule status
set "RC=%ERRORLEVEL%"
echo.
pause
popd
exit /b %RC%
