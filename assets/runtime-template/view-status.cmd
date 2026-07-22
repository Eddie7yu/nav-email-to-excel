@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul 2>&1
pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
  echo Cannot enter the NAV automation runtime directory.
  echo Keep this window open and ask the AI to inspect the deployment path.
  pause
  exit /b 2
)
title NAV email automation - status
if not exist "app\.venv\Scripts\python.exe" (
  echo Runtime Python was not found. Ask the AI to check this deployment.
  pause
  popd
  exit /b 2
)
if not exist "app\navctl.py" (
  echo Runtime controller was not found. Ask the AI to repair this deployment.
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
