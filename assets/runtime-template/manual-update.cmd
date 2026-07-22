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
title NAV email automation - manual update
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
