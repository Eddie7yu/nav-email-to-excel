@echo off
setlocal EnableExtensions DisableDelayedExpansion
pushd "%~dp0.." >nul 2>&1
if errorlevel 1 exit /b 2
if not exist "logs" mkdir "logs"
if not exist "app\.venv\Scripts\python.exe" (
  echo [launcher] Runtime Python was not found.>> "logs\update-launcher.log"
  popd
  exit /b 2
)
if not exist "app\navctl.py" (
  echo [launcher] Runtime controller was not found.>> "logs\update-launcher.log"
  popd
  exit /b 2
)
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set "STAMP=%%I"
if not defined STAMP (
  echo [launcher] Could not determine the local log date.>> "logs\update-launcher.log"
  popd
  exit /b 2
)
"app\.venv\Scripts\python.exe" -X utf8 "app\navctl.py" scheduled-update >> "logs\update-%STAMP%.log" 2>&1
set "RC=%ERRORLEVEL%"
popd
exit /b %RC%
