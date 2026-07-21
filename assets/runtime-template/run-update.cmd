@echo off
setlocal EnableExtensions
pushd "%~dp0.." >nul 2>&1
if errorlevel 1 exit /b 2
if not exist "logs" mkdir "logs"
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set "STAMP=%%I"
if not defined STAMP exit /b 2
"app\.venv\Scripts\python.exe" -X utf8 "app\navctl.py" scheduled-update >> "logs\update-%STAMP%.log" 2>&1
set "RC=%ERRORLEVEL%"
popd
exit /b %RC%
