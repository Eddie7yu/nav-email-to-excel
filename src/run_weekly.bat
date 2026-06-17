@echo off
rem Full run: write master (with backup) + send summary email. Use for the report run.
rem Uses pushd (not cd /d) so it works from a UNC network path.
pushd "%~dp0"
set PYTHONIOENCODING=utf-8
if not exist "logs" mkdir "logs"
python -u run_weekly.py --commit >> "logs\bat_last.txt" 2>&1
popd
