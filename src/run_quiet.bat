@echo off
rem Quiet run: write master (with backup), NO email. For mid-week scheduled runs.
rem Uses pushd (not cd /d) so it works from a UNC network path.
rem Self-diagnostics are written to logs\bat_quiet.txt (reset each run with > then >>).
pushd "%~dp0"
set PYTHONIOENCODING=utf-8
if not exist "logs" mkdir "logs"
echo === start %date% %time% === > "logs\bat_quiet.txt"
echo cwd=%CD% >> "logs\bat_quiet.txt"
echo --- where python --- >> "logs\bat_quiet.txt"
where python >> "logs\bat_quiet.txt" 2>&1
echo --- python --version --- >> "logs\bat_quiet.txt"
python --version >> "logs\bat_quiet.txt" 2>&1
echo --- run_weekly.py --- >> "logs\bat_quiet.txt"
python -u run_weekly.py --commit --no-notify >> "logs\bat_quiet.txt" 2>&1
echo === exit code %errorlevel% === >> "logs\bat_quiet.txt"
popd
