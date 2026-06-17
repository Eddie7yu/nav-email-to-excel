@echo off
rem ============================================================
rem  One-click: create 5 scheduled tasks. RUN THIS AS ADMINISTRATOR.
rem  Schedule: Mon 11:00 / 16:00 quiet write; Mon 17:00 write + email;
rem            Wed / Fri 15:00 quiet write.
rem  Tasks locate the .bat files via this folder (%~dp0): portable.
rem  /IT = run only when logged on (interactive token -> can reach the
rem        network share); /RL HIGHEST = highest privileges.
rem ============================================================
set "DIR=%~dp0"
echo Creating tasks. Folder: %DIR%
echo.
schtasks /Create /TN "NAV-Mon-1100"       /TR "\"%DIR%run_quiet.bat\""  /SC WEEKLY /D MON /ST 11:00 /IT /RL HIGHEST /F
schtasks /Create /TN "NAV-Mon-1600"       /TR "\"%DIR%run_quiet.bat\""  /SC WEEKLY /D MON /ST 16:00 /IT /RL HIGHEST /F
schtasks /Create /TN "NAV-Mon-1700-email" /TR "\"%DIR%run_weekly.bat\"" /SC WEEKLY /D MON /ST 17:00 /IT /RL HIGHEST /F
schtasks /Create /TN "NAV-Wed-1500"       /TR "\"%DIR%run_quiet.bat\""  /SC WEEKLY /D WED /ST 15:00 /IT /RL HIGHEST /F
schtasks /Create /TN "NAV-Fri-1500"       /TR "\"%DIR%run_quiet.bat\""  /SC WEEKLY /D FRI /ST 15:00 /IT /RL HIGHEST /F
echo.
echo Done. To remove all tasks later, run delete_tasks.bat.
pause
