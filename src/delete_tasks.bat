@echo off
rem Delete the 5 scheduled tasks. RUN THIS AS ADMINISTRATOR.
rem Use this when changing the schedule: delete, edit setup_tasks.bat times, re-run it.
schtasks /Delete /TN "NAV-Mon-1100"       /F
schtasks /Delete /TN "NAV-Mon-1600"       /F
schtasks /Delete /TN "NAV-Mon-1700-email" /F
schtasks /Delete /TN "NAV-Wed-1500"       /F
schtasks /Delete /TN "NAV-Fri-1500"       /F
echo Done.
pause
