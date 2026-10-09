@echo off
rem Daily code.org progress capture, started by the "CodeOrg Daily Progress" scheduled task.
cd /d "%~dp0"
if not exist output\logs mkdir output\logs
set PYTHONIOENCODING=utf-8
.venv\Scripts\python.exe -u codeorg_colleges.py >> output\logs\run.log 2>&1
