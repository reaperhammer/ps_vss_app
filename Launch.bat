@echo off

REM cd to the directory of this batch file so the GUI's relative asset
REM paths (e.g. .\assets\png\app_icon.png) resolve correctly when the
REM script is launched from any process working directory.
cd /d "%~dp0"

set "SCRIPT_PATH=%~dp0VSSManager.ps1"
set "WORK_DIR=%~dp0"

REM Launch PowerShell as administrator with the script path
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -Verb RunAs powershell -WorkingDirectory $env:WORK_DIR -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $env:SCRIPT_PATH"