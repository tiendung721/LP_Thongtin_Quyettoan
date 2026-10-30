@echo off
REM Run PAD flow: create the number of settlement rows passed by PM.
setlocal enableextensions enabledelayedexpansion

set "PROJECT_ROOT=%~dp0.."
for %%I in ("%PROJECT_ROOT%") do set "PROJECT_ROOT=%%~fI"
set "CREATE_COUNT=%~1"
if not defined CREATE_COUNT set "CREATE_COUNT=1"

if exist "%PROJECT_ROOT%\.venv\Scripts\python.exe" (
  "%PROJECT_ROOT%\.venv\Scripts\python.exe" "%PROJECT_ROOT%\scripts\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow create_new --count "%CREATE_COUNT%"
  exit /b !errorlevel!
)

where py >nul 2>nul
if not errorlevel 1 (
  py -3 "%PROJECT_ROOT%\scripts\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow create_new --count "%CREATE_COUNT%"
  exit /b !errorlevel!
)

where python >nul 2>nul
if not errorlevel 1 (
  python "%PROJECT_ROOT%\scripts\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow create_new --count "%CREATE_COUNT%"
  exit /b !errorlevel!
)

echo [ERROR] Python not found. Cannot launch PAD flow.
exit /b 2
