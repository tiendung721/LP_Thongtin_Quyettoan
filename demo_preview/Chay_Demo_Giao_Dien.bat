@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" demo_giao_dien.py
) else (
    python demo_giao_dien.py
)
if errorlevel 1 pause
