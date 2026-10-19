@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" demo_kiem_tra_boc_tach.py
) else (
    python demo_kiem_tra_boc_tach.py
)
if errorlevel 1 pause
