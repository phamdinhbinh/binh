@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment not found. Run setup_and_run.bat first.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" mf4_reader_v6.1.py
