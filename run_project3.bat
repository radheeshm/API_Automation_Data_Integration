@echo off
cd /d "%~dp0"
if exist "venv\Scripts\python.exe" (
    "venv\Scripts\python.exe" main.py
) else (
    echo Virtual environment not found.
    echo Run: py -m venv venv
    echo Then: venv\Scripts\python.exe -m pip install -r requirements.txt
    pause
)
