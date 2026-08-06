@echo off
cd /d "%~dp0"
if not exist venv\Scripts\python.exe (
    python -m venv venv
    call venv\Scripts\activate
    python -m pip install --upgrade pip
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate
)
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
pause
