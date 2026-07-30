@echo off
cd /d "%~dp0"
where pyw >nul 2>nul
if errorlevel 1 (
    python app.py
) else (
    pyw -3 app.py
)
