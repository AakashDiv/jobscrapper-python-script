@echo off
setlocal

cd /d "%~dp0"

echo ============================================
echo   HR Job Scraper v4 - Daily Mail Runner
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found
    exit /b 1
)

python daily_run.py
exit /b %errorlevel%
