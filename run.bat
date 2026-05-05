@echo off
echo ============================================
echo   HR Job Scraper v4 - Delhi NCR Last 24h
echo ============================================
echo.
python --version >nul 2>&1
if errorlevel 1 (echo ERROR: Python not found & pause & exit /b 1)
echo Installing / verifying dependencies...
pip install selenium webdriver-manager requests beautifulsoup4 pandas openpyxl tqdm fake-useragent undetected-chromedriver --quiet
echo.
echo Starting scraper...
echo.
python main.py
echo.
echo Done! Check HR_Jobs_Last24h.xlsx in this folder.
pause
