@echo off
title Grocery Scraper - Running...
cd /d "%~dp0"
echo.
echo Starting Fresh St. Market scraper - please wait 30-40 minutes.
echo.
python grocery_scraper.py
echo.
echo Starting Save-On-Foods scraper - please wait 1-3 hours.
echo A browser window may briefly appear if a security check needs solving.
echo.
python save_on_foods_scraper.py
if %errorlevel% neq 0 pause
