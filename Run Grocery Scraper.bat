@echo off
title Grocery Scraper - Running...
cd /d "%~dp0"
echo.
echo Starting scraper - please wait 30-90 minutes.
echo Results will be saved as Excel files on your Desktop.
echo.
python grocery_scraper.py
if %errorlevel% neq 0 pause
