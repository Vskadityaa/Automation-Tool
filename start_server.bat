@echo off
cd /d "%~dp0"
if exist "%~dp0PORT.txt" (set /p PORT=<"%~dp0PORT.txt") else (set PORT=6001)
set PORT=%PORT: =%
title BMS Point Tool - Server
echo Starting BMS Point Tool...
echo.
echo On THIS PC open:  http://localhost:%PORT%
echo On ANOTHER PC: paste the link below in the ADDRESS BAR (not search). Include http://
echo If it only loads forever: RIGHT-CLICK allow_firewall.bat - Run as administrator
echo.
python app.py
if errorlevel 1 (
    echo.
    echo Error. Installing dependencies...
    pip install -r requirements.txt
    python app.py
)
pause
