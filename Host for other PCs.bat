@echo off
title BMS Point Tool - Host (open from any PC)
cd /d "%~dp0"

echo.
echo  BMS Point Tool - HOST for other PCs
echo  You will get a public link that works from any PC.
echo.
echo  One-time: get free token at https://ngrok.com
echo  Then run:  set NGROK_AUTHTOKEN=your_token
echo.
echo  When someone opens the link, they may see an ngrok page
echo  - they must click "Visit Site" to open the app.
echo.
if exist "%~dp0.venv\Scripts\python.exe" (
    set "PY=%~dp0.venv\Scripts\python.exe"
) else (
    set "PY=python"
)
"%PY%" --version >nul 2>&1
if errorlevel 1 (
    echo Python not found.
    pause
    exit /b 1
)

"%PY%" app.py --host-public
pause
