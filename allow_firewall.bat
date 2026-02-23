@echo off
cd /d "%~dp0"
:: Run this ONCE as Administrator so other PCs can open the app
title Allow BMS Point Tool - Firewall
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo Please RIGHT-CLICK this file and choose "Run as administrator"
    echo Then run it again.
    echo.
    pause
    exit /b 1
)

if exist "%~dp0PORT.txt" (set /p PORT=<"%~dp0PORT.txt") else (set PORT=6001)
set PORT=%PORT: =%
echo Adding Windows Firewall rules for BMS Point Tool (port %PORT%)...
netsh advfirewall firewall delete rule name="BMS Point Tool" >nul 2>&1
netsh advfirewall firewall delete rule name="BMS Point Tool - Python" >nul 2>&1

:: Allow port for all network types (Private, Public, Domain)
netsh advfirewall firewall add rule name="BMS Point Tool" dir=in action=allow protocol=TCP localport=%PORT% profile=any
if errorlevel 1 (
    echo Port rule failed. Trying with profile=private,public...
    netsh advfirewall firewall add rule name="BMS Point Tool" dir=in action=allow protocol=TCP localport=%PORT% profile=private
    netsh advfirewall firewall add rule name="BMS Point Tool" dir=in action=allow protocol=TCP localport=%PORT% profile=public
)

:: Allow Python so the app can accept connections (use this folder's .venv if present)
set PYEXE=
if exist "%~dp0.venv\Scripts\python.exe" set PYEXE=%~dp0.venv\Scripts\python.exe
if exist "%~dp0venv\Scripts\python.exe" set PYEXE=%~dp0venv\Scripts\python.exe
if not defined PYEXE for /f "delims=" %%i in ('where python 2^>nul') do set PYEXE=%%i
if defined PYEXE (
    netsh advfirewall firewall add rule name="BMS Point Tool - Python" dir=in action=allow program="%PYEXE%" profile=any
    if errorlevel 1 (echo Python rule failed.) else (echo Python allowed: %PYEXE%)
) else (
    echo Python not found. Port rule only was added.
)

echo.
echo ============================================================
echo  Done. Port %PORT% is now allowed.
echo  On another PC open in browser:  http://THIS_PC_IP:%PORT%
echo  Get THIS_PC_IP by running "ipconfig" on this PC (IPv4).
echo  Both PCs must be on the same WiFi/LAN.
echo ============================================================
pause
