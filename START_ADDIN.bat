@echo off
title Data Matrix Generator - Add-in Server
color 0A

echo.
echo  ========================================
echo   Data Matrix Generator - Excel Add-in
echo  ========================================
echo.

:: Check Node is available
node --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Node.js is not installed or not in PATH.
    echo  Please install from https://nodejs.org/
    pause
    exit /b 1
)

:: Move to the script's own directory (works from any shortcut)
cd /d "%~dp0"

:: Install dependencies if node_modules is missing
if not exist "node_modules\" (
    echo  First run: Installing dependencies...
    echo  (This only happens once)
    echo.
    call npm install
    if errorlevel 1 (
        echo.
        echo  ERROR: npm install failed. Check your internet connection.
        pause
        exit /b 1
    )
)

:: Install SSL certificate if not already done
if not exist "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" (
    echo  Installing local SSL certificate (one-time setup)...
    call npx office-addin-dev-certs install
    if errorlevel 1 (
        echo.
        echo  ERROR: Certificate installation failed.
        echo  Try running this script as Administrator.
        pause
        exit /b 1
    )
    echo  Certificate installed successfully!
    echo.
)

echo  Starting add-in server at https://localhost:3000
echo  Keep this window open while using the add-in.
echo  Press Ctrl+C to stop.
echo.

node server.js

pause
