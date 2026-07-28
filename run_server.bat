@echo off
SETLOCAL EnableDelayedExpansion
title MEI Analytics Dashboard Server

:: ============================================================================
::  MEI Analytics Dashboard - Windows Server Deployment Script
::  Routes:  /alarm  |  /rejection
::  Engine:  Waitress WSGI (production-grade, multi-threaded)
:: ============================================================================

:: Navigate to the folder where this .bat file lives
cd /d "%~dp0"

echo.
echo ======================================================================
echo   MEI Analytics Dashboard - Server Setup
echo ======================================================================
echo   Working Directory: %CD%
echo.

:: ---------------------------------------------------------------------------
:: 1. Check Python Installation
:: ---------------------------------------------------------------------------
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not found on this machine!
    echo         Please install Python 3.10+ from https://www.python.org
    echo         Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYTHON_VER=%%v
echo   [OK] %PYTHON_VER% detected

:: ---------------------------------------------------------------------------
:: 2. Check / Create Virtual Environment
:: ---------------------------------------------------------------------------
if exist "venv\Scripts\activate.bat" (
    echo   [OK] Virtual environment found - activating 'venv'
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    echo   [OK] Virtual environment found - activating '.venv'
    call .venv\Scripts\activate.bat
) else (
    echo   [..] No virtual environment found - creating 'venv'...
    python -m venv venv
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to create virtual environment!
        pause
        exit /b 1
    )
    call venv\Scripts\activate.bat
    echo   [OK] Virtual environment created and activated
)

:: ---------------------------------------------------------------------------
:: 3. Install / Update Dependencies
:: ---------------------------------------------------------------------------
echo   [..] Installing dependencies from requirements.txt...
pip install -r requirements.txt --quiet --disable-pip-version-check
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install dependencies!
    echo         Check your internet connection and requirements.txt
    pause
    exit /b 1
)
echo   [OK] All dependencies installed

:: ---------------------------------------------------------------------------
:: 4. Ensure Required Directories Exist
:: ---------------------------------------------------------------------------
if not exist "storage"          mkdir storage
if not exist "storage\logs"     mkdir storage\logs
if not exist "uploads"          mkdir uploads
echo   [OK] Project directories verified

:: ---------------------------------------------------------------------------
:: 5. Set Production Environment Variables
:: ---------------------------------------------------------------------------
if "%PORT%"=="" set PORT=8090
if "%HOST%"=="" set HOST=0.0.0.0
set FLASK_DEBUG=0

echo.
echo ======================================================================
echo   MEI Dashboard is starting...
echo.
echo   URL:    http://%HOST%:%PORT%/alarm
echo   Engine: Waitress (production)
echo   Logs:   %CD%\storage\logs\server.log
echo ======================================================================
echo.
echo   Press Ctrl+C to stop the server.
echo.

:: ---------------------------------------------------------------------------
:: 6. Launch the Application
:: ---------------------------------------------------------------------------
python app.py

:: If the server stops or crashes, pause so the window stays open
echo.
echo [INFO] Server has stopped.
pause
