@echo off
setlocal
title PDF Word Local Converter

cd /d "%~dp0"

set "PYTHON_CMD=python"
where python >nul 2>&1
if errorlevel 1 (
    where py >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python 3.8 or newer was not found.
        echo Download: https://www.python.org/downloads/
        pause
        exit /b 1
    )
    set "PYTHON_CMD=py -3"
)

echo ============================================
echo   PDF Word Local Converter
echo ============================================
echo.
echo [1/2] Checking Python dependencies...
call %PYTHON_CMD% -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo [WARNING] Some dependencies could not be installed.
    echo Run this command manually: %PYTHON_CMD% -m pip install -r requirements.txt
)

echo [2/2] Starting the converter...
echo.
call %PYTHON_CMD% pdf_word_converter.py
set "APP_EXIT=%ERRORLEVEL%"

if not "%APP_EXIT%"=="0" (
    echo.
    echo [ERROR] The converter exited with code %APP_EXIT%.
    pause
)

exit /b %APP_EXIT%
