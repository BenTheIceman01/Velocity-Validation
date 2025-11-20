@echo off
title HD Supply Velocity Validator
color 0E
cls

echo ========================================
echo   HD SUPPLY VELOCITY VALIDATOR
echo   Developed by: Ben F. Benjamaa
echo ========================================
echo.
echo Starting application...
echo.

python velocity_validator_app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Application failed to start!
    echo Please ensure Python is installed and dependencies are met.
    echo.
    echo Run: pip install -r requirements.txt
    echo.
    pause
)
