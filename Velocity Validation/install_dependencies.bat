@echo off
title HD Supply Velocity Validator - Installer
color 0E
cls

echo ========================================
echo   HD SUPPLY VELOCITY VALIDATOR
echo   DEPENDENCY INSTALLER
echo   Developed by: Ben F. Benjamaa
echo ========================================
echo.
echo Installing required Python packages...
echo.

pip install -r requirements.txt

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo   Installation completed successfully!
    echo ========================================
    echo.
    echo You can now run the application using:
    echo   run_app.bat
    echo.
) else (
    echo.
    echo ========================================
    echo   Installation failed!
    echo ========================================
    echo.
    echo Please check your Python installation.
    echo.
)

pause
