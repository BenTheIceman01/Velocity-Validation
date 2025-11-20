@echo off
title HD Supply Velocity Validator - Build Executable
color 0E
cls

echo ========================================
echo   HD SUPPLY VELOCITY VALIDATOR
echo   EXECUTABLE BUILDER
echo   Developed by: Ben F. Benjamaa
echo ========================================
echo.
echo This will create a standalone executable file.
echo.
echo Building executable...
echo.

pyinstaller --clean --onefile velocity_validator.spec

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo   Build completed successfully!
    echo ========================================
    echo.
    echo Executable location:
    echo   dist\HD_Supply_Velocity_Validator.exe
    echo.
    echo You can now distribute this single file
    echo without requiring Python installation!
    echo.
) else (
    echo.
    echo ========================================
    echo   Build failed!
    echo ========================================
    echo.
    echo Please ensure PyInstaller is installed:
    echo   pip install pyinstaller
    echo.
)

pause
