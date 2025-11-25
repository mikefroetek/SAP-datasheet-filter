@echo off
echo ===================================
echo    Excel Level Processor
echo ===================================
echo.

:: Change to the script directory
cd /d "%~dp0"

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python first
    pause
    exit /b 1
)

:: Install required packages
echo Installing required packages...
pip install pandas openpyxl pyxlsb xlrd

echo.
echo Running Excel processor...
echo Current directory: %CD%
echo.

:: Run the Python script
python simple_excel_processor.py

echo.
echo Press any key to exit...
pause >nul