@echo off
REM DOCX Rebuilder - Windows Context Menu Installer
REM Double-click this file to install the right-click menu option

echo ================================================
echo DOCX Rebuilder - Context Menu Installer
echo ================================================
echo.

REM Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://python.org
    pause
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
pip install -r "%~dp0..\requirements.txt" >nul 2>&1

REM Run the installer
echo.
echo Installing context menu...
python "%~dp0install_context_menu.py"

echo.
echo ================================================
echo Installation complete!
echo.
echo Right-click any .docx file and select
echo "Rebuild with DOCX Rebuilder" to use.
echo ================================================
echo.
pause
