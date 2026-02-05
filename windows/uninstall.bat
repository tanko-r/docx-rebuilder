@echo off
REM DOCX Rebuilder - Context Menu Uninstaller
REM Double-click this file to remove the right-click menu option

echo ================================================
echo DOCX Rebuilder - Context Menu Uninstaller
echo ================================================
echo.

python "%~dp0install_context_menu.py" --uninstall

echo.
pause
