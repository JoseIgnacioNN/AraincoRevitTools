@echo off
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel% equ 0 (py enviar_prueba_email.py) else (python enviar_prueba_email.py)
echo.
pause
