@echo off
REM Recordatorios de incidencias (SMTP)
REM Detener con Ctrl+C

cd /d "%~dp0"

REM Buscar Python (py launcher o python directo)
where py >nul 2>&1
if %errorlevel% equ 0 (
    py recordatorios_smtp.py
) else (
    python recordatorios_smtp.py
)

if %errorlevel% neq 0 (
    echo.
    pause
)
