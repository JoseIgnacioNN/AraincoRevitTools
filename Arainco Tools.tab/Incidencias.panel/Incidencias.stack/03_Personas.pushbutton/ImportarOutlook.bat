@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Importando contactos de Outlook con @arainco.cl...
echo.
REM Outlook suele ser 32-bit: usar PowerShell 32-bit para coincidir
"%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -File "ImportarOutlook.ps1"
if errorlevel 1 (
    echo.
    echo Intentando con PowerShell 64-bit...
    powershell -ExecutionPolicy Bypass -NoProfile -File "ImportarOutlook.ps1"
)
echo.
echo Pulsa una tecla para cerrar...
pause >nul
