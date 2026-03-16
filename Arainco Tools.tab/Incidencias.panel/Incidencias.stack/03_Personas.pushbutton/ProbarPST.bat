@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo Prueba de carga de contactos desde archivo PST
echo =============================================
echo.
if "%~1"=="" (
    echo Uso: Arrastra un archivo .pst sobre este .bat
    echo O: ProbarPST.bat "C:\ruta\archivo.pst"
    echo.
    set /p PSTFILE="Ruta del archivo PST: "
) else (
    set PSTFILE=%~1
)
if not exist "%PSTFILE%" (
    echo ERROR: No existe el archivo: %PSTFILE%
    pause
    exit /b 1
)
echo.
echo Archivo PST: %PSTFILE%
echo.
set OUTPUT=%TEMP%\personas_test_%RANDOM%.json
echo Salida JSON: %OUTPUT%
echo.
"%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -File "LeerOutlookContactos.ps1" -PstPath "%PSTFILE%" -OutputPath "%OUTPUT%" -TodosContactos
if errorlevel 1 (
    echo Intentando con PowerShell 64-bit...
    powershell -ExecutionPolicy Bypass -NoProfile -File "LeerOutlookContactos.ps1" -PstPath "%PSTFILE%" -OutputPath "%OUTPUT%" -TodosContactos
)
if errorlevel 1 (
    echo.
    echo ERROR al ejecutar el script. Asegurate de tener Outlook abierto.
    pause
    exit /b 1
)
echo.
echo Script ejecutado correctamente.
if exist "%OUTPUT%" (
    echo.
    echo Contenido del archivo de salida:
    type "%OUTPUT%"
    echo.
    del "%OUTPUT%"
) else (
    echo No se genero el archivo de salida.
)
echo.
pause
