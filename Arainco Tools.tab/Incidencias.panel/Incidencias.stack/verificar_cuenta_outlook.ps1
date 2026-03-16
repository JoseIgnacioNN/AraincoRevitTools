# Verifica que la cuenta arainconotion@arainco.cl existe en Outlook CLASICO
# IMPORTANTE: Este script usa Outlook COM, que solo funciona con Outlook CLASICO.
# La nueva version de Outlook NO soporta COM. Las cuentas son independientes.
#
# Para que la cuenta aparezca aqui, debe estar agregada en Outlook CLASICO:
#   1. Abre Outlook CLASICO (no la nueva version)
#   2. Archivo > Configurar cuenta > Agregar cuenta
#   3. Ingresa arainconotion@arainco.cl
#
# Ejecutar: powershell -ExecutionPolicy Bypass -File verificar_cuenta_outlook.ps1

$cuentaBuscar = "jose.nunez@arainco.cl"

Write-Host "Verificando cuentas en Outlook CLASICO (COM)..." -ForegroundColor Cyan
Write-Host "Nota: La nueva version de Outlook usa un perfil distinto y no es accesible via COM." -ForegroundColor Gray
Write-Host ""

try {
    $ol = New-Object -ComObject Outlook.Application
    $cuentas = $ol.Session.Accounts
    
    Write-Host "Cuentas configuradas en Outlook CLASICO:" -ForegroundColor Cyan
    Write-Host ""
    
    $encontrada = $false
    for ($i = 1; $i -le $cuentas.Count; $i++) {
        $acc = $cuentas.Item($i)
        $smtp = $acc.SmtpAddress
        $display = $acc.DisplayName
        $user = $acc.UserName
        
        Write-Host "  [$i] DisplayName: $display"
        Write-Host "      SmtpAddress: $smtp"
        Write-Host "      UserName: $user"
        Write-Host ""
        
        if ($smtp -and $smtp -like "*$cuentaBuscar*") {
            $encontrada = $true
            Write-Host "  --> CUENTA ENCONTRADA: $cuentaBuscar" -ForegroundColor Green
        }
        if ($user -and $user -like "*$cuentaBuscar*") {
            $encontrada = $true
            Write-Host "  --> CUENTA ENCONTRADA (por UserName): $cuentaBuscar" -ForegroundColor Green
        }
    }
    
    Write-Host ""
    if ($encontrada) {
        Write-Host "RESULTADO: La cuenta $cuentaBuscar esta configurada en Outlook clasico. Se puede implementar el envio." -ForegroundColor Green
        exit 0
    } else {
        Write-Host "RESULTADO: La cuenta $cuentaBuscar NO fue encontrada en Outlook CLASICO." -ForegroundColor Red
        Write-Host ""
        Write-Host "Si la cuenta existe en la NUEVA version de Outlook:" -ForegroundColor Yellow
        Write-Host "  - Agregala tambien en Outlook CLASICO (Archivo > Agregar cuenta)" -ForegroundColor Yellow
        Write-Host "  - O usa Microsoft Graph API / SMTP como alternativa (sin Outlook)" -ForegroundColor Yellow
        exit 1
    }
} catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Posibles causas:" -ForegroundColor Yellow
    Write-Host "  - Outlook clasico no esta abierto (la nueva version no cuenta)" -ForegroundColor Yellow
    Write-Host "  - Outlook no esta instalado o hay error de permisos COM" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Sugerencia: Abre Outlook CLASICO (no la nueva version) y vuelve a ejecutar." -ForegroundColor Gray
    exit 1
}
