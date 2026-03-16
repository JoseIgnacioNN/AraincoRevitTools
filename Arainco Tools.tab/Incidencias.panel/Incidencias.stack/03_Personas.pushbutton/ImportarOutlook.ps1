# Importar contactos de Outlook con @arainco.cl al directorio Personas
# No requiere dependencias externas - usa COM nativo de Windows

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {

$ISSUES_DIR = "Y:\00_SERVIDOR DE INCIDENCIAS"
$PERSONAS_FILE = Join-Path $ISSUES_DIR "personas.json"
$DOMINIO = "@arainco.cl"
$MAILBOX_BUSCAR = "jose.nunez@arainco.cl"  # Buzón donde buscar contactos

# Conectar a Outlook: primero intentar instancia ya abierta, luego crear nueva
$outlook = $null
try {
    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
} catch { }
if (-not $outlook) {
    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        Write-Host "ERROR: No se pudo conectar con Outlook. Asegurate de tener Outlook abierto."
        Write-Host "Si Outlook esta abierto: prueba cerrar Outlook, ejecutar este script, y abrirlo cuando lo pida."
        Write-Host "O usa la version de PowerShell que coincida con Outlook (32 o 64 bits)."
        exit 1
    }
}
if (-not $outlook) {
    Write-Host "ERROR: No se pudo conectar con Outlook. Asegurate de tener Outlook abierto."
    exit 1
}

$ns = $outlook.GetNamespace("MAPI")
$contactsFolder = $null
foreach ($store in $ns.Stores) {
    if ($store.DisplayName -like "*$MAILBOX_BUSCAR*") {
        try {
            $contactsFolder = $store.GetDefaultFolder(10)  # 10 = olFolderContacts
            Write-Host "Buscando en buzón: $($store.DisplayName)"
            break
        } catch { continue }
    }
}
if (-not $contactsFolder) {
    $contactsFolder = $ns.GetDefaultFolder(10)
    Write-Host "Buzón $MAILBOX_BUSCAR no encontrado. Usando buzón por defecto."
}

function Get-EmailsFromContact {
    param($contact)
    $emails = @()
    foreach ($attr in @("Email1Address", "Email2Address", "Email3Address")) {
        $val = $contact.$attr
        if ($val -and $val.ToString().Trim()) { $emails += $val.ToString().Trim() }
    }
    # Fallback: Email1DisplayName puede tener formato "Nombre (email@dominio.com)"
    if ($emails.Count -eq 0) {
        foreach ($attr in @("Email1DisplayName", "Email2DisplayName", "Email3DisplayName")) {
            $val = $contact.$attr
            if ($val -and $val.ToString().Trim()) {
                if ($val -match '([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})') {
                    $emails += $matches[1]
                }
            }
        }
    }
    # Fallback: PropertyAccessor para contactos Exchange (Email1Address a veces vacio)
    if ($emails.Count -eq 0) {
        try {
            $pa = $contact.PropertyAccessor
            $schema = "http://schemas.microsoft.com/mapi/proptag/0x8083001E"
            $e = $pa.GetProperty($schema)
            if ($e -and $e.ToString().Trim()) { $emails += $e.ToString().Trim() }
        } catch { }
    }
    return $emails
}

function Get-ContactsFromFolder {
    param($folder, [ref]$totalItems)
    $found = @()
    try {
        $items = $folder.Items
        $count = $items.Count
        $totalItems.Value += $count
        for ($i = 1; $i -le $count; $i++) {
            try {
                $contact = $items.Item($i)
                if ($contact.Class -ne 40) { continue }
                $nombre = $contact.FullName
                if ([string]::IsNullOrWhiteSpace($nombre)) {
                    $nombre = ($contact.FirstName + " " + $contact.LastName).Trim()
                    if ([string]::IsNullOrWhiteSpace($nombre)) { $nombre = "(Sin nombre)" }
                }
                $emails = Get-EmailsFromContact $contact
                foreach ($email in $emails) {
                    if ($email -like "*$DOMINIO*") {
                        $found += @{ nombre = $nombre.Trim(); email = $email }
                        break
                    }
                }
            } catch { continue }
        }
        # Subcarpetas
        foreach ($sub in $folder.Folders) {
            $found += Get-ContactsFromFolder $sub $totalItems
        }
    } catch { }
    return $found
}

$totalItems = 0
$contactos = Get-ContactsFromFolder $contactsFolder ([ref]$totalItems)

Write-Host "Revisados $totalItems elementos en Contactos. Encontrados $($contactos.Count) con $DOMINIO"

if ($contactos.Count -eq 0) {
    Write-Host "No hay contactos con direccion $DOMINIO en Outlook."
    exit 0
}

# Cargar existentes
$existentes = @()
if (Test-Path $PERSONAS_FILE) {
    $raw = Get-Content $PERSONAS_FILE -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($raw) {
        $existentes = @($raw)
    }
}
$emailsExistentes = @{}
$nombresExistentes = @{}
foreach ($p in $existentes) {
    $emailsExistentes[$p.email.ToLower()] = $true
    $nombresExistentes[$p.nombre.ToLower()] = $true
}

# Agregar nuevos
$agregados = 0
foreach ($c in $contactos) {
    $emailLower = $c.email.ToLower()
    $nombreLower = $c.nombre.ToLower()
    if (-not $emailsExistentes.ContainsKey($emailLower) -and -not $nombresExistentes.ContainsKey($nombreLower)) {
        $existentes += [PSCustomObject]@{ nombre = $c.nombre; email = $c.email }
        $emailsExistentes[$emailLower] = $true
        $nombresExistentes[$nombreLower] = $true
        $agregados++
        Write-Host "  + Agregado: $($c.nombre) <$($c.email)>"
    }
}

if ($agregados -gt 0) {
    New-Item -ItemType Directory -Force -Path $ISSUES_DIR | Out-Null
    $existentes | ConvertTo-Json -Depth 3 | Out-File $PERSONAS_FILE -Encoding UTF8
    Write-Host ""
    Write-Host "Listo. Se agregaron $agregados personas al directorio."
} else {
    Write-Host ""
    Write-Host "Todos los contactos ya estaban en el directorio."
}

} catch {
    Write-Error $_.Exception.Message
    exit 1
}
