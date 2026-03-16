# Lee contactos de Outlook: desde archivo PST o desde buzón jose.nunez@arainco.cl
# Escribe resultado en archivo JSON (para que Python lo lea sin pywin32)
# Uso:
#   Desde PST:  .\LeerOutlookContactos.ps1 -PstPath "C:\ruta\archivo.pst" -OutputPath "C:\temp\contactos.json"
#   Desde buzón: .\LeerOutlookContactos.ps1 -OutputPath "C:\temp\contactos.json"

param(
    [string]$PstPath = "",
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    [string]$Dominio = "@arainco.cl",
    [string]$MailboxBuscar = "jose.nunez@arainco.cl",
    [switch]$TodosContactos = $false,
    [string]$DebugLog = ""  # Si se especifica, escribe diagnostico
)

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Get-EmailsFromContact {
    param($contact)
    $emails = @()
    foreach ($attr in @("Email1Address", "Email2Address", "Email3Address")) {
        $val = $contact.$attr
        if ($val -and $val.ToString().Trim()) { $emails += $val.ToString().Trim() }
    }
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
    param($folder)
    $found = @()
    try {
        $items = $folder.Items
        try {
            foreach ($item in $items) {
                try {
                    $isContact = ($item.Class -eq 40)
                    if (-not $isContact) {
                        $msgClass = $item.MessageClass
                        if ($msgClass -and $msgClass -like "*Contact*") { $isContact = $true }
                    }
                    if (-not $isContact) { continue }
                    $nombre = $item.FullName
                    if ([string]::IsNullOrWhiteSpace($nombre)) {
                        $nombre = (($item.FirstName + " " + $item.LastName) -replace "\s+", " ").Trim()
                        if ([string]::IsNullOrWhiteSpace($nombre)) { $nombre = "(Sin nombre)" }
                    }
                    $emails = Get-EmailsFromContact $item
                    $added = $false
                    foreach ($email in $emails) {
                        if ($TodosContactos -or $email -like "*$Dominio*") {
                            $found += [PSCustomObject]@{ nombre = $nombre.Trim(); email = $email }
                            $added = $true
                            break
                        }
                    }
                    if (-not $added -and $TodosContactos -and $emails.Count -eq 0 -and $nombre -ne "(Sin nombre)") {
                        $found += [PSCustomObject]@{ nombre = $nombre.Trim(); email = "" }
                    }
                } catch { continue }
            }
        } catch {
            $count = $items.Count
            for ($i = 1; $i -le $count; $i++) {
                try {
                    $item = $items.Item($i)
                    $isContact = ($item.Class -eq 40)
                    if (-not $isContact) {
                        $msgClass = $item.MessageClass
                        if ($msgClass -and $msgClass -like "*Contact*") { $isContact = $true }
                    }
                    if (-not $isContact) { continue }
                    $nombre = $item.FullName
                    if ([string]::IsNullOrWhiteSpace($nombre)) {
                        $nombre = (($item.FirstName + " " + $item.LastName) -replace "\s+", " ").Trim()
                        if ([string]::IsNullOrWhiteSpace($nombre)) { $nombre = "(Sin nombre)" }
                    }
                    $emails = Get-EmailsFromContact $item
                    $added = $false
                    foreach ($email in $emails) {
                        if ($TodosContactos -or $email -like "*$Dominio*") {
                            $found += [PSCustomObject]@{ nombre = $nombre.Trim(); email = $email }
                            $added = $true
                            break
                        }
                    }
                    if (-not $added -and $TodosContactos -and $emails.Count -eq 0 -and $nombre -ne "(Sin nombre)") {
                        $found += [PSCustomObject]@{ nombre = $nombre.Trim(); email = "" }
                    }
                } catch { continue }
            }
        }
        try {
            foreach ($sub in $folder.Folders) {
                $found += Get-ContactsFromFolder $sub
            }
        } catch {
            $foldersCount = $folder.Folders.Count
            for ($j = 1; $j -le $foldersCount; $j++) {
                try {
                    $sub = $folder.Folders.Item($j)
                    $found += Get-ContactsFromFolder $sub
                } catch { continue }
            }
        }
    } catch { }
    return $found
}

function Find-ContactsFolder {
    param($folder)
    try {
        foreach ($sub in $folder.Folders) {
            if ($sub.DefaultItemType -eq 40) { return $sub }
            $found = Find-ContactsFolder $sub
            if ($found) { return $found }
        }
    } catch { }
    return $null
}

# Fallback: usa GetTable y lee datos de la fila (schema MAPI) o GetItemFromID
function Get-ContactsViaTable {
    param($folder, $namespace)
    $found = @()
    $storeId = $null
    try { $storeId = $folder.Store.StoreID } catch { }
    $schemaDisplay = "http://schemas.microsoft.com/mapi/proptag/0x3001001E"
    $schemaEmail = "http://schemas.microsoft.com/mapi/proptag/0x8083001E"
    try {
        $table = $null
        try { $table = $folder.GetTable($null, 0) } catch { }
        if (-not $table) { return $found }
        try { $table.Columns.Add("EntryID") } catch { }
        try { $table.Columns.Add("MessageClass") } catch { }
        try { $table.Columns.Add($schemaDisplay) } catch { }
        try { $table.Columns.Add($schemaEmail) } catch { }
        try { $table.Columns.Add("FullName") } catch { }
        try { $table.Columns.Add("Email1Address") } catch { }
        while (-not $table.EndOfTable) {
            try {
                $row = $table.GetNextRow()
                $msgClass = $row["MessageClass"]
                if (-not $msgClass -or $msgClass -notlike "*Contact*") { continue }
                $nombre = $row[$schemaDisplay]
                if ([string]::IsNullOrWhiteSpace($nombre)) { $nombre = $row["FullName"] }
                $email = $row[$schemaEmail]
                if ([string]::IsNullOrWhiteSpace($email)) { $email = $row["Email1Address"] }
                $item = $null
                if (([string]::IsNullOrWhiteSpace($nombre) -or [string]::IsNullOrWhiteSpace($email)) -and $row["EntryID"]) {
                    if ($storeId) { try { $item = $namespace.GetItemFromID($row["EntryID"], $storeId) } catch { } }
                    if (-not $item) { try { $item = $namespace.GetItemFromID($row["EntryID"]) } catch { } }
                    if ($item) {
                        if ([string]::IsNullOrWhiteSpace($nombre)) {
                            $nombre = $item.FullName
                            if ([string]::IsNullOrWhiteSpace($nombre)) { $nombre = (($item.FirstName + " " + $item.LastName) -replace "\s+", " ").Trim() }
                        }
                        if ([string]::IsNullOrWhiteSpace($email)) {
                            $emails = Get-EmailsFromContact $item
                            if ($emails.Count -gt 0) { $email = $emails[0] }
                        }
                    }
                }
                $nombre = if ([string]::IsNullOrWhiteSpace($nombre)) { "(Sin nombre)" } else { $nombre.Trim() }
                $email = if ([string]::IsNullOrWhiteSpace($email)) { "" } else { ($email -replace "\s+", "").Trim() }
                if ($TodosContactos -or $email -like "*$Dominio*" -or ([string]::IsNullOrWhiteSpace($email) -and $nombre -ne "(Sin nombre)")) {
                    $found += [PSCustomObject]@{ nombre = $nombre; email = $email }
                }
            } catch { continue }
        }
    } catch { }
    return $found
}

# Recorre todas las carpetas y aplica GetTable en cada una (fallback para PST con Items vacío)
function Get-ContactsFromFolderViaTable {
    param($folder, $namespace)
    $found = @()
    try {
        $found += Get-ContactsViaTable $folder $namespace
        foreach ($sub in $folder.Folders) {
            $found += Get-ContactsFromFolderViaTable $sub $namespace
        }
    } catch { }
    return $found
}

try {
    $outlook = $null
    try {
        $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch { }
    if (-not $outlook) {
        try {
            $outlook = New-Object -ComObject Outlook.Application
        } catch {
            Write-Error "No se pudo conectar con Outlook. Asegurate de tener Outlook instalado y abierto."
            exit 1
        }
    }

    $ns = $outlook.GetNamespace("MAPI")
    $contactsFolder = $null
    $rootToRemove = $null

    if ($PstPath -and (Test-Path $PstPath)) {
        try {
            $pstAbs = (Resolve-Path -LiteralPath $PstPath -ErrorAction Stop).Path
        } catch {
            $pstAbs = [System.IO.Path]::GetFullPath($PstPath)
        }
        $storeCountBefore = $ns.Stores.Count
        try {
            $ns.AddStore($pstAbs)
        } catch {
            Write-Error "No se pudo abrir el archivo PST. Asegurate de que no este en uso y que Outlook este abierto. Detalle: $_"
            exit 1
        }
        $rootPst = $null
        for ($i = 1; $i -le $ns.Stores.Count; $i++) {
            try {
                $store = $ns.Stores.Item($i)
                if (-not $store.IsDataFileStore) { continue }
                $fp = $store.FilePath
                if ($fp) {
                    $fpNorm = [System.IO.Path]::GetFullPath($fp.Trim()).ToLower()
                    $pstNorm = $pstAbs.ToLower()
                    if ($fpNorm -eq $pstNorm) {
                        $rootPst = $store.GetRootFolder()
                        break
                    }
                    if ($fpNorm.EndsWith($pstNorm) -or $pstNorm.EndsWith($fpNorm)) {
                        $rootPst = $store.GetRootFolder()
                        break
                    }
                }
            } catch { continue }
        }
        if (-not $rootPst) {
            if ($ns.Stores.Count -gt $storeCountBefore) {
                $rootPst = $ns.Stores.Item($ns.Stores.Count).GetRootFolder()
            }
        }
        if (-not $rootPst) {
            for ($i = $ns.Stores.Count; $i -ge 1; $i--) {
                try {
                    $store = $ns.Stores.Item($i)
                    if ($store.IsDataFileStore) {
                        $rootPst = $store.GetRootFolder()
                        break
                    }
                } catch { continue }
            }
        }
        if (-not $rootPst) {
            Write-Error "No se pudo acceder a la raiz del archivo PST. Prueba cerrar el PST si esta abierto en Outlook."
            exit 1
        }
        $rootToRemove = $rootPst
        $contactsFolder = $null
        try {
            $store = $rootPst.Store
            if ($store) {
                $contactsFolder = $store.GetDefaultFolder(10)
            }
        } catch { }
        if (-not $contactsFolder) {
            $contactsFolder = Find-ContactsFolder $rootPst
        }
        if (-not $contactsFolder) {
            try { $contactsFolder = $rootPst.Folders.Item("Contactos") } catch { }
        }
        if (-not $contactsFolder) {
            try { $contactsFolder = $rootPst.Folders.Item("Contacts") } catch { }
        }
        if (-not $contactsFolder) {
            for ($k = 1; $k -le $rootPst.Folders.Count; $k++) {
                try {
                    $sub = $rootPst.Folders.Item($k)
                    if ($sub.Name -match "ontact") {
                        $contactsFolder = $sub
                        break
                    }
                } catch { continue }
            }
        }
        if (-not $contactsFolder) {
            $contactsFolder = $rootPst
        }
        if ($contactsFolder.Items.Count -eq 0 -and $contactsFolder.Folders.Count -eq 0) {
            $contactsFolder = $rootPst
        }
    } else {
        foreach ($store in $ns.Stores) {
            if ($store.DisplayName -like "*$MailboxBuscar*") {
                try {
                    $contactsFolder = $store.GetDefaultFolder(10)
                    break
                } catch { continue }
            }
        }
        if (-not $contactsFolder) {
            $contactsFolder = $ns.GetDefaultFolder(10)
        }
    }

    $contactos = Get-ContactsFromFolder $contactsFolder
    $usedTableFallback = $false
    if ($contactos.Count -eq 0 -and $contactsFolder) {
        $contactos = Get-ContactsFromFolderViaTable $contactsFolder $ns
        if ($contactos.Count -gt 0) { $usedTableFallback = $true }
    }

    if ($DebugLog -and $rootPst) {
        $dbg = @()
        $dbg += "PST: $PstPath"
        $dbg += "Carpeta usada: $($contactsFolder.Name) (DefaultItemType: $($contactsFolder.DefaultItemType))"
        $dbg += "ContactsFolder.Items.Count: $($contactsFolder.Items.Count)"
        $dbg += "ContactsFolder.Folders.Count: $($contactsFolder.Folders.Count)"
        $dbg += "Contactos encontrados: $($contactos.Count)"
        if ($usedTableFallback) { $dbg += "(via GetTable fallback)" }
        $contactsSub = $null
        for ($k = 1; $k -le $rootPst.Folders.Count; $k++) {
            try {
                $f = $rootPst.Folders.Item($k)
                if ($f.Name -match "ontact") { $contactsSub = $f; break }
            } catch { }
        }
        if ($contactsSub) {
            $dbg += "`nSubcarpetas de Contactos:"
            for ($k = 1; $k -le $contactsSub.Folders.Count; $k++) {
                try {
                    $sub = $contactsSub.Folders.Item($k)
                    $c = 0
                    try { $c = $sub.Items.Count } catch { }
                    $dbg += "  - $($sub.Name): $c items"
                } catch { }
            }
        }
        function Get-FolderTree { param($f, $indent)
            $info = @()
            try {
                for ($j = 1; $j -le $f.Folders.Count; $j++) {
                    $sub = $f.Folders.Item($j)
                    $items = 0
                    try { $items = $sub.Items.Count } catch { }
                    $info += "${indent}- $($sub.Name): $items items (Type: $($sub.DefaultItemType))"
                    $info += Get-FolderTree $sub "$indent  "
                }
            } catch { }
            return $info
        }
        $dbg += "`nEstructura de carpetas:"
        $dbg += Get-FolderTree $rootPst "  "
        try { $dbg -join "`n" | Out-File $DebugLog -Encoding UTF8 } catch { }
    }

    if ($rootToRemove) {
        try { $ns.RemoveStore($rootToRemove) } catch { }
    }

    $result = @($contactos | ForEach-Object { @{ nombre = $_.nombre; email = $_.email } })
    if ($result.Count -eq 0) {
        "[]" | Out-File $OutputPath -Encoding UTF8
    } else {
        $json = $result | ConvertTo-Json -Depth 3 -Compress
        if ($result.Count -eq 1) { $json = "[$json]" }
        $json | Out-File $OutputPath -Encoding UTF8
    }
    exit 0

} catch {
    Write-Error $_.Exception.Message
    exit 1
}
