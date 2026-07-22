# Sincroniza módulos canónicos de BIMTools.extension/scripts hacia scripts/ del pushbutton.
# Uso: .\sync_portable_scripts.ps1

$ErrorActionPreference = "Stop"
$PushbuttonDir = $PSScriptRoot
$RepoRoot = Resolve-Path (Join-Path $PushbuttonDir "..\..\..")
$Canonical = Join-Path $RepoRoot "scripts"
$Dest = Join-Path $PushbuttonDir "scripts"
$BootstrapSrc = Join-Path $PushbuttonDir "..\25_OrientacionMuroAlzado.pushbutton\bimtools_access_bootstrap.py"

$ToSync = @(
    "corporate_access.py",
    "bimtools_script_guard.py",
    "bimtools_instruction_dialog.py",
    "bimtools_ui_tokens.py",
    "bimtools_wpf_shell.py",
    "bimtools_wpf_dark_theme.py",
    "revit_wpf_window_position.py"
)

if (-not (Test-Path $Canonical)) {
    Write-Error "No se encontró carpeta canónica: $Canonical"
}

foreach ($name in $ToSync) {
    $src = Join-Path $Canonical $name
    if (-not (Test-Path $src)) {
        Write-Warning "Omitido (no existe en canónica): $name"
        continue
    }
    Copy-Item -Force $src (Join-Path $Dest $name)
    Write-Host "  + scripts/$name"
}

if (Test-Path $BootstrapSrc) {
    Copy-Item -Force $BootstrapSrc (Join-Path $PushbuttonDir "bimtools_access_bootstrap.py")
    Write-Host "  + bimtools_access_bootstrap.py"
} else {
    Write-Warning "No se encontró bootstrap portable de referencia: $BootstrapSrc"
}

Write-Host "Listo."
