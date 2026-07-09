# Sincroniza módulos canónicos de BIMTools.extension/scripts hacia scripts/ del pushbutton.
# Uso: .\sync_portable_scripts.ps1
# No sobrescribe orientacion_muro_alzado_ui.py (solo vive en el paquete portable).

$ErrorActionPreference = "Stop"
$PushbuttonDir = $PSScriptRoot
$RepoRoot = Resolve-Path (Join-Path $PushbuttonDir "..\..\..")
$Canonical = Join-Path $RepoRoot "scripts"
$Dest = Join-Path $PushbuttonDir "scripts"

$ToSync = @(
    "orientacion_muro_alzado.py",
    "bimtools_ui_tokens.py",
    "bimtools_wpf_shell.py",
    "bimtools_wpf_dark_theme.py",
    "revit_wpf_window_position.py",
    "corporate_access.py",
    "bimtools_script_guard.py",
    "bimtools_instruction_dialog.py"
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

Write-Host "Listo. Revise orientacion_muro_alzado_ui.py manualmente si cambió la UI."
