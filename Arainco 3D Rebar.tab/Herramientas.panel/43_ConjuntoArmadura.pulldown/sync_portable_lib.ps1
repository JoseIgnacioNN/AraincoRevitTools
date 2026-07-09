# Sincroniza lib/ y bootstrap desde 01_SeleccionarConjunto hacia los otros botones.
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcBtn = Join-Path $root "01_SeleccionarConjunto.pushbutton"
$srcLib = Join-Path $srcBtn "scripts\lib"
$targets = @(
    "02_OcultarConjunto.pushbutton",
    "03_EliminarConjunto.pushbutton"
)
foreach ($rel in $targets) {
    $dstBtn = Join-Path $root $rel
    $dstLib = Join-Path $dstBtn "scripts\lib"
    New-Item -ItemType Directory -Force -Path $dstLib | Out-Null
    Copy-Item (Join-Path $srcLib "*.py") $dstLib -Force
    Copy-Item (Join-Path $srcBtn "bimtools_access_bootstrap.py") $dstBtn -Force
    Write-Host "Synced -> $rel"
}
