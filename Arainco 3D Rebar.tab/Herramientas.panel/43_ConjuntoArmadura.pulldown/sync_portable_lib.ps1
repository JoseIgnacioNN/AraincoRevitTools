# Sincroniza scripts/lib/ desde 01_SeleccionarConjunto hacia los otros botones.
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "01_SeleccionarConjunto.pushbutton\scripts\lib"
$targets = @(
    "02_OcultarConjunto.pushbutton\scripts\lib",
    "03_EliminarConjunto.pushbutton\scripts\lib"
)
foreach ($rel in $targets) {
    $dst = Join-Path $root $rel
    New-Item -ItemType Directory -Force -Path $dst | Out-Null
    Copy-Item (Join-Path $src "*.py") $dst -Force
    Write-Host "Synced -> $rel"
}
