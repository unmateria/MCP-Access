# enable_vba_trust.ps1
# Activa "Confiar en el acceso al modelo de objetos de proyectos VBA" para Microsoft Access.
# El valor se guarda en HKCU (no requiere privilegios de administrador).

$ErrorActionPreference = "Stop"

$versions = @("16.0", "15.0", "14.0", "12.0")
$applied  = @()

foreach ($ver in $versions) {
    $parent = "HKCU:\Software\Microsoft\Office\$ver\Access"
    if (-not (Test-Path $parent)) { continue }

    $regPath = "$parent\Security"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord
    $applied += "Office $ver  ->  $regPath\AccessVBOM = 1"
}

if ($applied.Count -eq 0) {
    Write-Warning "No se encontro ninguna instalacion de Microsoft Access en el registro."
    exit 1
}

Write-Host ""
Write-Host "VBA Trust activado para:" -ForegroundColor Green
$applied | ForEach-Object { Write-Host "  $_" -ForegroundColor Cyan }
Write-Host ""
Write-Host "Si Access estaba abierto, cierra y vuelve a abrir para que el cambio surta efecto." -ForegroundColor Yellow
