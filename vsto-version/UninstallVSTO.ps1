# UninstallVSTO.ps1
# Uninstalls VSTO Add-ins (No Administrator privileges required)

$ErrorActionPreference = "Stop"

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "UNINSTALLING VSTO ADD-INS..." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

$apps = @(
    @{ Name = "Excel"; RegName = "AITranslate.Excel" },
    @{ Name = "Word"; RegName = "AITranslate.Word" },
    @{ Name = "PowerPoint"; RegName = "AITranslate.PPT" }
)

foreach ($app in $apps) {
    $regPath = "HKCU:\Software\Microsoft\Office\$($app.Name)\Addins\$($app.RegName)"
    
    if (Test-Path $regPath) {
        Write-Host "Removing registry keys for $($app.Name)..." -ForegroundColor Yellow
        Remove-Item -Path $regPath -Recurse -Force
        Write-Host "[OK] Unregistered $($app.Name) add-in." -ForegroundColor Green
    } else {
        Write-Host "[SKIP] No registration found for $($app.Name)." -ForegroundColor Gray
    }
}

# Clean AppData config directory
$appDataDir = Join-Path $env:APPDATA "AITranslateAddin"
if (Test-Path $appDataDir) {
    try {
        Write-Host "Cleaning AppData config folder..." -ForegroundColor Yellow
        Remove-Item -Path $appDataDir -Recurse -Force
        Write-Host "[OK] Cleared config folder." -ForegroundColor Green
    } catch {
        Write-Warning "Could not completely delete AppData folder (files may be locked)."
    }
}

Write-Host "========================================================" -ForegroundColor Green
Write-Host "UNINSTALLATION COMPLETED!" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
