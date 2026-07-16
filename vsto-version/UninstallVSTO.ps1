# UninstallVSTO.ps1
# Uninstalls VSTO Add-ins (No Administrator privileges required)

$ErrorActionPreference = "Stop"

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "UNINSTALLING VSTO ADD-INS..." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

$apps = @(
    @{ Name = "Excel"; RegNames = @("AITranslateExcel", "AITranslate.Excel") },
    @{ Name = "Word"; RegNames = @("AITranslateWord", "AITranslate.Word") },
    @{ Name = "PowerPoint"; RegNames = @("AITranslatePPT", "AITranslate.PPT") }
)

foreach ($app in $apps) {
    foreach ($regName in $app.RegNames) {
        $regPath = "HKCU:\Software\Microsoft\Office\$($app.Name)\Addins\$regName"
        if (Test-Path $regPath) {
            Write-Host "Removing registry keys for $($app.Name) ($regName)..." -ForegroundColor Yellow
            Remove-Item -Path $regPath -Recurse -Force
            Write-Host "[OK] Unregistered $($app.Name) ($regName) add-in." -ForegroundColor Green
        }
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
