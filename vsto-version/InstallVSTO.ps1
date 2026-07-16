# InstallVSTO.ps1
# Installs VSTO Add-ins for the current user (No Administrator privileges required)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "INSTALLING VSTO ADD-INS (NO ADMIN PRIVILEGES REQUIRED)..." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

# 1. Import certificate to the current user's Trusted Root & Trusted Publisher stores
$pfxFile = Join-Path $scriptDir "AITranslate.pfx"
$passwordText = "secure-password-123"

if (Test-Path $pfxFile) {
    Write-Host "Found certificate. Importing to Trusted Publisher..." -ForegroundColor Yellow
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $pwd = ConvertTo-SecureString -String $passwordText -Force -AsPlainText
        $cert.Import($pfxFile, $pwd, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet)
        
        # Add to Trusted Publishers for CurrentUser (No Admin required)
        $storeTP = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher", "CurrentUser")
        $storeTP.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $storeTP.Add($cert)
        $storeTP.Close()

        # Add to Trusted Root for CurrentUser (No Admin required)
        $storeRoot = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
        $storeRoot.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $storeRoot.Add($cert)
        $storeRoot.Close()
        
        Write-Host "[OK] Trusted the certificate successfully." -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not automatically trust the certificate: $($_.Exception.Message)"
        Write-Warning "Office may prompt a security warning on startup."
    }
} else {
    Write-Warning "AITranslate.pfx was not found in the script directory!"
}

# 2. Register Add-ins in Registry (CurrentUser)
$apps = @(
    @{ Name = "Excel"; Proj = "AITranslateExcel"; RegName = "AITranslateExcel"; Desc = "Excel Translation Add-in using AI" },
    @{ Name = "Word"; Proj = "AITranslateWord"; RegName = "AITranslateWord"; Desc = "Word Translation Add-in using AI" },
    @{ Name = "PowerPoint"; Proj = "AITranslatePPT"; RegName = "AITranslatePPT"; Desc = "PowerPoint Translation Add-in using AI" }
)

foreach ($app in $apps) {
    $vstoPath = Join-Path (Join-Path $scriptDir $app.Proj) "$($app.Proj).vsto"
    
    if (Test-Path $vstoPath) {
        $manifestPath = "file:///" + $vstoPath.Replace("\", "/") + "|vstolocal"
        $regPath = "HKCU:\Software\Microsoft\Office\$($app.Name)\Addins\$($app.RegName)"
        
        Write-Host "Registering add-in for $($app.Name)..." -ForegroundColor Yellow
        
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "AI Office Translate"
        Set-ItemProperty -Path $regPath -Name "Description" -Value $app.Desc
        Set-ItemProperty -Path $regPath -Name "Manifest" -Value $manifestPath
        Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
        
        Write-Host "[OK] Registered add-in for $($app.Name)!" -ForegroundColor Green
    } else {
        Write-Host "[SKIP] VSTO file not found for $($app.Proj)" -ForegroundColor Gray
    }
}

# 3. Create settings directory and file
$appDataDir = Join-Path $env:APPDATA "AITranslateAddin"
if (-not (Test-Path $appDataDir)) {
    New-Item -ItemType Directory -Path $appDataDir | Out-Null
}

$settingsFile = Join-Path $appDataDir "settings.txt"
$apiUrl = "http://localhost:3000"
$token = "secure-token-123"

$envPath = Join-Path $PSScriptRoot "..\.env"
if (Test-Path $envPath) {
    Get-Content $envPath | ForEach-Object {
        if ($_ -match "^\s*(PORT)\s*=\s*(.+)") {
            $apiUrl = "http://localhost:$($Matches[2].Trim().Trim("'").Trim('"'))"
        }
        elseif ($_ -match "^\s*(API_URL|BACKEND_URL|HOST_URL)\s*=\s*(.+)") {
            $apiUrl = $Matches[2].Trim().Trim("'").Trim('"')
        }
        elseif ($_ -match "^\s*(CLIENT_TOKEN|CLIENT_ACCESS_TOKEN)\s*=\s*(.+)") {
            $token = $Matches[2].Trim().Trim("'").Trim('"')
        }
    }
}

Set-Content -Path $settingsFile -Value @("API_URL=$apiUrl", "TOKEN=$token")
Write-Host "[OK] Written configuration file at: $settingsFile" -ForegroundColor Green

Write-Host "========================================================" -ForegroundColor Green
Write-Host "INSTALLATION COMPLETED!" -ForegroundColor Green
Write-Host "Please restart Excel, Word, and PowerPoint." -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
