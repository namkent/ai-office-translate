# BuildSetup.ps1
# Compiles Setup.cs and packages DLL dependencies and source C# code as embedded resources in setup.exe

$ErrorActionPreference = "Stop"

$scriptDir = $PSScriptRoot
$projectRoot = Split-Path $scriptDir -Parent

Write-Host "--------------------------------------------------------" -ForegroundColor Cyan
Write-Host "BUILDING NATIVE SINGLE-FILE INSTALLER (setup.exe)..." -ForegroundColor Cyan
Write-Host "--------------------------------------------------------" -ForegroundColor Cyan

# Locate csc.exe compiler
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $csc)) {
    Write-Error "Microsoft C# Compiler (csc.exe) not found on this machine! Cannot compile installer."
    Exit 1
}

# Source files and paths
$setupSource = Join-Path $scriptDir "Setup.cs"
$addinSource = Join-Path $scriptDir "AITranslateAddin.cs"
$extDll = Join-Path $scriptDir "Extensibility.dll"
$officeDll = Join-Path $scriptDir "Office.dll"
$outputExe = Join-Path $projectRoot "setup.exe"

# Stop any running setup.exe process to release lock
Stop-Process -Name setup -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

# Handle dynamic PNG to ICO icon conversion if icon.png is present in project root
$iconPng = Join-Path $projectRoot "icon.png"
$iconIco = Join-Path $scriptDir "icon.ico"
$iconArg = ""

if (Test-Path $iconPng) {
    Write-Host "Found icon.png. Converting to icon.ico for Win32 application icon..." -ForegroundColor Yellow
    $pngBytes = [System.IO.File]::ReadAllBytes($iconPng)
    $pngSize = $pngBytes.Length

    # 22-byte ICO header wrapping the raw PNG data (natively supported by Windows Explorer and WinForms)
    $icoHeader = [byte[]]@(
        0, 0,          # Reserved
        1, 0,          # Type (1 = ICO)
        1, 0,          # Image count (1)
        0,             # Width (0 = 256px)
        0,             # Height (0 = 256px)
        0,             # Color count (0 = >256 colors)
        0,             # Reserved
        1, 0,          # Color planes (1)
        32, 0,         # Bits per pixel (32)
        ($pngSize -band 0xFF), (($pngSize -shr 8) -band 0xFF), (($pngSize -shr 16) -band 0xFF), (($pngSize -shr 24) -band 0xFF), # PNG data size
        22, 0, 0, 0    # Offset where PNG data starts (22)
    )

    $icoBytes = New-Object byte[] ($icoHeader.Length + $pngBytes.Length)
    [Array]::Copy($icoHeader, 0, $icoBytes, 0, $icoHeader.Length)
    [Array]::Copy($pngBytes, 0, $icoBytes, $icoHeader.Length, $pngBytes.Length)
    [System.IO.File]::WriteAllBytes($iconIco, $icoBytes)
    
    $iconArg = "/win32icon:`"$iconIco`""
} else {
    Write-Host "icon.png not found. setup.exe will be built with the default system icon." -ForegroundColor Gray
}

# Read compile-time settings from .env file in the project root
$envPath = Join-Path $projectRoot ".env"
$defaultApiUrl = "https://localhost:3000"
$defaultToken = "secure-token-123"

if (Test-Path $envPath) {
    Write-Host "Reading compile-time configuration from .env..." -ForegroundColor Yellow
    $envLines = Get-Content $envPath
    $port = "3000"
    $token = "secure-token-123"
    $customApiUrl = ""

    foreach ($line in $envLines) {
        $trimmed = $line.Trim()
        if ($trimmed -and -not $trimmed.StartsWith("#") -and $trimmed.Contains("=")) {
            $idx = $trimmed.IndexOf('=')
            $key = $trimmed.Substring(0, $idx).Trim()
            $val = $trimmed.Substring($idx + 1).Trim()

            if ($key -eq "PORT") {
                $port = $val
            } elseif ($key -eq "CLIENT_ACCESS_TOKEN" -or $key -eq "CLIENT_TOKEN") {
                $token = $val
            } elseif ($key -eq "API_URL") {
                $customApiUrl = $val
            }
        }
    }

    if ($customApiUrl -ne "") {
        $defaultApiUrl = $customApiUrl
    } else {
        $defaultApiUrl = "https://localhost:$port"
    }
    $defaultToken = $token

    Write-Host "  -> Default API URL: $defaultApiUrl" -ForegroundColor Gray
    Write-Host "  -> Default Access Token: $defaultToken" -ForegroundColor Gray
}

Write-Host "Preprocessing source code with environment configurations..." -ForegroundColor Yellow

# Read original source codes
$setupCode = Get-Content $setupSource -Raw -Encoding UTF8
$addinCode = Get-Content $addinSource -Raw -Encoding UTF8

# Inject environment values into defaults
$setupCode = $setupCode.Replace('string apiUrl = "https://localhost:3000";', "string apiUrl = `"$defaultApiUrl`";")
$setupCode = $setupCode.Replace('string token = "secure-token-123";', "string token = `"$defaultToken`";")

$addinCode = $addinCode.Replace('protected string serverUrl = "https://localhost:3000";', "protected string serverUrl = `"$defaultApiUrl`";")
$addinCode = $addinCode.Replace('protected string clientToken = "secure-token-123";', "protected string clientToken = `"$defaultToken`";")

# Create temporary source files for compilation
$setupTemp = Join-Path $scriptDir "Setup_temp.cs"
$addinTemp = Join-Path $scriptDir "AITranslateAddin_temp.cs"

[System.IO.File]::WriteAllText($setupTemp, $setupCode, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($addinTemp, $addinCode, [System.Text.Encoding]::UTF8)

Write-Host "Embedding resources and compiling..." -ForegroundColor Yellow

# Compile from temporary files to keep source directory clean
if ($iconArg -ne "") {
    & $csc /target:winexe /out:"$outputExe" $iconArg `
        /resource:"$addinTemp,AITranslateAddin.cs" `
        /resource:"$extDll,Extensibility.dll" `
        /resource:"$officeDll,Office.dll" `
        /resource:"$iconIco,icon.ico" `
        /resource:"$iconPng,icon.png" `
        "$setupTemp"
} else {
    & $csc /target:winexe /out:"$outputExe" `
        /resource:"$addinTemp,AITranslateAddin.cs" `
        /resource:"$extDll,Extensibility.dll" `
        /resource:"$officeDll,Office.dll" `
        "$setupTemp"
}

$compileResult = $LASTEXITCODE

# Cleanup temporary preprocessed source files
Remove-Item $setupTemp -Force -ErrorAction SilentlyContinue
Remove-Item $addinTemp -Force -ErrorAction SilentlyContinue

if ($compileResult -eq 0) {
    Write-Host "--------------------------------------------------------" -ForegroundColor Green
    Write-Host "SUCCESS: $outputExe built successfully!" -ForegroundColor Green
    Write-Host "You can now distribute this single 'setup.exe' file." -ForegroundColor Green
    Write-Host "--------------------------------------------------------" -ForegroundColor Green
} else {
    Write-Error "Failed to build setup.exe"
    Exit 1
}
