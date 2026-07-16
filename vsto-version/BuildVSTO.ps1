# BuildVSTO.ps1
# Automatically generates self-signed certificate and compiles VSTO projects using MSBuild.

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "BUILDING VSTO ADD-IN PROJECTS (EXCEL, WORD, PPT)..." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

# 1. Create Self-Signed Code Signing Certificate (.pfx)
$pfxFile = Join-Path $scriptDir "AITranslate.pfx"
$passwordText = "secure-password-123"

if (-not (Test-Path $pfxFile)) {
    Write-Host "Generating self-signed code signing certificate..." -ForegroundColor Yellow
    try {
        # Create Certificate using PowerShell
        $cert = New-SelfSignedCertificate -Type CodeSigning `
                                          -Subject "CN=AI Office Translate, O=Namkent, C=VN" `
                                          -KeyUsage DigitalSignature `
                                          -FriendlyName "AI Office Translate Certificate" `
                                          -CertStoreLocation "Cert:\CurrentUser\My" `
                                          -NotAfter (Get-Date).AddYears(5)
        
        $pwd = ConvertTo-SecureString -String $passwordText -Force -AsPlainText
        Export-PfxCertificate -Cert $cert -FilePath $pfxFile -Password $pwd
        
        # Import to CurrentUser My store so MSBuild doesn't prompt for password
        Import-PfxCertificate -FilePath $pfxFile -CertStoreLocation "Cert:\CurrentUser\My" -Password $pwd | Out-Null
        
        Write-Host "-> Created and imported certificate: $pfxFile" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not automatically generate self-signed certificate. Will fallback if error occurs."
    }
} else {
    Write-Host "-> Certificate already exists: $pfxFile" -ForegroundColor Gray
    try {
        $pwd = ConvertTo-SecureString -String $passwordText -Force -AsPlainText
        Import-PfxCertificate -FilePath $pfxFile -CertStoreLocation "Cert:\CurrentUser\My" -Password $pwd | Out-Null
        Write-Host "-> Verified and imported existing certificate into store." -ForegroundColor Green
    } catch {
        Write-Warning "Could not import existing certificate: $($_.Exception.Message)"
    }
}

# Copy the .pfx for each project to use for manifest signing
$projects = @("AITranslateExcel", "AITranslateWord", "AITranslatePPT")
foreach ($project in $projects) {
    $destPfx = Join-Path (Join-Path $scriptDir $project) "$($project)_TemporaryKey.pfx"
    Copy-Item $pfxFile -Destination $destPfx -Force
}

# 2. Locate MSBuild.exe
$msbuild = ""
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $vsPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
    if ($vsPath) {
        $msbuild = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
    }
}

if (-not (Test-Path $msbuild)) {
    $paths = @(
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            $msbuild = $p
            break
        }
    }
}

if (-not (Test-Path $msbuild)) {
    $msbuild = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
}

if (-not (Test-Path $msbuild)) {
    Write-Error "MSBuild.exe was not found! Please install Visual Studio or Build Tools."
    Exit 1
}

Write-Host "Using MSBuild from: $msbuild" -ForegroundColor Gray

# 3. Build Solution
Write-Host "Building Visual Studio VSTO solution..." -ForegroundColor Yellow
$slnPath = Join-Path $scriptDir "AITranslateAddin.sln"

& $msbuild $slnPath /t:Rebuild /p:Configuration=Release /p:SignManifests=false

if ($LASTEXITCODE -eq 0) {
    # 4. Copy build artifacts to the common publish directory
    $publishDir = Join-Path $scriptDir "publish"
    if (Test-Path $publishDir) {
        Remove-Item $publishDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $publishDir | Out-Null
    
    foreach ($project in $projects) {
        $binDir = Join-Path (Join-Path $scriptDir $project) "bin\Release"
        if (Test-Path $binDir) {
            $projPubDir = Join-Path $publishDir $project
            New-Item -ItemType Directory -Path $projPubDir | Out-Null
            
            # Copy DLL, DLL.manifest, VSTO file
            Copy-Item "$binDir\$project.dll" -Destination $projPubDir -Force
            Copy-Item "$binDir\$project.dll.manifest" -Destination $projPubDir -Force
            Copy-Item "$binDir\$project.vsto" -Destination $projPubDir -Force
            
            # Copy dependency DLLs
            Get-ChildItem $binDir -Filter "*.dll" | Where-Object { $_.Name -ne "$project.dll" } | ForEach-Object {
                Copy-Item $_.FullName -Destination $projPubDir -Force
            }
        }
    }
    
    # Copy certificate to publish directory
    Copy-Item $pfxFile -Destination $publishDir -Force

    # 5. Compile Setup.cs into single-file setup.exe
    Write-Host "Compiling native installer (setup.exe) for VSTO..." -ForegroundColor Yellow
    $csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    if (-not (Test-Path $csc)) {
        $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
    }

    if (Test-Path $csc) {
        $setupSource = Join-Path $scriptDir "Setup.cs"
        $outputExe = Join-Path $scriptDir "setup.exe"
        $iconIco = Join-Path $scriptDir "icon.ico"
        
        # Copy icon from com-version if available
        $sourceIconIco = Join-Path $scriptDir "../com-version/scripts/icon.ico"
        if (Test-Path $sourceIconIco) {
            Copy-Item $sourceIconIco -Destination $iconIco -Force
        }
        $sourceIconPng = Join-Path $scriptDir "../icon.png"
        $destIconPng = Join-Path $scriptDir "icon.png"
        if (Test-Path $sourceIconPng) {
            Copy-Item $sourceIconPng -Destination $destIconPng -Force
        }

        # Build resources arguments
        $resourceArgs = @(
            "/resource:`"$pfxFile`",AITranslate.pfx",
            "/resource:`"$(Join-Path $publishDir "AITranslateExcel\AITranslateExcel.dll")`",AITranslateExcel.dll",
            "/resource:`"$(Join-Path $publishDir "AITranslateExcel\AITranslateExcel.dll.manifest")`",AITranslateExcel.dll.manifest",
            "/resource:`"$(Join-Path $publishDir "AITranslateExcel\AITranslateExcel.vsto")`",AITranslateExcel.vsto",
            "/resource:`"$(Join-Path $publishDir "AITranslateWord\AITranslateWord.dll")`",AITranslateWord.dll",
            "/resource:`"$(Join-Path $publishDir "AITranslateWord\AITranslateWord.dll.manifest")`",AITranslateWord.dll.manifest",
            "/resource:`"$(Join-Path $publishDir "AITranslateWord\AITranslateWord.vsto")`",AITranslateWord.vsto",
            "/resource:`"$(Join-Path $publishDir "AITranslatePPT\AITranslatePPT.dll")`",AITranslatePPT.dll",
            "/resource:`"$(Join-Path $publishDir "AITranslatePPT\AITranslatePPT.dll.manifest")`",AITranslatePPT.dll.manifest",
            "/resource:`"$(Join-Path $publishDir "AITranslatePPT\AITranslatePPT.vsto")`",AITranslatePPT.vsto"
        )

        $iconArg = ""
        if (Test-Path $iconIco) {
            $iconArg = "/win32icon:`"$iconIco`""
            $resourceArgs += "/resource:`"$iconIco`",icon.ico"
        }
        if (Test-Path $destIconPng) {
            $resourceArgs += "/resource:`"$destIconPng`",icon.png"
        }

        # Run csc
        $cscArgs = @("/target:winexe", "/out:`"$outputExe`"")
        if ($iconArg -ne "") { $cscArgs += $iconArg }
        $cscArgs += $resourceArgs
        $cscArgs += "`"$setupSource`""

        # Execute compilation
        Write-Host "Compiling setup.exe using: $csc" -ForegroundColor Gray
        & $csc $cscArgs
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "-> Successfully compiled single-file installer: $outputExe" -ForegroundColor Green
        } else {
            Write-Error "Failed to compile setup.exe installer"
            Exit 1
        }
    } else {
        Write-Warning "csc.exe compiler not found, skipping setup.exe installer creation."
    }

    Write-Host "========================================================" -ForegroundColor Green
    Write-Host "BUILD AND PUBLISH SUCCESSFUL!" -ForegroundColor Green
    Write-Host "Artifacts are stored at: $publishDir" -ForegroundColor Green
    Write-Host "Single-file installer is ready at: $outputExe" -ForegroundColor Green
    Write-Host "========================================================" -ForegroundColor Green
} else {
    Write-Error "Build failed!"
    Exit 1
}
