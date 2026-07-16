# ExtractVSTODlls.ps1
# Run this script AFTER installing the "Microsoft 365 development" workload.
# It copies all required VSTO reference DLLs into a local 'lib' folder,
# allowing the project to be built on other computers without installing the workload.

$scriptDir = $PSScriptRoot
$libDir = Join-Path $scriptDir "lib"

if (-not (Test-Path $libDir)) {
    New-Item -ItemType Directory -Path $libDir | Out-Null
}

# Standard search paths for VSTO reference assemblies
$searchPaths = @(
    "C:\Program Files (x86)\Reference Assemblies\Microsoft\VSTO40\v4.0.Framework",
    "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0",
    "C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0",
    "C:\Windows\Microsoft.NET\assembly\GAC_MSIL"
)

# DLLs we want to extract
$dllNames = @(
    "Microsoft.Office.Tools.dll",
    "Microsoft.Office.Tools.Common.dll",
    "Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
    "Microsoft.Office.Tools.Excel.dll",
    "Microsoft.Office.Tools.Excel.v4.0.Utilities.dll",
    "Microsoft.Office.Tools.Word.dll",
    "Microsoft.Office.Tools.Word.v4.0.Utilities.dll",
    "Microsoft.Office.Tools.v4.0.Framework.dll",
    "Microsoft.VisualStudio.Tools.Applications.Runtime.dll"
)

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "EXTRACTING VSTO REFERENCE DLLS TO LOCAL 'lib' FOLDER..." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

$copiedCount = 0

foreach ($dll in $dllNames) {
    $found = $false
    foreach ($path in $searchPaths) {
        if (-not (Test-Path $path)) { continue }
        
        # Search recursively in GAC or directly in Reference folders
        $file = Get-ChildItem -Path $path -Filter $dll -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        
        if ($file -and (Test-Path $file.FullName)) {
            $dest = Join-Path $libDir $dll
            Copy-Item $file.FullName -Destination $dest -Force
            Write-Host "-> Extracted: $dll" -ForegroundColor Green
            Write-Host "   From: $($file.FullName)" -ForegroundColor Gray
            $found = $true
            $copiedCount++
            break
        }
    }
    
    if (-not $found) {
        Write-Warning "Could not find reference assembly: $dll"
    }
}

Write-Host "--------------------------------------------------------" -ForegroundColor Cyan
if ($copiedCount -eq $dllNames.Count) {
    Write-Host "SUCCESS: All $copiedCount DLLs extracted to '$libDir'!" -ForegroundColor Green
    Write-Host "Once committed, you can build this project on any machine." -ForegroundColor Green
} else {
    Write-Warning "Only $copiedCount of $($dllNames.Count) DLLs were extracted. Please make sure the workload is fully installed."
}
Write-Host "========================================================" -ForegroundColor Cyan
