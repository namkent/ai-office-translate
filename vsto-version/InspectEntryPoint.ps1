# InspectEntryPoint.ps1
try {
    $dllPath = "p:\web\translate-plugin\vsto-version\lib\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
    if (Test-Path $dllPath) {
        $assembly = [System.Reflection.Assembly]::LoadFile($dllPath)
        $type = $assembly.GetType("Microsoft.Office.Tools.EntryPoint")
        Write-Host "Type: $($type.FullName)"
        Write-Host "Base: $($type.BaseType.FullName)"
        Write-Host "Is Interface: $($type.IsInterface)"
        Write-Host "Is Attribute: $($type.IsSubclassOf([System.Attribute]))"
        $type.GetProperties() | ForEach-Object {
            Write-Host "Property: $($_.Name) ($($_.PropertyType.FullName))"
        }
        $type.GetMethods() | ForEach-Object {
            Write-Host "Method: $($_.Name) ($($_.ReturnType.FullName))"
        }
    }
}
catch {
    Write-Error $_.Exception.Message
}
