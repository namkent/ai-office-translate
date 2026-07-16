# InspectExcelFactory.ps1
try {
    $dllPath = "p:\web\translate-plugin\vsto-version\lib\Microsoft.Office.Tools.Excel.v4.0.Utilities.dll"
    if (Test-Path $dllPath) {
        $assembly = [System.Reflection.Assembly]::LoadFile($dllPath)
        $type = $assembly.GetType("Microsoft.Office.Tools.Excel.ApplicationFactory")
        Write-Host "--- Methods in Microsoft.Office.Tools.Excel.ApplicationFactory ---"
        $type.GetMethods([System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Instance) | ForEach-Object {
            $params = $_.GetParameters() | ForEach-Object { "$($_.ParameterType.FullName) $($_.Name)" }
            Write-Host "Method: $($_.ReturnType.FullName) $($_.Name) ($($params -join ', '))"
        }
    }
}
catch {
    Write-Error $_.Exception.Message
}
