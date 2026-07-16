$desktop = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "*.docx")
$files = Get-ChildItem -Path $desktop
$f = $null
foreach ($file in $files) {
    if ($file.Name.StartsWith("~$")) { continue }
    $f = $file
    break
}

if ($f -eq $null) {
    Write-Host "No docx file found!"
    exit 1
}

# Kill any existing Word instances to prevent lock issues
taskkill /F /IM winword.exe 2>$null | Out-Null
Start-Sleep -Milliseconds 500

$tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "temp_copy_test.docx")
Copy-Item -Path $f.FullName -Destination $tempFile -Force

$word = New-Object -ComObject Word.Application
$word.DisplayAlerts = 0
$word.Visible = $false

try {
    $doc = $word.Documents.Open($tempFile, $false, $true)
    Write-Host "Source document opened. Shapes count: $($doc.Shapes.Count)"
    
    if ($doc.Shapes.Count -gt 0) {
        $backupDoc = $word.Documents.Add()
        Write-Host "Backup document created."
        
        $s = $doc.Shapes.Item(1)
        Write-Host "Attempting to select and copy shape: $($s.Name), Type=$($s.Type)"
        
        $s.Select()
        $word.Selection.Copy()
        Write-Host "Shape copied via Selection."
        
        $shapesBefore = $backupDoc.Shapes.Count
        $inlineBefore = $backupDoc.InlineShapes.Count
        
        $backupDoc.Content.InsertParagraphAfter()
        $pasteRange = $backupDoc.Paragraphs.Item($backupDoc.Paragraphs.Count).Range
        $pasteRange.Paste()
        
        Write-Host "Shape pasted into backupDoc."
        Write-Host "  backupDoc.Shapes count: $($backupDoc.Shapes.Count) (Before: $shapesBefore)"
        Write-Host "  backupDoc.InlineShapes count: $($backupDoc.InlineShapes.Count) (Before: $inlineBefore)"
        
        $backupDoc.Close($false)
    }
    
    $doc.Close($false)
} catch {
    Write-Host "Error occurred: $_"
} finally {
    $word.Quit()
    if (Test-Path $tempFile) { Remove-Item -Path $tempFile -Force }
}
