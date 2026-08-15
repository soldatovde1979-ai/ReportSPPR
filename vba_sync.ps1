param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("export", "import")]
    [string]$Action,
    [string]$WorkbookPath = "$PSScriptRoot\Workbook.xlsm",
    [string]$SourceDir = "$PSScriptRoot\src"
)
if (-not (Test-Path $SourceDir)) { New-Item -ItemType Directory -Path $SourceDir | Out-Null }
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    if (-not (Test-Path $WorkbookPath)) {
        Write-Host "Создание новой книги $WorkbookPath..." -ForegroundColor Yellow
        $wb = $excel.Workbooks.Add()
        $wb.SaveAs($WorkbookPath, 52)
    } else {
        $wb = $excel.Workbooks.Open($WorkbookPath)
    }
    $vbProj = $wb.VBProject
    if ($Action -eq "export") {
        Write-Host "📥 Экспорт модулей из Excel в $SourceDir..." -ForegroundColor Cyan
        foreach ($comp in $vbProj.VBComponents) {
            $ext = switch ($comp.Type) { 1 { ".bas" } 2 { ".cls" } 3 { ".frm" } default { $null } }
            if ($ext) {
                $targetFile = Join-Path $SourceDir ($comp.Name + $ext)
                $comp.Export($targetFile)
                Write-Host "  -> Экспортирован: $($comp.Name)$ext"
            }
        }
        Write-Host "✅ Экспорт завершен!" -ForegroundColor Green
    }
    elseif ($Action -eq "import") {
        Write-Host "📤 Импорт кода из $SourceDir в Excel..." -ForegroundColor Cyan
        $files = Get-ChildItem -Path $SourceDir -Include *.bas, *.cls, *.frm -File
        foreach ($file in $files) {
            $modName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $existing = $null
            try { $existing = $vbProj.VBComponents.Item($modName) } catch {}
            if ($existing -and $existing.Type -in @(1, 2, 3)) {
                $vbProj.VBComponents.Remove($existing)
            }
            $vbProj.VBComponents.Import($file.FullName) | Out-Null
            Write-Host "  <- Импортирован: $($file.Name)"
        }
        $wb.Save()
        Write-Host "✅ Импорт завершен, книга сохранена!" -ForegroundColor Green
    }
}
catch {
    Write-Error "Ошибка синхронизации: $_"
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}