$excelPath = Join-Path $PSScriptRoot 'Ops files\Monthly Review Meeting.xlsx'
$extractScript = Join-Path $PSScriptRoot '_extract_data.ps1'
$folder = Split-Path $excelPath
$fileName = Split-Path $excelPath -Leaf

Write-Host "HALO Dashboard - Watching for changes to:" -ForegroundColor Cyan
Write-Host "  $excelPath" -ForegroundColor Yellow
Write-Host "Press Ctrl+C to stop." -ForegroundColor Gray
Write-Host ""

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $folder
$watcher.Filter = $fileName
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $false

$lastRun = [datetime]::MinValue

while ($true) {
    $result = $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::Changed, 5000)
    if ($result.TimedOut) { continue }

    # Debounce: Excel saves trigger multiple events, wait for file to settle
    $now = Get-Date
    if (($now - $lastRun).TotalSeconds -lt 30) { continue }

    Write-Host "[$($now.ToString('HH:mm:ss'))] File changed! Waiting for Excel to finish saving..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5

    Write-Host "[$($now.ToString('HH:mm:ss'))] Extracting latest data..." -ForegroundColor Cyan
    try {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $extractScript
        $lastRun = Get-Date
        Write-Host "[$($lastRun.ToString('HH:mm:ss'))] Done! Refresh the dashboard in your browser (F5)." -ForegroundColor Green
    } catch {
        Write-Host "[$($now.ToString('HH:mm:ss'))] Extraction failed: $_" -ForegroundColor Red
    }
    Write-Host ""
}
