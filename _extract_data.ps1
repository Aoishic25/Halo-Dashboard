$ErrorActionPreference = 'Stop'
$excelPath = Join-Path $PSScriptRoot 'Ops files\Monthly Review Meeting.xlsx'
$outputPath = Join-Path $PSScriptRoot '_data_block.js'

if (-not (Test-Path $excelPath)) {
    Write-Host "ERROR: Excel file not found at $excelPath" -ForegroundColor Red
    exit 1
}

Write-Host "Opening Excel..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($excelPath, $false, $true)

    function Export-Sheet($name, $maxRows, $headerRow, $startCol, $endCol) {
        if (-not $headerRow) { $headerRow = 1 }
        if (-not $startCol) { $startCol = 1 }
        $ws = $null
        foreach ($s in $wb.Sheets) { if ($s.Name -eq $name) { $ws = $s; break } }
        if (-not $ws) { return '[]' }
        $used = $ws.UsedRange
        $rows = [Math]::Min($used.Rows.Count, $maxRows + $headerRow)
        $cols = if ($endCol) { $endCol } else { $used.Columns.Count }
        $headers = @()
        for ($c = $startCol; $c -le $cols; $c++) {
            $h = $used.Cells.Item($headerRow, $c).Text.Trim()
            if ($h -eq '') { $h = "Col$c" }
            $headers += $h
        }
        $arr = @()
        for ($r = ($headerRow + 1); $r -le $rows; $r++) {
            $obj = [ordered]@{}
            $hasData = $false
            $ci = 0
            for ($c = $startCol; $c -le $cols; $c++) {
                $raw = $used.Cells.Item($r, $c).Value2
                if ($null -eq $raw) { $v = '' } else { $v = [string]$raw }
                $obj[$headers[$ci]] = $v
                if ($v -ne '') { $hasData = $true }
                $ci++
            }
            if ($hasData) { $arr += New-Object PSObject -Property $obj }
        }
        if ($arr.Count -eq 0) { return '[]' }
        if ($arr.Count -eq 1) { return '[' + ($arr | ConvertTo-Json -Compress -Depth 3) + ']' }
        return ($arr | ConvertTo-Json -Compress -Depth 3)
    }

    $allSheets = [ordered]@{
        'stockFlow'          = @('19. Stock Flow', 200)
        'costingDet'         = @('14. Costing Details', 1200, 1, 1, 11)
        'costingCases'       = @('14. Costing Details', 1200, 1, 13, 25)
        'waitingCharges'     = @('8. Waiting Charges', 1100, 1, 16, 32)
        'weeklyOrders'       = @('5. Week-wise Orders Details', 200)
        'damages'            = @('4. % of Damages', 200, 3)
        'qualityIssues'      = @('3. Quality Issues', 1000)
        'manualOrders'       = @('6. Manual orders', 1100)
        'palletAging'        = @('13. Mov. Ageing', 1200)
        'copackingWeekly'    = @('29. COPACKING ORDERS-WEEK WISE', 200)
        'whDamages'          = @('12. WH Handling Damages', 200)
        'inboundFlow'        = @('16. Inbound Flow-25', 300, 1, 1, 4)
        'invoiceSummary'     = @('24. WH Invoice Summary', 300)
        'tempReport'         = @('20. Temp report', 1100)
        'subjects'           = @('1. Subjects', 300)
        'hubInbound'         = @('15. Hubwise Inbound Summary', 1100)
        'hubOutbound'        = @('18. Hubwise Outbound Summary', 700)
        'storage'            = @('10. Storage', 1100, 1, 16, 45)
        'storageHub'         = @('10. Storage', 50, 1, 1, 14)
        'copackingOrders'    = @('28. Co-Packing Orders', 700, 1, 1, 4)
        'coPacking'          = @('26. Co-Packing', 1400)
        'customerComplaints' = @('27. Customers Complaints', 1100)
        'inboundIssues'      = @('25. Inbound Issues', 300)
        'inbGrnMonthly'      = @('2. Inbound Plan vs GRN', 1100, 1, 8, 18)
        'expiredStock'       = @('11. Expiry & Near Expire', 1400)
        'freights'           = @('7. Freights', 1100)
        'outboundCases'      = @('22. Outbound Summary', 1100)
        'truckTurnover'      = @('9. Truck Turnover Time', 3600)
        'inboundSummary'     = @('17. Inbound Summary', 300)
        'inboundDetail'      = @('21. Inbound Detail', 2000)
        'ytdTrends'          = @('23. YTD Trends', 400)
        'fzeSkuGw'           = @('30. FZE SKU''s GW Details', 100)
    }

    $result = 'var EMBEDDED_DATA={'
    $first = $true
    $total = $allSheets.Count
    $i = 0

    foreach ($key in $allSheets.Keys) {
        $i++
        $info = $allSheets[$key]
        Write-Host "  [$i/$total] Extracting: $key ($($info[0]))..." -ForegroundColor Yellow
        $json = Export-Sheet $info[0] $info[1] $info[2] $info[3] $info[4]
        if (-not $json -or $json -eq '') { $json = '[]' }
        if (-not $first) { $result += ',' }
        $result += """$key"":$json"
        $first = $false
    }

    $result += '};'

    [System.IO.File]::WriteAllText($outputPath, $result, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "SUCCESS: Data extracted to _data_block.js ($([Math]::Round($result.Length / 1KB)) KB)" -ForegroundColor Green

    $wb.Close($false)
}
finally {
    try { $excel.Quit() } catch {}
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
}

exit 0
