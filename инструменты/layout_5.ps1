Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
$wbs = $xl.Workbooks
$files = Get-ChildItem "E:\ИИТ\ЭлектронныйБизнес\СДАЧА" -Recurse -Filter *.xlsx | Where-Object { $_.Name -match "^(analiz_|klassifikaciya|6_2_|6_3_)" }
$report = @()
foreach ($f in $files) {
    try {
        $wb = $wbs.Open($f.FullName)
        $done = 0; $skipped = 0
        foreach ($ws in $wb.Worksheets) {
            $ur = $ws.UsedRange
            $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
            if ($nr -lt 1 -or $nc -lt 1) { continue }
            $hasChart = ($ws.ChartObjects().Count -gt 0)
            $ur.WrapText = $true
            $ur.VerticalAlignment = -4160
            if ($hasChart) { $skipped++; continue }
            if ($nr -eq 1 -and $nc -eq 1) { continue }
            $vals = $ur.Value2
            $c0 = $ur.Column; $r0 = $ur.Row
            for ($c = 1; $c -le $nc; $c++) {
                $max = 0
                $lim = [Math]::Min($nr, 400)
                for ($r = 1; $r -le $lim; $r++) {
                    $v = $vals[$r, $c]
                    if ($v -is [string]) {
                        $len = $v.Length
                        $nl = $v.IndexOf("`n")
                        if ($nl -ge 0) { $len = [Math]::Max($nl, 10) }
                        if ($len -gt $max) { $max = $len }
                    } elseif ($v -ne $null) {
                        $len = ([string]$v).Length
                        if ($len -gt $max) { $max = $len }
                    }
                }
                $col = $ws.Columns.Item($c0 + $c - 1)
                $cur = $col.ColumnWidth
                $target = [Math]::Min(55, [Math]::Max(9, [Math]::Ceiling($max * 1.05)))
                if ($cur -lt $target) { $col.ColumnWidth = $target }
                elseif ($cur -gt 60) { $col.ColumnWidth = 60 }
            }
            $ur.Rows.AutoFit() | Out-Null
            $merged = $ur.MergeCells
            if ($merged -ne $false) {
                for ($r = 1; $r -le $nr; $r++) {
                    for ($c = 1; $c -le $nc; $c++) {
                        $v = $vals[$r, $c]
                        if ($v -is [string] -and $v.Length -gt 40) {
                            $cell = $ws.Cells.Item($r0 + $r - 1, $c0 + $c - 1)
                            if ($cell.MergeCells) {
                                $area = $cell.MergeArea
                                $w = 0
                                foreach ($cc in $area.Columns) { $w += $cc.ColumnWidth }
                                $rowsInArea = $area.Rows.Count
                                $lines = [Math]::Ceiling(($v.Length * 1.1) / [Math]::Max($w, 8))
                                $need = [Math]::Min(409, $lines * 15)
                                $have = 0
                                foreach ($rr in $area.Rows) { $have += $rr.RowHeight }
                                if ($need -gt $have) { $area.Rows.Item(1).RowHeight = [Math]::Min(409, $area.Rows.Item(1).RowHeight + ($need - $have)) }
                            }
                        }
                    }
                }
            }
            $done++
        }
        $xl.CalculateFull()
        $wb.Save(); $wb.Close($false)
        $report += [pscustomobject]@{ File = $f.Name; Status = "OK, листов настроено: $done, пропущено (диаграммы): $skipped" }
    } catch {
        $report += [pscustomobject]@{ File = $f.Name; Status = "ОШИБКА: " + $_.Exception.Message }
        try { $wb.Close($false) } catch {}
    }
}
$xl.Quit()
$report | Format-Table -AutoSize

