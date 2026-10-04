Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$dash = [string][char]0x2013
$files = Get-ChildItem "E:\ИИТ\ЭлектронныйБизнес\СДАЧА" -Recurse -Filter *.xlsx | Where-Object { -not $_.Name.StartsWith("~") -and -not $_.Name.StartsWith("XLT") -and $_.Name -notlike "SW_рабочая*" }
$report = @()

function Snapshot($wb) {
    $h = @{}
    foreach ($ws in $wb.Worksheets) {
        $ur = $ws.UsedRange
        if ($ur.Cells.Count -gt 1) { $h[$ws.Name] = @($ur.Value2, $ur.Row, $ur.Column) }
    }
    return $h
}
function CountDiff($before, $after, $written) {
    $diff = 0
    foreach ($k in $before.Keys) {
        if (-not $after.ContainsKey($k)) { continue }
        $b = $before[$k][0]; $a = $after[$k][0]; $r0 = $before[$k][1]; $c0 = $before[$k][2]
        $nr = $b.GetLength(0); $nc = $b.GetLength(1)
        if ($a.GetLength(0) -ne $nr -or $a.GetLength(1) -ne $nc) { $diff++; continue }
        for ($r = 1; $r -le $nr; $r++) {
            for ($c = 1; $c -le $nc; $c++) {
                if ($written.ContainsKey("$k|$($r0 + $r - 1)|$($c0 + $c - 1)")) { continue }
                if ([string]$b[$r, $c] -ne [string]$a[$r, $c]) { $diff++ }
            }
        }
    }
    return $diff
}

foreach ($f in $files) {
    try {
        $wb = $xl.Workbooks.Open($f.FullName)
        $xl.CalculateFull()
        $filled = 0; $reverted = @()
        foreach ($ws in $wb.Worksheets) {
            if ($ws.ListObjects.Count -eq 0) { continue }
            $before = Snapshot $wb
            $written = @{}
            $cells = @()
            foreach ($lo in $ws.ListObjects) {
                $body = $lo.DataBodyRange
                if ($body -eq $null) { continue }
                $r0 = $body.Row; $c0 = $body.Column
                $nr = $body.Rows.Count; $nc = $body.Columns.Count
                $fm = $body.Formula
                for ($r = 1; $r -le $nr; $r++) {
                    $rowObj = $ws.Rows.Item($r0 + $r - 1)
                    if ($rowObj.Hidden) { continue }
                    for ($c = 1; $c -le $nc; $c++) {
                        $v = $fm[$r, $c]
                        if ($v -eq $null -or ([string]$v).Trim() -eq "") {
                            $cell = $ws.Cells.Item($r0 + $r - 1, $c0 + $c - 1)
                            if ($cell.MergeCells) { continue }
                            $cell.Value2 = $dash
                            $written["$($ws.Name)|$($r0 + $r - 1)|$($c0 + $c - 1)"] = 1
                            $cells += $cell
                        }
                    }
                }
            }
            if ($written.Count -eq 0) { continue }
            $xl.CalculateFull()
            $after = Snapshot $wb
            $d = CountDiff $before $after $written
            if ($d -gt 0) {
                foreach ($cell in $cells) { $cell.ClearContents() | Out-Null }
                $xl.CalculateFull()
                $reverted += "$($ws.Name) (изменилось бы $d расчетных ячеек)"
            } else { $filled += $written.Count }
        }
        $wb.Save(); $wb.Close($false)
        $report += [pscustomobject]@{ File = $f.Name; Filled = $filled; Skipped = ($reverted -join "; ") }
    } catch {
        $report += [pscustomobject]@{ File = $f.Name; Filled = "ОШИБКА: " + $_.Exception.Message; Skipped = "" }
        try { $wb.Close($false) } catch {}
    }
}
$xl.Quit()
$report | Format-Table -AutoSize -Wrap
