$t = "C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
New-Item -ItemType Directory -Force "$t\prefinal3" | Out-Null
$files = Get-ChildItem "E:\ИИТ\ЭлектронныйБизнес\СДАЧА" -Recurse -Filter *.xlsx | Where-Object { -not $_.Name.StartsWith("~") -and -not $_.Name.StartsWith("XLT") -and $_.Name -notlike "SW_рабочая*" }
foreach ($f in $files) { Copy-Item -LiteralPath $f.FullName -Destination "$t\prefinal3\$($f.Name)" -Force }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wbs = $xl.Workbooks
$report = @()

function CountErrors($wb) {
    $n = 0
    foreach ($ws in $wb.Worksheets) { try { $r = $ws.UsedRange.SpecialCells(-4123, 16); $n += $r.Count } catch {} }
    return $n
}
function IsFillerV($v) {
    if ($v -eq $null) { return $true }
    if ($v -is [double]) { return ($v -eq 0) }
    $s = ([string]$v).Trim()
    if ($s -eq "") { return $true }
    return ($s -match '^(Не заполнено|Не начато|Нет|нет данных|Заполнить.*|[-–]|0|0,0+|0%|FALSE|ЛОЖЬ)$')
}
function IsIdV($v) {
    if ($v -eq $null) { return $true }
    if ($v -is [double]) { return $true }
    $s = ([string]$v).Trim()
    if ($s -eq "") { return $true }
    return ($s -match '^[A-Za-zА-Яа-я]{0,4}[-_]?\d{1,4}$')
}

foreach ($f in $files) {
    $log = @()
    try {
        $wb = $wbs.Open($f.FullName)
        $e0 = CountErrors $wb
        $total = 0
        foreach ($ws in $wb.Worksheets) {
            if ($ws.ChartObjects().Count -gt 0) { continue }
            $ur = $ws.UsedRange
            $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
            if ($nr -lt 6 -or $nc -lt 2) { continue }
            $r0 = $ur.Row
            $vals = $ur.Value2
            $empty = New-Object bool[] ($nr + 1)
            for ($r = 1; $r -le $nr; $r++) {
                $ok = $true
                if (-not (IsIdV $vals[$r, 1])) { $ok = $false }
                if ($ok) {
                    $anyId = $false
                    $v1 = $vals[$r, 1]
                    if ($v1 -ne $null -and ([string]$v1).Trim() -ne "") { $anyId = $true }
                    for ($c = 2; $c -le $nc; $c++) { if (-not (IsFillerV $vals[$r, $c])) { $ok = $false; break } }
                    if ($ok -and -not $anyId) { $ok = $true }
                }
                $empty[$r] = $ok
            }
            $blocks = @()
            $r = 1
            while ($r -le $nr) {
                if ($empty[$r]) {
                    $s = $r
                    while ($r -le $nr -and $empty[$r]) { $r++ }
                    $len = $r - $s
                    if ($len -ge 3) { $blocks += ,@($s, ($r - 1)) }
                } else { $r++ }
            }
            for ($i = $blocks.Count - 1; $i -ge 0; $i--) {
                $a = $r0 + $blocks[$i][0] - 1; $b = $r0 + $blocks[$i][1] - 1
                try { $ws.Rows("$a`:$b").Delete() | Out-Null; $total += ($b - $a + 1) } catch {}
            }
            if ($blocks.Count -gt 0) { $log += "$($ws.Name): -$(($blocks | ForEach-Object { $_[1] - $_[0] + 1 } | Measure-Object -Sum).Sum)" }
        }
        $xl.CalculateFull()
        $e1 = CountErrors $wb
        if ($e1 -gt $e0) {
            $wb.Close($false)
            $report += [pscustomobject]@{ File = $f.Name; Status = "ОТКАТ (ошибок $e0 -> $e1)"; Detail = ($log -join "; ") }
        } else {
            $wb.Save(); $wb.Close($false)
            $report += [pscustomobject]@{ File = $f.Name; Status = "OK, удалено строк: $total"; Detail = ($log -join "; ") }
        }
    } catch {
        $report += [pscustomobject]@{ File = $f.Name; Status = "ОШИБКА: " + $_.Exception.Message; Detail = "" }
        try { $wb.Close($false) } catch {}
    }
}
$xl.Quit()
$report | Format-Table -AutoSize -Wrap
