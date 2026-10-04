$t = "C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
Start-Sleep 3
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
New-Item -ItemType Directory -Force "$t\prefinal" | Out-Null
$files = Get-ChildItem "E:\ИИТ\ЭлектронныйБизнес\СДАЧА" -Recurse -Filter *.xlsx | Where-Object { $_.Name -match "^(analiz_|klassifikaciya|6_2_|6_3_)" }
foreach ($f in $files) { Copy-Item -LiteralPath $f.FullName -Destination "$t\prefinal\$($f.Name)" -Force }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wbs = $xl.Workbooks
$report = @()

function CountErrors($wb) {
    $n = 0
    foreach ($ws in $wb.Worksheets) {
        try { $r = $ws.UsedRange.SpecialCells(-4123, 16); $n += $r.Count } catch {}
    }
    return $n
}
function ChartState($wb) {
    $s = 0; $bad = 0
    foreach ($ws in $wb.Worksheets) {
        foreach ($co in $ws.ChartObjects()) {
            foreach ($ser in $co.Chart.SeriesCollection()) {
                $s++
                try { if ($ser.Formula -match "#REF") { $bad++ } } catch {}
            }
        }
    }
    return "$s/$bad"
}
function IsFiller($v) {
    if ($v -eq $null) { return $true }
    if ($v -is [double]) { return ($v -eq 0) }
    $s = [string]$v
    $s = $s.Trim()
    if ($s -eq "") { return $true }
    if ($s -match '^(Не заполнено|нет данных|Заполнить.*|[-–]|0|0,0+|0%|FALSE|ЛОЖЬ)$') { return $true }
    return $false
}
function IsId($v) {
    if ($v -eq $null) { return $true }
    if ($v -is [double]) { return $true }
    $s = ([string]$v).Trim()
    if ($s -eq "") { return $true }
    return ($s -match '^[A-Za-zА-Яа-я]{0,4}[-_]?\d{1,4}$')
}

foreach ($f in $files) {
    $before = ""; $log = @()
    try {
        $wb = $wbs.Open($f.FullName)
        $e0 = CountErrors $wb
        $c0 = ChartState $wb
        $total = 0; $skip=@(); foreach($w2 in $wb.Worksheets){ foreach($co in $w2.ChartObjects()){ $skip += $w2.Name; foreach($se in $co.Chart.SeriesCollection()){ foreach($m in [regex]::Matches($se.Formula,"'?([^'!,()]+)'?!")){ $skip += $m.Groups[1].Value } } } }
        foreach ($ws in $wb.Worksheets) {
            if ($skip -contains $ws.Name) { continue }
            $ur = $ws.UsedRange
            $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
            if ($nr -lt 6) { continue }
            $r0 = $ur.Row; $c0col = $ur.Column
            $vals = $ur.Value2
            $last = 0
            for ($r = 1; $r -le $nr; $r++) {
                $filled = $false
                for ($c = 1; $c -le $nc; $c++) {
                    $v = $vals[$r, $c]
                    if ($c -eq 1) { if (-not (IsId $v)) { $filled = $true; break } else { if (-not (IsFiller $v) -and $v -isnot [double]) { } ; continue } }
                    if (-not (IsFiller $v)) { $filled = $true; break }
                }
                if ($filled) { $last = $r }
            }
            if ($last -lt 1) { continue }
            $delN = $nr - $last
            if ($delN -ge 3) {
                $startRow = $r0 + $last
                $endRow = $r0 + $nr - 1
                try {
                    $ws.Rows("$startRow`:$endRow").Delete() | Out-Null
                    $total += $delN
                    $log += "$($ws.Name): -$delN"
                } catch { $log += "$($ws.Name): не удалось ($($_.Exception.Message))" }
            }
        }
        $xl.CalculateFull()
        $e1 = CountErrors $wb
        $c1 = ChartState $wb
        if ($e1 -gt $e0 -or $c1 -ne $c0) {
            $wb.Close($false)
            $report += [pscustomobject]@{ File = $f.Name; Status = "ОТКАТ (ошибок $e0 -> $e1, диаграммы $c0 -> $c1)"; Detail = ($log -join "; ") }
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


