Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wbs = $xl.Workbooks
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
$root = "E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
# 1. klassifikaciya 07_Карта: shrink table, clear tail
$f = Get-ChildItem $root -Recurse -Filter "klassifikaciya*.xlsx" | Select-Object -First 1
$wb = $wbs.Open($f.FullName)
$ws = $wb.Worksheets.Item("07_Карта")
foreach ($lo in $ws.ListObjects) {
    $rng = $lo.Range
    $last = 25
    $endRow = $rng.Row + $rng.Rows.Count - 1
    if ($endRow -gt $last) {
        $lo.Resize($ws.Range($ws.Cells.Item($rng.Row, $rng.Column), $ws.Cells.Item($last, $rng.Column + $rng.Columns.Count - 1)))
        $tail = $ws.Range($ws.Cells.Item($last + 1, $rng.Column), $ws.Cells.Item($endRow, $rng.Column + $rng.Columns.Count - 1))
        $tail.Clear() | Out-Null
    }
}
$xl.CalculateFull(); $wb.Save(); $wb.Close($false)
"klassifikaciya 07_Карта: таблица сужена до строки 25"
# 2. hide empty template blocks in the three files that rolled back on delete
foreach ($pat in "ПК_*", "ПЛ_*", "СВ_*") {
    $f = Get-ChildItem "$root\ЛР2-3" -Filter "$pat.xlsx" | Select-Object -First 1
    $wb = $wbs.Open($f.FullName); $hidden = 0
    foreach ($ws in $wb.Worksheets) {
        if ($ws.ChartObjects().Count -gt 0) { continue }
        $ur = $ws.UsedRange; $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
        if ($nr -lt 6 -or $nc -lt 2) { continue }
        $r0 = $ur.Row; $vals = $ur.Value2
        $empty = New-Object bool[] ($nr + 1)
        for ($r = 1; $r -le $nr; $r++) {
            $ok = (IsIdV $vals[$r, 1])
            if ($ok) { for ($c = 2; $c -le $nc; $c++) { if (-not (IsFillerV $vals[$r, $c])) { $ok = $false; break } } }
            $empty[$r] = $ok
        }
        $r = 1
        while ($r -le $nr) {
            if ($empty[$r]) {
                $s = $r
                while ($r -le $nr -and $empty[$r]) { $r++ }
                if (($r - $s) -ge 3) {
                    $a = $r0 + $s - 1; $b = $r0 + $r - 2
                    $ws.Rows("$a`:$b").Hidden = $true; $hidden += ($b - $a + 1)
                }
            } else { $r++ }
        }
    }
    $wb.Save(); $wb.Close($false)
    "$($f.Name): скрыто строк $hidden"
}
$xl.Quit()
