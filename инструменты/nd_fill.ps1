Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$f = Get-ChildItem "E:\ИИТ\ЭлектронныйБизнес\СДАЧА\ЛР4" -Filter "ocenka*.xlsx" | Select-Object -First 1
$wb = $xl.Workbooks.Open($f.FullName)
$nd = "н/д"
function Snapshot($wb) {
    $h = @{}
    foreach ($ws in $wb.Worksheets) { $ur = $ws.UsedRange; if ($ur.Cells.Count -gt 1) { $h[$ws.Name] = @($ur.Value2, $ur.Row, $ur.Column) } }
    return $h
}
$xl.CalculateFull()
$before = Snapshot $wb
$written = @{}
foreach ($sn in "02_Конкуренты_SW", "03_Каналы_трафика", "05_Отзывы_рейтинги", "06_Технологии_финансы") {
    $ws = $wb.Worksheets.Item($sn)
    $ur = $ws.UsedRange; $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
    $r0 = $ur.Row; $c0 = $ur.Column
    $fm = $ur.Formula
    $hdr = 0
    for ($r = 1; $r -le $nr; $r++) { $v = $fm[$r, 1]; if ($v -is [string] -and $v -match '^Конкурент') { $hdr = $r; break } }
    if ($hdr -eq 0) { continue }
    $lastc = 1
    for ($c = 1; $c -le $nc; $c++) { if ($fm[$hdr, $c] -ne $null -and ([string]$fm[$hdr, $c]).Trim() -ne "") { $lastc = $c } }
    for ($r = $hdr + 1; $r -le $nr; $r++) {
        $a = $fm[$r, 1]
        if (-not ($a -is [string] -and $a -match '\.')) { continue }
        for ($c = 2; $c -le $lastc; $c++) {
            $v = $fm[$r, $c]
            if ($v -eq $null -or ([string]$v).Trim() -eq "") {
                $ws.Cells.Item($r0 + $r - 1, $c0 + $c - 1).Value2 = $nd
                $written["$sn|$($r0 + $r - 1)|$($c0 + $c - 1)"] = 1
            }
        }
    }
}
$xl.CalculateFull()
$after = Snapshot $wb
$diff = 0
foreach ($k in $before.Keys) {
    $b = $before[$k][0]; $a = $after[$k][0]; $r0 = $before[$k][1]; $c0 = $before[$k][2]
    $nr = $b.GetLength(0); $nc = $b.GetLength(1)
    for ($r = 1; $r -le $nr; $r++) {
        for ($c = 1; $c -le $nc; $c++) {
            if ($written.ContainsKey("$k|$($r0 + $r - 1)|$($c0 + $c - 1)")) { continue }
            if ([string]$b[$r, $c] -ne [string]$a[$r, $c]) { $diff++ }
        }
    }
}
if ($diff -eq 0) { $wb.Save(); "OK: записано н/д в $($written.Count) ячеек, расчеты не изменились" }
else { "ОТКАТ: изменились $diff расчетных ячеек" }
$wb.Close($false)
$xl.Quit()
