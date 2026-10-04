Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wbs = $xl.Workbooks
$rows = @()
$files = Get-ChildItem "E:\ИИТ\ЭлектронныйБизнес\СДАЧА" -Recurse -Filter *.xlsx | Where-Object { $_.Name -match "^(analiz_|klassifikaciya|6_2_|6_3_)" }
$dash = [string][char]0x2014
$en = [string][char]0x2013
$yo = [string][char]0x0451
$Yo = [string][char]0x0401
$e = [string][char]0x0435
$E = [string][char]0x0415
foreach ($f in $files) {
    $info = ""
    try {
        $wb = $wbs.Open($f.FullName)
        $dates = 0
        foreach ($ws in $wb.Worksheets) {
            $ur = $ws.UsedRange
            $vals = $ur.Value2
            $r0 = $ur.Row; $c0 = $ur.Column
            $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
            if ($nr -lt 2 -or $nc -lt 1) { continue }
            $hdrCols = @()
            for ($c = 1; $c -le $nc; $c++) {
                for ($r = 1; $r -le [Math]::Min(6, $nr); $r++) {
                    $v = $vals[$r, $c]
                    if ($v -is [string] -and $v -match "^Дата") { $hdrCols += $c; break }
                }
            }
            foreach ($c in $hdrCols) {
                for ($r = 1; $r -le $nr; $r++) {
                    $v = $vals[$r, $c]
                    if ($v -is [double] -and $v -gt 40000 -and $v -lt 60000) {
                        $cell = $ws.Cells.Item($r0 + $r - 1, $c0 + $c - 1)
                        if (-not $cell.HasFormula) {
                            $cell.NumberFormat = "@"
                            $cell.Value2 = [DateTime]::FromOADate($v).ToString("dd.MM.yyyy")
                            $dates++
                        }
                    }
                }
            }
            [void]$ws.Cells.Replace($dash, $en, 2, 1, $true)
            [void]$ws.Cells.Replace($yo, $e, 2, 1, $true)
            [void]$ws.Cells.Replace($Yo, $E, 2, 1, $true)
            [void]$ws.Cells.Replace("USD", "у.е.", 2, 1, $true)
        }
        $xl.CalculateFull()
        $wb.Save()
        $wb.Close($false)
        $info = "OK, дат: $dates"
    } catch {
        $info = "ОШИБКА: " + $_.Exception.Message
        try { $wb.Close($false) } catch {}
    }
    $rows += [pscustomobject]@{ File = $f.Name; Status = $info }
}
$xl.Quit()
$rows | Format-Table -AutoSize

