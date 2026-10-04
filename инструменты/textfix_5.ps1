Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
$wbs = $xl.Workbooks
$box = [char]::ConvertFromUtf32(0x1F532)
$abbr = @("у.е","т.д","т.п","т.е","др","мес","тыс","млн","млрд","руб","шт","стр","рис","табл","им","г","гг","ч","мин","сек","п","пп","см","напр","пр","ср","прим","е","д","к","ред","изд","вып","ст","обл","гр","ул","пос","англ","рус")
$report = @()
$files = Get-ChildItem "E:\ИИТ\ЭлектронныйБизнес\СДАЧА" -Recurse -Filter *.xlsx | Where-Object { $_.Name -match "^(analiz_|klassifikaciya|6_2_|6_3_)" }
foreach ($f in $files) {
    try {
        $wb = $wbs.Open($f.FullName)
        $dots = 0; $fx = 0; $boxes = 0
        foreach ($ws in $wb.Worksheets) {
            $ur = $ws.UsedRange
            $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
            if ($ur.Cells.Count -lt 1) { continue }
            # 1. formulas: XLOOKUP -> INDEX/MATCH, drop _xlfn prefix
            $fm = $ur.Formula
            $r0 = $ur.Row; $c0 = $ur.Column
            if ($ur.Cells.Count -eq 1) { continue }
            for ($r = 1; $r -le $nr; $r++) {
                for ($c = 1; $c -le $nc; $c++) {
                    $v = $fm[$r, $c]
                    if ($v -is [string] -and $v.StartsWith("=") -and ($v.Contains("XLOOKUP") -or $v.Contains("_xlfn."))) {
                        $n = $v -replace '_xlfn\.', ''
                        $n = [regex]::Replace($n, 'XLOOKUP\(([^,()]+),([^,()]+),([^,()]+)\)', 'INDEX($3,MATCH($1,$2,0))')
                        if ($n -ne $v) { $ws.Cells.Item($r0 + $r - 1, $c0 + $c - 1).Formula = $n; $fx++ }
                    }
                }
            }
            # 2. trailing dot in constant text cells
            $fm = $ur.Formula
            for ($r = 1; $r -le $nr; $r++) {
                for ($c = 1; $c -le $nc; $c++) {
                    $v = $fm[$r, $c]
                    if ($v -is [string] -and -not $v.StartsWith("=") -and $v.Length -gt 3 -and $v.EndsWith(".") -and -not $v.EndsWith("..")) {
                        $body = $v.Substring(0, $v.Length - 1)
                        $m = [regex]::Match($body, '([A-Za-zА-Яа-яЁё]+)$')
                        $last = if ($m.Success) { $m.Groups[1].Value.ToLower() } else { "" }
                        $isAbbr = ($abbr -contains $last) -and ($body -match '(^|[\s.])' + [regex]::Escape($last) + '$') -and ($last.Length -le 4)
                        if ($body -match '\d\s*(у\.е)$' -or $body -match 'у\.е$') { $isAbbr = $true }
                        if (-not $isAbbr) {
                            $cell = $ws.Cells.Item($r0 + $r - 1, $c0 + $c - 1)
                            $cell.Value2 = $body
                            $dots++
                        }
                    }
                }
            }
            # 3. placeholder boxes
            [void]$ws.Cells.Replace("$box ДОСНЯТЬ", "не собрано", 2, 1, $true)
            [void]$ws.Cells.Replace("$box ", "", 2, 1, $true)
            [void]$ws.Cells.Replace($box, "", 2, 1, $true)
        }
        $xl.CalculateFull()
        $wb.Save(); $wb.Close($false)
        $report += [pscustomobject]@{ File = $f.Name; Status = "OK: формул $fx, точек $dots" }
    } catch {
        $report += [pscustomobject]@{ File = $f.Name; Status = "ОШИБКА: " + $_.Exception.Message }
        try { $wb.Close($false) } catch {}
    }
}
$xl.Quit()
$report | Format-Table -AutoSize

