Set-Location "E:\ИИТ\ЭлектронныйБизнес"
$t = "C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wbs = $xl.Workbooks
$rows = @()
$root = "E:\ИИТ\ЭлектронныйБизнес\СДАЧА\"
$files = Get-ChildItem $root -Recurse -Filter *.xlsx
foreach ($f in $files) {
    $relWin = $f.FullName.Substring($root.Length)
    $rel = "СДАЧА/" + ($relWin -replace '\\', '/')
    $src = $null
    foreach ($c in @("9735d9f", "efd0279")) {
        git cat-file -e "${c}:$rel" 2>$null
        if ($LASTEXITCODE -eq 0) { $src = $c; break }
    }
    if (-not $src) { $rows += [pscustomobject]@{File = $f.Name; Status = "нет оригинала в git" }; continue }
    $tmp = "$t\orig\$($f.Name)"
    cmd /c "git show ${src}:$rel > `"$tmp`""
    try {
        $wb = $wbs.Open($tmp)
        $fixed = 0
        try {
            $ws = $wb.Sheets.Item("Журнал_выгрузки")
            $last = $ws.UsedRange.Rows.Count
            for ($r = 2; $r -le $last; $r++) {
                $cell = $ws.Cells.Item($r, 1)
                $v = $cell.Value2
                if ($v -is [double] -and $v -gt 40000) {
                    $cell.NumberFormat = "@"
                    $cell.Value2 = [DateTime]::FromOADate($v).ToString("dd.MM.yyyy")
                    $fixed++
                }
            }
        } catch {}
        $xl.CalculateFull()
        Copy-Item -LiteralPath $f.FullName -Destination "$t\broken\$($f.Name)" -Force
        $out = "$t\orig\out_$($f.Name)"
        $wb.SaveAs($out, 51)
        $wb.Close($false)
        Copy-Item -LiteralPath $out -Destination $f.FullName -Force
        $rows += [pscustomobject]@{File = $f.Name; Status = "OK из $src, дат исправлено: $fixed" }
    } catch {
        $rows += [pscustomobject]@{File = $f.Name; Status = "ОШИБКА: " + $_.Exception.Message }
        try { $wb.Close($false) } catch {}
    }
}
$xl.Quit()
$rows | Format-Table -AutoSize
