$t = "C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$items = Get-Content "$t\yo.json" -Raw -Encoding utf8 | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wbs = $xl.Workbooks
foreach ($g in ($items | Group-Object p)) {
    $wb = $wbs.Open($g.Name)
    $n = 0
    foreach ($it in $g.Group) {
        $cell = $wb.Worksheets.Item($it.s).Range($it.c)
        if ($it.f) { $cell.Formula = $it.v } else { $cell.Value2 = $it.v }
        $n++
    }
    $wb.Save(); $wb.Close($false)
    "$($g.Name | Split-Path -Leaf): $n"
}
$xl.Quit()
