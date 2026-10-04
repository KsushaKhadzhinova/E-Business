$t = "C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$items = Get-Content "$t\port2.json" -Raw -Encoding utf8 | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wbs = $xl.Workbooks
$groups = $items | Group-Object p
foreach ($g in $groups) {
    $wb = $wbs.Open($g.Name)
    $n = 0
    foreach ($it in $g.Group) {
        $ws = $wb.Worksheets.Item($it.s)
        $cell = $ws.Range($it.c)
        if ($it.v -is [long] -or $it.v -is [int]) { $cell.Value2 = [double]$it.v } else { $cell.Value2 = $it.v }
        $n++
    }
    foreach ($ws in $wb.Worksheets) {
        foreach ($lo in $ws.ListObjects) {
            $rng = $lo.Range
            $lastRow = $ws.Cells.Item($ws.Rows.Count, $rng.Column).End(-4162).Row
            $endRow = $rng.Row + $rng.Rows.Count - 1
            if ($lastRow -gt $endRow) {
                try { $lo.Resize($ws.Range($ws.Cells.Item($rng.Row, $rng.Column), $ws.Cells.Item($lastRow, $rng.Column + $rng.Columns.Count - 1))) } catch {}
            }
        }
    }
    $xl.CalculateFull()
    $wb.Save(); $wb.Close($false)
    "$($g.Name | Split-Path -Leaf): записано $n"
}
$xl.Quit()

