$t = "C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$items = Get-Content "$t\fill3.json" -Raw -Encoding utf8 | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
foreach ($g in ($items | Group-Object p)) {
    $wb = $xl.Workbooks.Open($g.Name); $n = 0
    foreach ($it in $g.Group) {
        $cell = $wb.Worksheets.Item($it.s).Range($it.c)
        if ($cell.HasFormula) { continue }
        $cur = $cell.Value2
        if ($cur -ne $null -and ([string]$cur).Trim() -ne "") { continue }
        if ($it.v -is [long] -or $it.v -is [int]) { $cell.Value2 = [double]$it.v } else { $cell.Value2 = $it.v }
        $n++
    }
    $xl.CalculateFull(); $wb.Save(); $wb.Close($false)
    "$($g.Name | Split-Path -Leaf): заполнено $n"
}
$xl.Quit()
