Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wbs = $xl.Workbooks
$pairs = @(
    @("Опрос студентов сгенерирован (ЛР4)", "Опрос студентов пробный, по условным профилям (ЛР4)"),
    @("опрос студентов сгенерирован", "опрос студентов пробный, по условным профилям"),
    @("данные сгенерированы", "данные пробного опроса"),
    @(". не собрано:", ". Планируется собрать:")
)
$files = Get-ChildItem "E:\ИИТ\ЭлектронныйБизнес\СДАЧА" -Recurse -Filter *.xlsx | Where-Object { $_.Name -match "^(analiz_|klassifikaciya|6_2_|6_3_)" }
foreach ($f in $files) {
    try {
        $wb = $wbs.Open($f.FullName)
        foreach ($ws in $wb.Worksheets) {
            foreach ($p in $pairs) { [void]$ws.Cells.Replace($p[0], $p[1], 2, 1, $false) }
        }
        $wb.Save(); $wb.Close($false)
    } catch { "ERR $($f.Name): $($_.Exception.Message)"; try { $wb.Close($false) } catch {} }
}
$xl.Quit()
"done"

