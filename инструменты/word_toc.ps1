$base = "E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
$files = @(
    "ЛР1\ОТЧЕТ_ЛР1_NotaCode.docx",
    "ЛР2-3\ОТЧЕТ_ЛР2-3_NotaCode.docx",
    "ЛР4\ОТЧЕТ_ЛР4_NotaCode.docx",
    "ЛР5\ОТЧЕТ_ЛР5_NotaCode.docx",
    "ЛР6\ОТЧЕТ_ЛР6_NotaCode.docx",
    "ЛР7\ОТЧЕТ_ЛР7_NotaCode.docx",
    "ЛР7\РЕЗЮМЕ_ПРОЕКТА_NotaCode.docx",
    "Финальный_отчет\ОТЧЕТ_по_лабораторным_работам_1-7_NotaCode.docx"
)
Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow() | Out-Null }
Start-Sleep 3
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
foreach ($f in $files) {
    $p = Join-Path $base $f
    try {
        $d = $w.Documents.Open($p)
        $d.Repaginate()
        $n = $d.TablesOfContents.Count
        for ($i = 1; $i -le $n; $i++) { $d.TablesOfContents.Item($i).Update() }
        $d.Fields.Update() | Out-Null
        for ($i = 1; $i -le $d.TablesOfContents.Count; $i++) { $d.TablesOfContents.Item($i).Update() }
        $d.Repaginate()
        $pages = $d.ComputeStatistics(2)
        $d.Save(); $d.Close()
        "$f : TOC=$n, pages=$pages"
    } catch {
        "$f : ERROR $($_.Exception.Message)"
        try { $d.Close($false) } catch {}
    }
}
$w.Quit()
