Attribute VB_Name = "LR1_NotaCode_Trend"
'==============================================================================
' ЛР1 "Анализ интереса аудитории и фиксация границ рынка" - проект NotaCode
' Модуль заполнения Excel-шаблона методички
'   shablon_excel_ocenka_trenda_sezonnosti_delovyh_ciklov_google_trends_wordstat.xlsx
'
' ЧТО ДЕЛАЕТ МАКРОС Main
'   1. Проверяет, что активная книга - копия шаблона (есть листы Ввод_данных,
'      Настройки, Расчет, Сезонность, Оценка_тренда, Деловые_циклы, Панель,
'      Журнал_выгрузки).
'   2. Создаёт (если нет) служебный лист "Ввод_NotaCode" - сюда вы вносите
'      помесячную частотность Яндекс Вордстат (столбец C) после досбора и,
'      при желании, собственный ряд Google Trends по Беларуси (столбец B).
'   3. Берёт помесячный ряд Google Trends:
'        - по умолчанию из архивного файла старой версии
'          ЛР1_Тренд_Сезонность_DiagramCode.xlsx (лист "01_Данные_и_расчёт",
'          A3:B42 - PlantUML, Worldwide, 5 лет, снято 2026-09-17, 40 месяцев);
'        - если файл не найден/отменён выбор или на листе "Ввод_NotaCode"
'          в ячейке G4 стоит "ЛИСТ" - со столбцов A:B листа "Ввод_NotaCode".
'   4. Очищает демонстрационные строки листа "Ввод_данных" (A5:F64) и
'      вносит ряд NotaCode: Период, Кластер, GT, Вордстат (если внесён), Регион,
'      Комментарий.
'   5. Заполняет "Журнал_выгрузки" (GT - фактические параметры замера,
'      Вордстат - строка-заготовка с пометкой [ДОСНЯТЬ], либо параметры,
'      если вы их внесли на листе "Ввод_NotaCode").
'   6. Пересчитывает книгу и строит лист "График_NotaCode": динамика
'      сводного индекса + скользящее среднее 12 мес., сезонные индексы,
'      циклический коэффициент.
'   7. Выводит итоговую формулировку с листа "Панель" (ячейка A17).
'
' ДОПОЛНИТЕЛЬНЫЕ МАКРОСЫ
'   ApplyCenteredMA  - поправка к шаблону: заменяет в "Расчет"!H5:H64
'                      скользящее среднее 12 мес. "назад" на центрированное
'                      2x12 (в растущем ряду "хвостовое" среднее отстаёт, и
'                      циклический коэффициент систематически > 1,1).
'   RestoreTemplateMA - вернуть исходные формулы шаблона в "Расчет"!H5:H64.
'
' КАК ЗАПУСТИТЬ
'   1. Скопируйте шаблон из папки "Материалы" в папку ЛР1, например как
'      ЛР1_Тренд_Сезонность_NotaCode.xlsm (Сохранить как -> "Книга Excel с
'      поддержкой макросов").
'   2. Откройте копию, Alt+F11 -> File -> Import File... -> выберите этот
'      файл LR1_NotaCode_Trend.bas.
'   3. Alt+F8 -> Main -> Выполнить.
'   4. После досбора Вордстата: внесите числа в "Ввод_NotaCode"!C5:C64 и
'      параметры в G6:G9, снова запустите Main.
'
' Файл сохранён в кодировке Windows-1251 (иначе кириллица в редакторе VBA
' отобразится некорректно).
'==============================================================================
Option Explicit

Private Const SH_INPUT As String = "Ввод_данных"
Private Const SH_SETTINGS As String = "Настройки"
Private Const SH_CALC As String = "Расчет"
Private Const SH_SEASON As String = "Сезонность"
Private Const SH_TREND As String = "Оценка_тренда"
Private Const SH_CYCLE As String = "Деловые_циклы"
Private Const SH_PANEL As String = "Панель"
Private Const SH_LOG As String = "Журнал_выгрузки"
Private Const SH_NOTA As String = "Ввод_NotaCode"
Private Const SH_CHART As String = "График_NotaCode"

Private Const ARCHIVE_NAME As String = "ЛР1_Тренд_Сезонность_DiagramCode.xlsx"
Private Const ARCHIVE_ABS As String = _
    "E:\ИИТ\ЭлектронныйБизнес\_архив_DiagramCode\labs\lab1-demand\ЛР1_Тренд_Сезонность_DiagramCode.xlsx"
Private Const ARCHIVE_SHEET As String = "01_Данные_и_расчёт"

Private Const CLUSTER_NAME As String = "diagram as code: PlantUML (прокси категории NotaCode)"
Private Const MAX_ROWS As Long = 60          ' шаблон рассчитан на строки 5..64

'------------------------------------------------------------------------------
Public Sub Main()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If Not CheckTemplate(wb) Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    EnsureNotaSheet wb

    Dim periods() As Date, gt() As Variant, n As Long, src As String
    n = LoadGT(wb, periods, gt, src)
    If n = 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "Нет данных Google Trends: ни архивного файла, ни значений на листе " & _
               SH_NOTA & " (A5:B64).", vbExclamation
        Exit Sub
    End If

    Dim ws() As Variant, nWs As Long
    nWs = LoadWordstat(wb, periods, n, ws)

    FillInput wb, periods, gt, ws, n, src
    FillLog wb, periods, n, src, nWs

    ' Если Вордстата нет - сводный индекс шаблона сам берёт только GT
    ' (формула Расчет!G: при одном источнике используется имеющийся).
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    BuildCharts wb, n

    Application.ScreenUpdating = True
    wb.Worksheets(SH_PANEL).Activate

    MsgBox "Готово. Внесено месяцев: " & n & " (" & Format(periods(1), "yyyy-mm") & " ... " & _
           Format(periods(n), "yyyy-mm") & ")." & vbCrLf & _
           "Источник GT: " & src & vbCrLf & _
           "Вордстат: " & IIf(nWs > 0, nWs & " мес. внесено", "не внесён [ДОСНЯТЬ]") & vbCrLf & vbCrLf & _
           "Итог (Панель!A17):" & vbCrLf & CStr(wb.Worksheets(SH_PANEL).Range("A17").Value), _
           vbInformation, "ЛР1 NotaCode"
End Sub

'------------------------------------------------------------------------------
Private Function CheckTemplate(wb As Workbook) As Boolean
    Dim names As Variant, i As Long
    names = Array(SH_INPUT, SH_SETTINGS, SH_CALC, SH_SEASON, SH_TREND, SH_CYCLE, SH_PANEL, SH_LOG)
    For i = LBound(names) To UBound(names)
        If Not SheetExists(wb, CStr(names(i))) Then
            MsgBox "В активной книге нет листа " & names(i) & ". Откройте копию шаблона " & _
                   "shablon_excel_ocenka_trenda_... и запустите макрос снова.", vbCritical
            CheckTemplate = False
            Exit Function
        End If
    Next i
    CheckTemplate = True
End Function

Private Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Sheets(nm)
    SheetExists = Not s Is Nothing
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Служебный лист для личного досбора
Private Sub EnsureNotaSheet(wb As Workbook)
    Dim sh As Worksheet
    If SheetExists(wb, SH_NOTA) Then Exit Sub
    Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    sh.Name = SH_NOTA
    With sh
        .Range("A1").Value = "NotaCode: данные для досбора (заполняется студенткой)"
        .Range("A1").Font.Bold = True
        .Range("A2").Value = "Столбец B - свой ряд Google Trends (Беларусь, 5 лет), если нужен вместо архивного; " & _
                             "столбец C - Яндекс Вордстат, число запросов в месяц по кластеру."
        .Range("A4:D4").Value = Array("Период (1-е число)", "Google Trends, 0-100", _
                                      "Яндекс Вордстат, запросов", "Комментарий")
        .Range("A4:D4").Font.Bold = True
        .Range("F4").Value = "Источник GT (АРХИВ / ЛИСТ)"
        .Range("G4").Value = "АРХИВ"
        .Range("F6").Value = "Вордстат: дата выгрузки"
        .Range("F7").Value = "Вордстат: запрос/кластер (с операторами)"
        .Range("G7").Value = "(диаграмма|диаграммы) онлайн -скачать -бесплатно"
        .Range("F8").Value = "Вордстат: регион"
        .Range("G8").Value = "Беларусь"
        .Range("F9").Value = "Вордстат: имя файла"
        .Range("G9").Value = "YW_ГГГГ-ММ-ДД_BY_diagram-online_dynamic.xlsx"
        .Range("F11").Value = "GT (если ЛИСТ): регион / период / имя файла"
        .Range("G11").Value = "Беларусь"
        .Range("H11").Value = "5 лет"
        .Range("I11").Value = "GT_ГГГГ-ММ-ДД_BY_5y_plantuml_time.csv"
        .Range("A5:A64").NumberFormat = "yyyy-mm"
        .Columns("A:D").ColumnWidth = 22
        .Columns("F").ColumnWidth = 40
        .Columns("G").ColumnWidth = 45
        .Range("G6,G7,G8,G9,C5:C64").Interior.Color = RGB(255, 242, 204)
    End With
End Sub

'------------------------------------------------------------------------------
' Загрузка GT. Возвращает число месяцев.
Private Function LoadGT(wb As Workbook, periods() As Date, gt() As Variant, src As String) As Long
    Dim nota As Worksheet, mode As String
    Set nota = wb.Worksheets(SH_NOTA)
    mode = UCase$(Trim$(CStr(nota.Range("G4").Value)))

    If mode <> "ЛИСТ" Then
        Dim path As String
        path = FindArchive(wb)
        If path <> "" Then
            LoadGT = LoadFromArchive(path, periods, gt)
            If LoadGT > 0 Then
                src = "архив DiagramCode (" & ARCHIVE_NAME & "), PlantUML, Worldwide, 5 лет, снято 2026-09-17"
                ' Для наглядности копируем ряд на служебный лист
                Dim i As Long
                For i = 1 To LoadGT
                    nota.Cells(4 + i, 1).Value = periods(i)
                    If IsEmpty(nota.Cells(4 + i, 2).Value) Then nota.Cells(4 + i, 2).Value = gt(i)
                Next i
                Exit Function
            End If
        End If
    End If

    ' Резерв: лист Ввод_NotaCode
    LoadGT = LoadFromNota(nota, periods, gt)
    src = "лист " & SH_NOTA & " (данные внесены вручную)"
End Function

Private Function FindArchive(wb As Workbook) As String
    Dim cands(1 To 4) As String, i As Long, base As String
    base = wb.path
    cands(1) = base & "\" & ARCHIVE_NAME
    cands(2) = base & "\..\_архив_DiagramCode\labs\lab1-demand\" & ARCHIVE_NAME
    cands(3) = base & "\..\..\_архив_DiagramCode\labs\lab1-demand\" & ARCHIVE_NAME
    cands(4) = ARCHIVE_ABS
    For i = 1 To 4
        If Len(cands(i)) > 0 Then
            If Dir(cands(i)) <> "" Then FindArchive = cands(i): Exit Function
        End If
    Next i
    Dim f As Variant
    f = Application.GetOpenFilename("Excel (*.xlsx), *.xlsx", , _
        "Укажите архивный файл " & ARCHIVE_NAME & " (Отмена - взять данные с листа " & SH_NOTA & ")")
    If VarType(f) = vbString Then FindArchive = CStr(f)
End Function

Private Function LoadFromArchive(path As String, periods() As Date, gt() As Variant) As Long
    Dim src As Workbook, sh As Worksheet, r As Long, n As Long, v As Variant
    On Error GoTo Fail
    Set src = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    Set sh = src.Worksheets(ARCHIVE_SHEET)
    ReDim periods(1 To MAX_ROWS): ReDim gt(1 To MAX_ROWS)
    For r = 3 To 3 + MAX_ROWS - 1
        v = sh.Cells(r, 1).Value
        If IsEmpty(v) Then Exit For
        If Not IsNumeric(sh.Cells(r, 2).Value) Or IsEmpty(sh.Cells(r, 2).Value) Then Exit For
        n = n + 1
        periods(n) = ParsePeriod(v)
        gt(n) = CDbl(sh.Cells(r, 2).Value)
    Next r
    src.Close SaveChanges:=False
    LoadFromArchive = n
    Exit Function
Fail:
    If Not src Is Nothing Then src.Close SaveChanges:=False
    LoadFromArchive = 0
End Function

Private Function LoadFromNota(nota As Worksheet, periods() As Date, gt() As Variant) As Long
    Dim r As Long, n As Long
    ReDim periods(1 To MAX_ROWS): ReDim gt(1 To MAX_ROWS)
    For r = 5 To 64
        If IsEmpty(nota.Cells(r, 1).Value) Then Exit For
        If IsNumeric(nota.Cells(r, 2).Value) And Not IsEmpty(nota.Cells(r, 2).Value) Then
            n = n + 1
            periods(n) = ParsePeriod(nota.Cells(r, 1).Value)
            gt(n) = CDbl(nota.Cells(r, 2).Value)
        End If
    Next r
    LoadFromNota = n
End Function

' "2021-09" (текст) или дата -> 1-е число месяца
Private Function ParsePeriod(v As Variant) As Date
    Dim s As String
    If IsDate(v) And VarType(v) = vbDate Then
        ParsePeriod = DateSerial(Year(v), Month(v), 1)
    Else
        s = Trim$(CStr(v))
        ParsePeriod = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), 1)
    End If
End Function

'------------------------------------------------------------------------------
' Вордстат: сопоставление по периоду со столбцом C листа Ввод_NotaCode
Private Function LoadWordstat(wb As Workbook, periods() As Date, n As Long, ws() As Variant) As Long
    Dim nota As Worksheet, r As Long, i As Long, cnt As Long, p As Date
    Set nota = wb.Worksheets(SH_NOTA)
    ReDim ws(1 To n)
    For r = 5 To 64
        If Not IsEmpty(nota.Cells(r, 1).Value) And IsNumeric(nota.Cells(r, 3).Value) _
           And Not IsEmpty(nota.Cells(r, 3).Value) Then
            p = ParsePeriod(nota.Cells(r, 1).Value)
            For i = 1 To n
                If periods(i) = p Then ws(i) = CDbl(nota.Cells(r, 3).Value): cnt = cnt + 1: Exit For
            Next i
        End If
    Next r
    LoadWordstat = cnt
End Function

'------------------------------------------------------------------------------
Private Sub FillInput(wb As Workbook, periods() As Date, gt() As Variant, ws() As Variant, _
                      n As Long, src As String)
    Dim sh As Worksheet, i As Long, r As Long
    Set sh = wb.Worksheets(SH_INPUT)
    sh.Range("A5:F64").ClearContents
    sh.Range("A2").Value = "NotaCode: кластер "diagram as code" (прокси - PlantUML). GT - индекс 0-100; " & _
        "Вордстат - абсолютное число запросов по кластеру (досбор). Источник GT: " & src & "."
    For i = 1 To n
        r = 4 + i
        sh.Cells(r, 1).Value = periods(i)
        sh.Cells(r, 1).NumberFormat = "yyyy-mm"
        sh.Cells(r, 2).Value = CLUSTER_NAME
        sh.Cells(r, 3).Value = gt(i)
        If Not IsEmpty(ws(i)) Then sh.Cells(r, 4).Value = ws(i)
        If InStr(src, "Worldwide") > 0 Then
            sh.Cells(r, 5).Value = "Весь мир (прокси; [ДОСНЯТЬ] ряд по Беларуси)"
        Else
            sh.Cells(r, 5).Value = "Республика Беларусь"
        End If
        sh.Cells(r, 6).Value = "GT: первая неделя месяца" & IIf(IsEmpty(ws(i)), "; Вордстат [ДОСНЯТЬ]", "")
    Next i
End Sub

'------------------------------------------------------------------------------
Private Sub FillLog(wb As Workbook, periods() As Date, n As Long, src As String, nWs As Long)
    Dim sh As Worksheet, nota As Worksheet, per As String
    Set sh = wb.Worksheets(SH_LOG)
    Set nota = wb.Worksheets(SH_NOTA)
    sh.Range("A4:H30").ClearContents
    per = Format(periods(1), "yyyy-mm") & "-" & Format(periods(n), "yyyy-mm")

    If InStr(src, "архив") > 0 Then
        sh.Range("A4:H4").Value = Array(DateSerial(2026, 9, 17), "Google Trends", _
            "diagram as code, PlantUML, Mermaid.js", "Весь мир (Worldwide)", "5 лет", "веб-график", _
            "(снято вручную через браузер, CSV не сохранён)", _
            "Все категории, веб-поиск; средние: diagram as code 3, PlantUML 41, Mermaid.js 2")
        sh.Range("A5:H5").Value = Array(DateSerial(2026, 9, 17), "Google Trends", _
            "UML diagram generator, ERD tool, text to diagram", "Весь мир (Worldwide)", "5 лет", "веб-график", _
            "(CSV не сохранён)", "Средние: 2 / 4 / 26; регионы не снимались (мало данных)")
        sh.Range("A6:H6").Value = Array(DateSerial(2026, 9, 17), "Google Trends", _
            "PlantUML, text to diagram (2 термина)", "Весь мир (Worldwide)", "5 лет; в шаблон " & per, _
            "значения перенесены в XLSX", ARCHIVE_NAME, _
            "Помесячно (1-я неделя месяца); регионы: PlantUML - Беларусь №5, Россия №4; " & _
            "related: text to diagram ai - Breakout")
        sh.Range("A7:H7").Value = Array("[ДОСНЯТЬ]", "Google Trends", _
            "PlantUML; diagram as code; UML диаграмма онлайн; BPMN; IDEF0", "Беларусь", "5 лет + 12 мес.", _
            "CSV", "GT_ГГГГ-ММ-ДД_BY_5y_<запрос>_time.csv", _
            "Повторить замер по Беларуси + карточки регионов и связанных запросов")
    Else
        sh.Range("A4:H4").Value = Array(Date, "Google Trends", CLUSTER_NAME, _
            nota.Range("G11").Value, nota.Range("H11").Value & "; в шаблон " & per, "CSV", _
            nota.Range("I11").Value, "Ряд внесён на лист " & SH_NOTA)
    End If

    Dim rw As Long
    rw = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row + 1
    If rw < 5 Then rw = 5
    If nWs > 0 Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 8)).Value = Array(nota.Range("G6").Value, "Яндекс Вордстат", _
            nota.Range("G7").Value, nota.Range("G8").Value, per & " (" & nWs & " мес.)", "XLSX/CSV", _
            nota.Range("G9").Value, "Вкладка Динамика, все устройства, помесячно")
    Else
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 8)).Value = Array("[ДОСНЯТЬ]", "Яндекс Вордстат", _
            nota.Range("G7").Value, "Беларусь", per, "XLSX/CSV", _
            "YW_ГГГГ-ММ-ДД_BY_<кластер>_dynamic.xlsx", _
            "Требуется вход в Яндекс ID; вкладки Динамика, Регионы, Топы запросов")
    End If
    sh.Columns("A:H").AutoFit
End Sub

'------------------------------------------------------------------------------
Private Sub BuildCharts(wb As Workbook, n As Long)
    Dim sh As Worksheet, co As ChartObject, lastR As Long
    Application.DisplayAlerts = False
    If SheetExists(wb, SH_CHART) Then wb.Worksheets(SH_CHART).Delete
    Application.DisplayAlerts = True
    Set sh = wb.Worksheets.Add(After:=wb.Worksheets(SH_PANEL))
    sh.Name = SH_CHART
    sh.Range("A1").Value = "NotaCode - графики по данным шаблона (строятся макросом Main)"
    sh.Range("A1").Font.Bold = True
    lastR = 4 + n

    Dim calc As Worksheet: Set calc = wb.Worksheets(SH_CALC)

    ' 1. Сводный индекс + тренд
    Set co = sh.ChartObjects.Add(Left:=10, Top:=30, Width:=720, Height:=300)
    With co.Chart
        .ChartType = xlLine
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "Сводный индекс (GT" & IIf(Application.WorksheetFunction.Count(calc.Range("E5:E64")) > 0, _
                    " + Вордстат", "") & ")"
            .Values = calc.Range("G5:G" & lastR)
            .XValues = calc.Range("A5:A" & lastR)
            .Format.Line.ForeColor.RGB = RGB(42, 120, 214)
            .Format.Line.Weight = 2
        End With
        With .SeriesCollection.NewSeries
            .Name = "Тренд: скользящее среднее 12 мес."
            .Values = calc.Range("H5:H" & lastR)
            .XValues = calc.Range("A5:A" & lastR)
            .Format.Line.ForeColor.RGB = RGB(235, 104, 52)
            .Format.Line.Weight = 2
            .Format.Line.DashStyle = msoLineDash
        End With
        .HasTitle = True
        .ChartTitle.Text = "Динамика интереса к diagram as code (NotaCode) и тренд"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm"
    End With

    ' 2. Сезонные индексы
    Set co = sh.ChartObjects.Add(Left:=10, Top:=345, Width:=350, Height:=260)
    With co.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "Сезонный индекс"
            .Values = wb.Worksheets(SH_SEASON).Range("D4:D15")
            .XValues = wb.Worksheets(SH_SEASON).Range("B4:B15")
            .Format.Fill.ForeColor.RGB = RGB(42, 120, 214)
        End With
        .HasTitle = True
        .ChartTitle.Text = "Сезонные индексы (1,0 = средний уровень)"
        .HasLegend = False
    End With

    ' 3. Циклический коэффициент
    Set co = sh.ChartObjects.Add(Left:=380, Top:=345, Width:=350, Height:=260)
    With co.Chart
        .ChartType = xlLineMarkers
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "K = Y / (T x S)"
            .Values = calc.Range("K5:K" & lastR)
            .XValues = calc.Range("A5:A" & lastR)
            .Format.Line.ForeColor.RGB = RGB(27, 175, 122)
        End With
        .HasTitle = True
        .ChartTitle.Text = "Циклический коэффициент (норма 0,9-1,1)"
        .HasLegend = False
        .Axes(xlCategory).TickLabels.NumberFormat = "yyyy-mm"
    End With

    ' Сводка показателей рядом с графиками
    sh.Range("L3").Value = "Показатель": sh.Range("M3").Value = "Значение"
    sh.Range("L4").Value = "Месяцев": sh.Range("M4").Formula = "='" & SH_TREND & "'!B4"
    sh.Range("L5").Value = "Среднее первых 12": sh.Range("M5").Formula = "='" & SH_TREND & "'!B5"
    sh.Range("L6").Value = "Среднее последних 12": sh.Range("M6").Formula = "='" & SH_TREND & "'!B6"
    sh.Range("L7").Value = "Относит. изменение": sh.Range("M7").Formula = "='" & SH_TREND & "'!B8"
    sh.Range("L8").Value = "Наклон, п./мес.": sh.Range("M8").Formula = "='" & SH_TREND & "'!B9"
    sh.Range("L9").Value = "Тренд": sh.Range("M9").Formula = "='" & SH_TREND & "'!B10"
    sh.Range("L10").Value = "Амплитуда сезонности": sh.Range("M10").Formula = "='" & SH_SEASON & "'!B21"
    sh.Range("L11").Value = "Сила сезонности": sh.Range("M11").Formula = "='" & SH_SEASON & "'!B22"
    sh.Range("L12").Value = "Месяц максимума": sh.Range("M12").Formula = "='" & SH_SEASON & "'!B23"
    sh.Range("L13").Value = "Месяц минимума": sh.Range("M13").Formula = "='" & SH_SEASON & "'!B24"
    sh.Range("L14").Value = "Текущая фаза": sh.Range("M14").Formula = "='" & SH_CYCLE & "'!B7"
    sh.Range("L15").Value = "Текущий K": sh.Range("M15").Formula = "='" & SH_CYCLE & "'!B6"
    sh.Range("M7").NumberFormat = "0.0%"
    sh.Range("M5:M6,M8,M10,M15").NumberFormat = "0.000"
    sh.Range("L3:M3").Font.Bold = True
    sh.Columns("L:M").AutoFit
End Sub

'------------------------------------------------------------------------------
' Поправка: центрированное скользящее среднее 2x12 в Расчет!H
Public Sub ApplyCenteredMA()
    Dim calc As Worksheet, r As Long
    If Not CheckTemplate(ActiveWorkbook) Then Exit Sub
    Set calc = ActiveWorkbook.Worksheets(SH_CALC)
    For r = 5 To 64
        If r >= 11 And r <= 58 Then
            calc.Cells(r, 8).FormulaR1C1 = _
                "=IF(OR(RC7="""",COUNT(R[-6]C7:R[6]C7)<13),""""," & _
                "(AVERAGE(R[-6]C7:R[5]C7)+AVERAGE(R[-5]C7:R[6]C7))/2)"
        Else
            calc.Cells(r, 8).Value = ""
        End If
    Next r
    calc.Range("H4").Value = "Тренд: центрированное скользящее среднее 2x12 (поправка NotaCode)"
    Application.CalculateFull
    MsgBox "В Расчет!H установлено центрированное скользящее среднее 2x12. " & _
           "Последние 6 месяцев ряда тренда не имеют - текущая фаза на листе Деловые_циклы " & _
           "будет пустой; смотрите последний рассчитанный месяц в таблице ниже.", vbInformation
End Sub

' Возврат исходных формул шаблона
Public Sub RestoreTemplateMA()
    Dim calc As Worksheet, r As Long
    If Not CheckTemplate(ActiveWorkbook) Then Exit Sub
    Set calc = ActiveWorkbook.Worksheets(SH_CALC)
    For r = 5 To 64
        If r >= 16 Then
            calc.Cells(r, 8).FormulaR1C1 = _
                "=IF(OR(RC7="""",COUNT(R5C7:RC7)<12),"""",AVERAGE(R[-11]C7:RC7))"
        Else
            calc.Cells(r, 8).FormulaR1C1 = "=IF(OR(RC7="""",COUNT(R5C7:RC7)<12),"""","""")"
        End If
    Next r
    calc.Range("H4").Value = "Тренд: скользящее среднее 12 мес."
    Application.CalculateFull
End Sub
