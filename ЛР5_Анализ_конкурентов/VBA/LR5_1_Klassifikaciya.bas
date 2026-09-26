Attribute VB_Name = "LR5_1_Klassifikaciya"
Option Explicit
' =====================================================================================
' ЛР5 «Анализ конкурентов» — NotaCode. Модуль 1: классификация по уровням конкуренции
' Шаблон: Материалы\Шаблон_классификации_компаний_по_уровням_конкуренции.xlsx
'
' КАК ЗАПУСТИТЬ
'   1. Скопируйте шаблон в папку ЛР5_Анализ_конкурентов и сохраните копию как .xlsm
'      (например, ЛР5_1_Классификация_NotaCode.xlsm). Рядом должна лежать папка VBA\data (TSV-файлы).
'   2. Alt+F11 -> File -> Import File... -> VBA\LR5_1_Klassifikaciya.bas
'   3. Alt+F8 -> LR5_1_Klassifikaciya.Main -> Выполнить.
'   Excel 2016+ (модуль 1 заменяет XLOOKUP на INDEX/MATCH), ADODB.Stream (есть в Windows).
'
' ПОЧЕМУ ДАННЫЕ В TSV. Редактор VBA импортирует .bas в ANSI-кодировке, кириллица в строках
' кода ломается. Поэтому код — только ASCII, а все тексты NotaCode читаются из UTF-8
' файлов VBA\data\k_*.tsv (их создаёт Python\export_vba_data.py из Python\lr5_data.py).
' Если комментарии в редакторе отображаются «кракозябрами» — это косметика, код работает.
'
' ЧТО ДЕЛАЕТ МАКРОС
'   01_Компании (20 компаний), 02_Доказательства (32 факта), 04_Оценка (баллы C1-C10, 2 корректировки,
'   обоснования), 07_Карта; ИСПРАВЛЯЕТ 04_Оценка N4:N103 (#VALUE!: строка x столбец) и 06_Сводка B19:F28 (MATCH давал дубли
'   при равных баллах -> AGGREGATE); вывод и исключённые — 06_Сводка A30; лист Ввод_NotaCode;
'   scatter-карта на 07_Карта и bar-диаграмма баллов на 04_Оценка. Веса 03_Критерии не меняются.
' Повторный запуск безопасен: значения перезаписываются, диаграммы пересоздаются.
' =====================================================================================

' ---------------- общие процедуры (Private — модули не конфликтуют) ----------------
Private gStr As Object

Private Function DataDir(wb As Workbook, ByVal probe As String) As String
    Dim c As Variant, cands(2) As String
    cands(0) = wb.Path & "\VBA\data\"
    cands(1) = wb.Path & "\data\"
    cands(2) = wb.Path & "\..\VBA\data\"
    DataDir = ""
    If wb.Path <> "" Then
        For Each c In cands
            If Dir(CStr(c) & probe) <> "" Then DataDir = CStr(c): Exit Function
        Next c
    End If
    With Application.FileDialog(4)   ' msoFileDialogFolderPicker
        .Title = "LR5: select folder VBA\data"
        If .Show = -1 Then DataDir = .SelectedItems(1) & "\"
    End With
    If DataDir <> "" Then
        If Dir(DataDir & probe) = "" Then DataDir = ""
    End If
End Function

Private Function ReadUtf8(ByVal f As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile f
    ReadUtf8 = st.ReadText(-1)
    st.Close
End Function

Private Sub LoadStrings(ByVal f As String)
    Dim lines() As String, i As Long, p As Long
    Set gStr = CreateObject("Scripting.Dictionary")
    lines = Split(Replace(ReadUtf8(f), vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        p = InStr(lines(i), vbTab)
        If p > 0 Then gStr(Left(lines(i), p - 1)) = Replace(Mid(lines(i), p + 1), "\n", vbLf)
    Next i
End Sub

Private Function S(ByVal key As String) As String
    S = key
    If Not gStr Is Nothing Then
        If gStr.Exists(key) Then S = gStr(key)
    End If
End Function

Private Function SheetByPrefix(wb As Workbook, ByVal p As String) As Worksheet
    Dim ws As Worksheet
    If Left(p, 1) = "+" Then
        p = Mid(p, 2)
        For Each ws In wb.Worksheets
            If ws.Name = p Then Set SheetByPrefix = ws: Exit Function
        Next ws
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = p
        Set SheetByPrefix = ws
        Exit Function
    End If
    For Each ws In wb.Worksheets
        If Left(ws.Name, Len(p)) = p Then Set SheetByPrefix = ws: Exit Function
    Next ws
    Err.Raise vbObjectError + 1, , "Sheet not found: " & p
End Function

Private Function IsNumText(ByVal t As String) As Boolean
    Dim i As Long, ch As String, dots As Long
    IsNumText = False
    If Len(t) = 0 Or Len(t) > 15 Then Exit Function
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        If ch = "." Then
            dots = dots + 1
        ElseIf ch < "0" Or ch > "9" Then
            Exit Function
        End If
    Next i
    IsNumText = (dots <= 1 And Left(t, 1) <> "." And Right(t, 1) <> ".")
End Function

Private Sub PutVal(c As Range, ByVal t As String)
    t = Replace(t, "\n", vbLf)
    If Left(t, 1) = "=" Then
        c.Formula = t
    ElseIf IsNumText(t) Then
        c.Value = Val(t)
    Else
        c.NumberFormat = "@"
        c.Value = t
        c.WrapText = (Len(t) > 40)
    End If
End Sub

' Заполняет лист из TSV: строка 1 = SHEET<TAB>префикс<TAB>ячейка; пустое поле не трогает ячейку
Private Sub LoadBlock(wb As Workbook, ByVal f As String)
    Dim lines() As String, hdr() As String, flds() As String
    Dim ws As Worksheet, r0 As Long, c0 As Long, i As Long, j As Long
    If Dir(f) = "" Then Debug.Print "missing: " & f: Exit Sub
    lines = Split(Replace(ReadUtf8(f), vbCr, ""), vbLf)
    hdr = Split(lines(0), vbTab)
    Set ws = SheetByPrefix(wb, hdr(1))
    r0 = ws.Range(hdr(2)).Row
    c0 = ws.Range(hdr(2)).Column
    For i = 1 To UBound(lines)
        If Len(lines(i)) > 0 Then
            flds = Split(lines(i), vbTab)
            For j = 0 To UBound(flds)
                If Len(flds(j)) > 0 Then PutVal ws.Cells(r0 + i - 1, c0 + j), flds(j)
            Next j
        End If
    Next i
End Sub

Private Function LastRow(ws As Worksheet, ByVal col As String, ByVal r0 As Long) As Long
    Dim r As Long
    r = r0
    Do While Len(CStr(ws.Cells(r + 1, col).Value)) > 0
        r = r + 1
    Loop
    LastRow = r
End Function

Private Sub DelChart(ws As Worksheet, ByVal nm As String)
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        If co.Name = nm Then co.Delete
    Next co
End Sub

Private Sub BarChart(ws As Worksheet, ByVal nm As String, cats As Range, vals As Range, ByVal ttl As String, _
                     ByVal leftCell As String, ByVal barH As Boolean)
    Dim co As ChartObject
    DelChart ws, nm
    Set co = ws.ChartObjects.Add(ws.Range(leftCell).Left, ws.Range(leftCell).Top, 520, 300)
    co.Name = nm
    With co.Chart
        .ChartType = IIf(barH, 57, 51)   ' xlBarClustered / xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .XValues = cats
            .Values = vals
            .HasDataLabels = True
        End With
        .HasTitle = True
        .ChartTitle.Text = ttl
        .HasLegend = False
    End With
End Sub

' Совместимость с Excel 2016/2019: XLOOKUP(a,b,c) -> INDEX(c,MATCH(a,b,0)) (результат тот же)
Private Sub FixXlookup(wb As Workbook)
    Dim ws As Worksheet, c As Range, rx As Object, f As String
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "(_xlfn\.)?XLOOKUP\(([^,()]+),([^,()]+),([^,()]+)\)"
    For Each ws In wb.Worksheets
        For Each c In ws.UsedRange.Cells
            If c.HasFormula Then
                f = c.Formula
                If InStr(1, f, "XLOOKUP", vbTextCompare) > 0 Then c.Formula = rx.Replace(f, "INDEX($4,MATCH($2,$3,0))")
            End If
        Next c
    Next ws
End Sub

Public Sub Main()
    Dim wb As Workbook, p As String, ws As Worksheet, co As ChartObject, n As Long, i As Long
    Set wb = ActiveWorkbook
    p = DataDir(wb, "k_01_kompanii.tsv")
    If p = "" Then MsgBox "LR5: data folder VBA\data not found (k_01_kompanii.tsv)": Exit Sub
    LoadStrings p & "k_strings.tsv"
    Application.ScreenUpdating = False
    Application.Calculation = -4135   ' xlCalculationManual
    FixXlookup wb                          ' XLOOKUP есть только в Excel 2021/365
    LoadBlock wb, p & "k_01_kompanii.tsv"
    LoadBlock wb, p & "k_02_dokazatelstva.tsv"
    LoadBlock wb, p & "k_04_ocenka.tsv"
    LoadBlock wb, p & "k_04_fix.tsv"      ' исправление N: SUMPRODUCT строки и столбца весов давал #VALUE!
    LoadBlock wb, p & "k_07_karta.tsv"
    LoadBlock wb, p & "k_06_top10_fix.tsv"   ' исправление: топ-10 без дублей при равных баллах
    LoadBlock wb, p & "k_06_vyvod.tsv"
    LoadBlock wb, p & "k_vvod.tsv"
    Application.Calculation = -4105   ' xlCalculationAutomatic
    Application.CalculateFull

    ' Карта конкурентной близости: X = среднее(C1;C3), Y = среднее(C2;C7)
    Set ws = SheetByPrefix(wb, "07_")
    n = LastRow(SheetByPrefix(wb, "01_"), "B", 3) - 3
    DelChart ws, "NotaCode_Map"
    Set co = ws.ChartObjects.Add(ws.Range("M3").Left, ws.Range("M3").Top, 620, 440)
    co.Name = "NotaCode_Map"
    With co.Chart
        .ChartType = -4169   ' xlXYScatter
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .XValues = ws.Range("C4").Resize(n, 1)
            .Values = ws.Range("D4").Resize(n, 1)
            .MarkerSize = 8
            For i = 1 To n
                .Points(i).HasDataLabel = True
                .Points(i).DataLabel.Text = CStr(ws.Cells(3 + i, 2).Value)
                .Points(i).DataLabel.Font.Size = 8
            Next i
        End With
        .HasTitle = True
        .ChartTitle.Text = S("chart_title")
        .HasLegend = False
        With .Axes(1)
            .MinimumScale = 0: .MaximumScale = 5: .MajorUnit = 1
            .HasTitle = True: .AxisTitle.Text = S("x_title")
        End With
        With .Axes(2)
            .MinimumScale = 0: .MaximumScale = 5: .MajorUnit = 1
            .HasTitle = True: .AxisTitle.Text = S("y_title")
        End With
    End With
    ' Столбчатая диаграмма взвешенных баллов на 04_Оценка
    Set ws = SheetByPrefix(wb, "04_")
    BarChart ws, "NotaCode_Scores", ws.Range("B4").Resize(n, 1), ws.Range("N4").Resize(n, 1), "N (0-5)", "X3", True
    Application.ScreenUpdating = True
    MsgBox S("done")
End Sub
