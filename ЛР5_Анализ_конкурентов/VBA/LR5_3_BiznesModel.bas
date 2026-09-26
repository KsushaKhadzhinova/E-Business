Attribute VB_Name = "LR5_3_BiznesModel"
Option Explicit
' =====================================================================================
' ЛР5 «Анализ конкурентов» — NotaCode. Модуль 3: бизнес-модель конкурентов
' Шаблон: Материалы\Шаблон_анализа_бизнес_модели_конкурентов.xlsx
'
' КАК ЗАПУСТИТЬ
'   1. Скопируйте шаблон в папку ЛР5_Анализ_конкурентов и сохраните копию как .xlsm
'      (например, ЛР5_3_Бизнес_модель_NotaCode.xlsm). Рядом должна лежать папка VBA\data (TSV-файлы).
'   2. Alt+F11 -> File -> Import File... -> VBA\LR5_3_BiznesModel.bas
'   3. Alt+F8 -> LR5_3_BiznesModel.Main -> Выполнить.
'   Excel 2016+ (модуль 1 заменяет XLOOKUP на INDEX/MATCH), ADODB.Stream (есть в Windows).
'
' ПОЧЕМУ ДАННЫЕ В TSV. Редактор VBA импортирует .bas в ANSI-кодировке, кириллица в строках
' кода ломается. Поэтому код — только ASCII, а все тексты NotaCode читаются из UTF-8
' файлов VBA\data\b_*.tsv (их создаёт Python\export_vba_data.py из Python\lr5_data.py).
' Если комментарии в редакторе отображаются «кракозябрами» — это косметика, код работает.
'
' ЧТО ДЕЛАЕТ МАКРОС
'   02_Конкуренты (8), 03_Факты_сайта (29), 04-08 профили, 09_Матрица: ВЕСА строки 2
'   (10/15/15/10/5/15/5/5/5/10/5) и баллы 1-5; 10_Стандарт (13 практик), 11_Преимущества (8),
'   12_Сводка C23:C27 (выводы), лист 00_Паспорт_NotaCode; ИСПРАВЛЯЕТ объединение 09_Матрица!A2:Q2 и 10_Стандарт E4:E203
'   (делитель '12_Сводка'!B5); диаграмма итогов 09_Матрица!Q.
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

Public Sub Main()
    Dim wb As Workbook, p As String, ws As Worksheet, f As Variant, n As Long
    Set wb = ActiveWorkbook
    p = DataDir(wb, "b_02_konkurenty.tsv")
    If p = "" Then MsgBox "LR5: data folder VBA\data not found (b_02_konkurenty.tsv)": Exit Sub
    LoadStrings p & "b_strings.tsv"
    Application.ScreenUpdating = False
    Application.Calculation = -4135
    ' строка весов 09_Матрица!A2:Q2 в шаблоне объединена -> веса в C2:M2 некуда записать
    Set ws = SheetByPrefix(wb, "09_")
    ws.Range("A2:Q2").UnMerge
    For Each f In Array("b_02_konkurenty", "b_03_fakty", "b_04_profil", "b_05_tovar", "b_06_monet", "b_07_kanaly", _
                        "b_08_ops", "b_09_ves", "b_09_matrica", "b_10_standart", "b_10_fix", "b_11_preim", _
                        "b_12_vyvod", "b_00_pasport", "b_vvod")
        LoadBlock wb, p & f & ".tsv"
    Next f
    ' b_09_ves: веса строки 2 листа 09_Матрица (в шаблоне пусто -> деление на 0)
    ' b_10_fix: доля в 10_Стандарт делилась на '12_Сводка'!B4 (текст «Значение») -> заменено на B5
    Application.Calculation = -4105
    Application.CalculateFull
    Set ws = SheetByPrefix(wb, "09_")
    n = LastRow(ws, "A", 3) - 3
    BarChart ws, "NotaCode_BM", ws.Range("B4").Resize(n, 1), ws.Range("Q4").Resize(n, 1), S("chart_title"), "S3", True
    Application.ScreenUpdating = True
    MsgBox S("done")
End Sub
