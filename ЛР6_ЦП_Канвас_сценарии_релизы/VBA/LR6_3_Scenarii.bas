Attribute VB_Name = "LR6_3_Scenarii"
' =====================================================================
' ЛР6 «Электронный бизнес», NotaCode. Методика 6.3 Бизнес-логика и сценарии.
' Выполнила: Хаджинова К.
' Назначение: заполнить Excel-шаблон «Шаблон_бизнес_логики_и_пользовательских_сценариев.xlsx»
'   данными NotaCode из файла VBA\data\lr6_3.tsv (UTF-8, выгружен скриптом Python\lr6_calc.py),
'   не трогая зелёные ячейки с формулами (в TSV помечены ~), пересчитать книгу и построить диаграмму.
' Инструкция:
'   1) Скопируйте шаблон из папки Материалы (оригинал не изменяйте) в папку VBA\data и откройте копию.
'   2) Файл lr6_3.tsv ищется рядом с копией, в подпапке data; иначе макрос откроет диалог выбора.
'   3) Alt+F11 -> File -> Import File -> LR6_3_Scenarii.bas.
'   4) Alt+F8 -> Main -> Выполнить. Сохраните книгу (.xlsx: модуль не сохранится, данные — да).
' Код содержит только ASCII: кириллица берётся из TSV (UTF-8) и из \uXXXX-последовательностей,
' поэтому макрос работает при любой локали Windows (в т.ч. en-US). Русские комментарии в cp1251.
' =====================================================================
Option Explicit

Public Sub Main()
    RunLoad "lr6_3.tsv"
    FixScenarioValidation
    AddBarChart ActiveWorkbook.Worksheets(U("\u0414\u0430\u0448\u0431\u043E\u0440\u0434")), ActiveWorkbook.Worksheets(U("\u041F\u043E\u043B\u044C\u0437\u043E\u0432\u0430\u0442\u0435\u043B\u044C\u0441\u043A\u0438\u0435 \u0441\u0446\u0435\u043D\u0430\u0440\u0438\u0438")), "A4:A16", "R4:R16", U("\u0418\u043D\u0434\u0435\u043A\u0441 \u0441\u0446\u0435\u043D\u0430\u0440\u0438\u044F"), 20, 360
    MsgBox "LR6 LR6_3_Scenarii: OK", vbInformation
End Sub

Private Sub FixScenarioValidation()
    ' Исправление шаблона: на Q4:Q203 наложены две проверки (1-5 и 1-4). Оставляем N:P = 1-5, Q = 1-4
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(U("\u041F\u043E\u043B\u044C\u0437\u043E\u0432\u0430\u0442\u0435\u043B\u044C\u0441\u043A\u0438\u0435 \u0441\u0446\u0435\u043D\u0430\u0440\u0438\u0438"))
    ws.Range("N4:Q203").Validation.Delete
    ws.Range("N4:P203").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="1,2,3,4,5"
    ws.Range("Q4:Q203").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="1,2,3,4"
End Sub

' ---------- общий загрузчик данных (одинаковый во всех модулях ЛР6) ----------
Private Function U(ByVal s As String) As String
    ' Раскодирует \uXXXX в символы Unicode
    Dim i As Long, r As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" Then
            r = r & ChrW(CLng("&H" & Mid$(s, i + 2, 4)))
            i = i + 6
        Else
            r = r & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = r
End Function

Private Function ReadUtf8(ByVal path As String) As String
    ' Чтение текстового файла UTF-8 через ADODB.Stream
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

Private Function PickTsv(ByVal wb As Workbook, ByVal fileName As String) As String
    ' Ищет TSV рядом с книгой и в подпапках data, VBA\data; иначе — диалог выбора файла
    Dim fso As Object, cand As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each cand In Array(wb.Path & "\" & fileName, wb.Path & "\data\" & fileName, wb.Path & "\VBA\data\" & fileName)
        If fso.FileExists(cand) Then
            PickTsv = cand
            Exit Function
        End If
    Next cand
    With Application.FileDialog(3)
        .Title = "Select " & fileName & " (LR6\VBA\data)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "TSV", "*.tsv"
        If .Show = -1 Then PickTsv = .SelectedItems(1)
    End With
End Function

Private Function ColNum(ByVal letters As String) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(letters)
        n = n * 26 + (Asc(UCase$(Mid$(letters, i, 1))) - 64)
    Next i
    ColNum = n
End Function

Private Function IsPlainNumber(ByVal v As String) As Boolean
    ' Число в формате 123 или 3.25 (точка — десятичный разделитель в TSV)
    Dim i As Long, ch As String, dots As Long
    If Len(v) = 0 Or Len(v) > 15 Then Exit Function
    For i = 1 To Len(v)
        ch = Mid$(v, i, 1)
        If ch = "." Then
            dots = dots + 1
        ElseIf ch = "-" And i = 1 And Len(v) > 1 Then
            ' знак минуса
        ElseIf ch < "0" Or ch > "9" Then
            Exit Function
        End If
    Next i
    IsPlainNumber = (dots <= 1)
End Function

Private Function LoadTsv(ByVal wb As Workbook, ByVal path As String) As Long
    ' Формат: #SHEET<TAB>лист | R<TAB>строка<TAB>столбец<TAB>значения...
    ' "~" — ячейка с формулой шаблона (не трогать), пусто — очистить ячейку
    Dim txt As String, lines() As String, f() As String, i As Long, j As Long
    Dim ws As Worksheet, r As Long, c As Long, v As String, cnt As Long
    txt = ReadUtf8(path)
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(lines(i)) > 0 Then
            f = Split(lines(i), vbTab)
            If f(0) = "#SHEET" Then
                Set ws = Nothing
                On Error Resume Next
                Set ws = wb.Worksheets(f(1))
                On Error GoTo 0
                If ws Is Nothing Then MsgBox "Sheet not found: " & f(1), vbExclamation
            ElseIf f(0) = "R" And Not ws Is Nothing Then
                r = CLng(f(1))
                c = ColNum(f(2))
                For j = 3 To UBound(f)
                    v = f(j)
                    If v = "~" Then
                        ' формула шаблона
                    ElseIf v = "" Then
                        ws.Cells(r, c + j - 3).ClearContents
                    ElseIf IsPlainNumber(v) Then
                        ws.Cells(r, c + j - 3).Value = Val(v)
                    Else
                        ws.Cells(r, c + j - 3).Value = v
                    End If
                Next j
                cnt = cnt + 1
            End If
        End If
    Next i
    LoadTsv = cnt
End Function

Private Sub AddBarChart(ByVal target As Worksheet, ByVal src As Worksheet, ByVal labels As String, ByVal values As String, ByVal title As String, ByVal leftPos As Double, ByVal topPos As Double)
    ' Линейчатая диаграмма по диапазонам листа src; повторный запуск пересоздаёт её
    Dim co As ChartObject, nm As String
    nm = "LR6_" & src.Index & "_" & Replace(values, ":", "_")
    On Error Resume Next
    target.ChartObjects(nm).Delete
    On Error GoTo 0
    Set co = target.ChartObjects.Add(leftPos, topPos, 520, 320)
    co.Name = nm
    With co.Chart
        .ChartType = xlBarClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Values = src.Range(values)
            .XValues = src.Range(labels)
        End With
        .HasTitle = True
        .ChartTitle.Text = title
        .HasLegend = False
    End With
End Sub

Private Sub RunLoad(ByVal fileName As String)
    ' Загрузка данных NotaCode в активную книгу-копию шаблона и полный пересчёт
    Dim p As String, n As Long
    p = PickTsv(ActiveWorkbook, fileName)
    If p = "" Then End
    Application.ScreenUpdating = False
    n = LoadTsv(ActiveWorkbook, p)
    Application.CalculateFull
    Application.ScreenUpdating = True
    Debug.Print fileName & ": " & n & " rows"
End Sub
