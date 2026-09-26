Attribute VB_Name = "NC6_ScenarSW"
' NOTE (ASCII): comments are in Russian (UTF-8). Read them in VS Code / GitHub if VBE shows garbage.
' String literals are ASCII transliteration decoded by R() - works on any Windows ANSI code page.
'
' =====================================================================================
' ЛР2-3 «Оценка объёма рынка», дисциплина «Электронный бизнес». Проект NotaCode.
' Выполнила: Хаджинова К. Модуль NC6_ScenarSW: сценарная оценка рынка с данными Similarweb + сверка всех методов
' Шаблон: Материалы\shablon_excel_scenarnaya_ocenka_rynka_s_similarweb.xlsx
'
' Как запустить:
'   1) скопируйте шаблон из папки Материалы (оригинал не трогайте), откройте копию в Excel;
'   2) Alt+F11 -> File -> Import File... -> выберите этот файл NC6_ScenarSW.bas;
'   3) Alt+F8 -> NC6_ScenarSW.Main -> Выполнить (или F5 внутри процедуры Main);
'   4) сохраните книгу как .xlsm (если нужен макрос) или .xlsx (только результат).
'
' Что делает макрос:
'   1) 01_Параметры: чек 120 BYN; охват неучтённых конкурентов = 1/покрытие (1,25 / 1,67 / 2,22);
'      конверсии визит->регистрация и Free->Pro; множитель чека 0,6/1/1,2 (72/120/144 BYN); доля 1/3/7 %;
'   2) 02_Similarweb: реальные данные 4 доменов (демо-строка 5 удаляется);
'   3) 05_Платформа обнуляется (NotaCode - не платформа), в 06_Сверка строка платформы = «н/п»;
'   4) 06_Сверка: строки 7-9 заменяются результатами других методов NotaCode (SOM, BYN/год),
'      «Рекомендуемый диапазон» - формула с TEXT(), чтобы не было 15 знаков после запятой;
'   5) лист NotaCode_Итог с диаграммой сверки (лог. шкала).
'
' Пометки в книге: оранжевая заливка + примечание «ДОСНЯТЬ» - данные снять самой;
' примечание «(допущение)» - экспертное допущение, его можно заменить своим значением.
' Курс пересчёта (допущение): 1 USD = 3,00 BYN на 25.09.2026 (уточнить на nbrb.by).
' Pro = 4 USD/мес = 12 BYN/мес; 40 USD/год = 120 BYN/год.
' Все числа совпадают с Python\nc_model.py (единая модель расчёта).
' =====================================================================================
Option Explicit

Public Sub Main()
    NC_Run True
End Sub

Public Sub MainNoMsg()
    NC_Run False
End Sub

Private Sub NC_Run(ByVal showMsg As Boolean)
    Dim p As Worksheet, w As Worksheet, pl As Worksheet, v As Worksheet, src As Worksheet, res As Worksheet
    Dim sw As Variant, i As Long, col As Variant
    Set p = SH("01_"): Set w = SH("02_"): Set pl = SH("05_"): Set v = SH("06_"): Set src = SH("08_")

    p.Range("B4").Value = R("{NotaCode} - {web}-{IDE} {diagram} {as} {code} dlya formal'nykh notatsij ({Pro})")
    p.Range("B5").Value = R("Russkoyazychnyj segment (RF + RB) po trafiku {Similarweb}")
    p.Range("B6").Value = "BYN": p.Range("B7").Value = 12
    p.Range("B8").Value = 120: p.Range("F8").Value = R("godovoj tarif {Pro} 40 {USD} {x} 3,00 {BYN} (kurs - dopushchenie)")
    p.Range("B9").Value = 1: p.Range("F9").Value = R("godovoj chek uzhe vklyuchaet 12 mes. podpiski")
    p.Range("B10").Value = 0.0907: p.Range("B11").Value = 0.35: p.Range("B12").Value = 0.75
    p.Range("B13").Value = DateSerial(2026, 9, 25)
    PutRow p, "B16", Array(1.25, 1 / 0.6, 1 / 0.45)
    PutRow p, "B19", Array(0.02, 0.04, 0.06)
    PutRow p, "B20", Array(0.02, 0.03, 0.05)
    PutRow p, "B21", Array(0.6, 1, 1.2)
    PutRow p, "B22", Array(1, 1, 1)
    PutRow p, "B23", Array(0.01, 0.03, 0.07)
    p.Range("A19").Value = R("Konversiya vizit {\u2192} registratsiya {Free}")
    p.Range("A20").Value = R("Konversiya {Free} {\u2192} {Pro}")
    Assume p.Range("B16:D16"), R("= 1 / koe'ffitsient okhvata vyborki (0,80 / 0,60 / 0,45)")
    Assume p.Range("B19:D23"), R("{SaaS}-orientiry; {KPI}-05 = 3 %; chek 72/120/144 {BYN}; dolya 1/3/7 % (K5)")

    sw = Array( _
        Array("PlantUML", "plantuml.com", R("pryamoj konkurent"), 507000, 0.0907, 0.35, 0.8, R("Rossiya 9,07 % (real'nye); 17.09.2026")), _
        Array("Eraser", "eraser.io", R("pryamoj konkurent ({AI})"), 678500, 0.01, 0.175, 0.8, R("RF/RB vne top-5 - 1 % (dopushchenie); 23.09.2026")), _
        Array("draw.io", "app.diagrams.net", R("zamenitel' ({GUI})"), 7800000, 0.01, 0.07, 0.7, R("RF/RB vne top-5 - 1 % (dopushchenie); 23.09.2026")), _
        Array("Mermaid Chart", "mermaidchart.com", R("pryamoj konkurent"), 20000, 0.01, 0.245, 0.7, R("<20 tys. vizitov; 17.09.2026")))
    For i = 0 To 3
        w.Cells(5 + i, 2).Value = sw(i)(0): w.Cells(5 + i, 3).Value = sw(i)(1): w.Cells(5 + i, 4).Value = sw(i)(2)
        w.Cells(5 + i, 5).Value = sw(i)(3): w.Cells(5 + i, 6).Value = sw(i)(4): w.Cells(5 + i, 7).Value = sw(i)(5)
        w.Cells(5 + i, 8).Value = sw(i)(6)
        w.Cells(5 + i, 12).Value = "https://www.similarweb.com/website/" & sw(i)(1) & "/"
        w.Cells(5 + i, 13).Value = sw(i)(7)
    Next i
    w.Range("B9:H9").ClearContents: w.Range("L9:M9").ClearContents
    Todo w.Range("B9"), R("eshchyo 1-11 domenov (metodika: 5-15 konkurentov): {mermaid}.{live}, {dbdiagram}.{io}, {lucidchart}.{com} ...")
    Assume w.Range("G5:G8"), R("dolya kategorii ""formal'nye notatsii"" (1 / 0,5 / 0,2 / 0,7) {x} kommercheskaya dolya 0,35")

    ' --- платформа не применима
    For Each col In Array("C", "E", "F", "H", "J")
        pl.Range(col & "5:" & col & "7").Value = 0
    Next col
    pl.Range("A9").Value = R("Ne primenimo: {NotaCode} - podpiska {SaaS}, a ne platforma sdelok; {GMV} ne schitaet|sya.")

    ' --- сверка методов (SOM, BYN/год)
    v.Range("B6:D6").Value = R("n/p")
    v.Range("E6").Value = R("Ne primenimo ({NotaCode} - ne platforma); tekst ne uchastvuet v {MIN}/{AVERAGE}/{MAX}")
    v.Range("A7").Value = R("Tsifrovaya voronka (kanaly {NotaCode})")
    PutRow v, "B7", Array(441, 2943, 13244)
    v.Range("E7").Value = R("{shablon}_..._{voronku}.{xlsx}, 06_Povtory_{LTV}!{E}4:{E}6 (modul' {NC}2). Poiskovyj spros - posle Vordstat")
    v.Range("A8").Value = R("Kolichestvo potentsial'nykh klientov ({SOM})")
    PutRow v, "B8", Array(146, 1516, 7410)
    v.Range("E8").Value = R("..._{potencialnyh}_{klientov}.{xlsx}, 03_Stsenarii!{J}4:{J}6 (modul' {NC}3)")
    v.Range("A9").Value = R("Sverkhu vniz ot VVP ({SOM} = {SAM} {x} dolya)")
    PutRow v, "B9", Array(2157, 25888, 158561)
    v.Range("E9").Value = R("..._{sverhu}_{vniz}_{snizu}_{vverh}....{xlsx}, 02_Sverkhu_vniz {x} 1/3/7 % (modul' {NC}5)")
    For Each col In Array("B", "C", "D")
        v.Range(col & "15").Formula = "=TEXT(" & col & "12,""0"")&"" " & ChrW(8211) & " ""&TEXT(" & col & "14,""0"")"
    Next col
    Note v.Range("B15"), R("IZMENENO: v shablone ={B}12&... bez okrugleniya - vyvodilis' lishnie znaki")
    Note v.Range("B7"), R("znacheniya iz drugikh shablonov; pri izmenenii dopushchenij obnovit' vruchnuyu (sm. {Python}\{nc}_{model}.{py})")

    src.Range("E4").Value = DateSerial(2026, 9, 23)
    src.Range("C7").Value = "https://www.eraser.io/pricing": src.Range("E7").Value = DateSerial(2026, 9, 17)
    src.Range("C8").Value = "https://trends.google.com/": src.Range("E8").Value = DateSerial(2026, 9, 17)
    src.Range("G8").Value = R("{GT} snyat ({PlantUML}, {text} {to} {diagram}); Vordstat - DOSNYaT'")

    Application.CalculateFull
    Set res = ResultSheet()
    PutRow res, "A3", Array(R("Metod ({SOM}, {BYN}/god)"), R("Ostorozhnyj"), R("Bazovyj"), R("Optimistichnyj"))
    Dim rr As Variant, n As Long
    n = 4
    For Each rr In Array(5, 7, 8, 9, 12, 13, 14)
        res.Cells(n, 1).Formula = "=" & Q(v) & "A" & rr
        res.Cells(n, 2).Formula = "=" & Q(v) & "B" & rr
        res.Cells(n, 3).Formula = "=" & Q(v) & "C" & rr
        res.Cells(n, 4).Formula = "=" & Q(v) & "D" & rr
        n = n + 1
    Next rr
    res.Range("A12").Value = R("Vyruchka rynka ({SAM}) po {SW}-voronke")
    res.Range("B12").Formula = "=" & Q(SH("04_")) & "H5"
    res.Range("C12").Formula = "=" & Q(SH("04_")) & "H6"
    res.Range("D12").Formula = "=" & Q(SH("04_")) & "H7"
    res.Range("B4:D12").NumberFormat = "# ##0"
    res.Columns("A:D").AutoFit
    AddChart res, res.Range("A3:D7"), R("Sverka metodov: dostizhimaya vyruchka {NotaCode}, {BYN}/god (log. shkala)"), "F3", True
    If showMsg Then MsgBox R("Gotovo. Rekomenduemyj diapazon (baza): ") & v.Range("C15").Value & R(" {BYN}/god"), vbInformation, "NotaCode"
End Sub

' ============================== служебные функции ==============================
' R() переводит транслит в кириллицу (см. шапку). {...} - текст как есть, \uXXXX - символ Unicode,
' | - разделитель, если буквы нельзя склеивать (ot|sut|stvie = «отсутствие»).
Private Function R(ByVal s As String) As String
    Static keys As Variant, vals As Variant
    Dim i As Long, n As Long, j As Long, idx As Long, code As Long
    Dim res As String, ch As String, low As String, found As Boolean, up As Boolean
    If IsEmpty(keys) Then
        keys = Array("shch", "zh", "kh", "ts", "ch", "sh", "yo", "yu", "ya", "e'", "''", _
                     "a", "b", "v", "g", "d", "e", "z", "i", "j", "k", "l", "m", "n", "o", "p", _
                     "r", "s", "t", "u", "f", "y", "'")
        vals = Array(1097, 1078, 1093, 1094, 1095, 1096, 1105, 1102, 1103, 1101, 1098, _
                     1072, 1073, 1074, 1075, 1076, 1077, 1079, 1080, 1081, 1082, 1083, 1084, 1085, 1086, 1087, _
                     1088, 1089, 1090, 1091, 1092, 1099, 1100)
    End If
    i = 1: n = Len(s)
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "{" Then
            j = InStr(i, s, "}")
            res = res & DecodeU(Mid$(s, i + 1, j - i - 1))
            i = j + 1
        ElseIf ch = "|" Then
            i = i + 1
        Else
            low = LCase$(Mid$(s, i, 4))
            found = False
            For idx = 0 To UBound(keys)
                If Left$(low, Len(keys(idx))) = keys(idx) Then
                    code = vals(idx)
                    up = (ch <> LCase$(ch))
                    If Not up And (keys(idx) = "'" Or keys(idx) = "''") And Len(res) >= 2 Then
                        up = IsUpCyr(Right$(res, 1)) And IsUpCyr(Mid$(res, Len(res) - 1, 1))
                    End If
                    If up Then
                        If code = 1105 Then code = 1025 Else code = code - 32
                    End If
                    res = res & ChrW(code)
                    i = i + Len(keys(idx))
                    found = True
                    Exit For
                End If
            Next idx
            If Not found Then
                res = res & ch
                i = i + 1
            End If
        End If
    Loop
    R = res
End Function

Private Function DecodeU(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW(CLng("&H" & Mid$(s, p + 2, 4))) & Mid$(s, p + 6)
        p = InStr(1, s, "\u")
    Loop
    DecodeU = s
End Function

Private Function IsUpCyr(ByVal c As String) As Boolean
    Dim a As Long
    If Len(c) = 0 Then Exit Function
    a = AscW(c)
    IsUpCyr = (a >= 1040 And a <= 1071) Or a = 1025
End Function

' Лист по префиксу имени ("01_", "02_" ...) - так код не зависит от кодировки имён листов.
Private Function SH(ByVal prefix As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If Left$(ws.Name, Len(prefix)) = prefix Then
            Set SH = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 513, , "Sheet with prefix " & prefix & " not found. Open a copy of the right template."
End Function

Private Function Q(ByVal ws As Worksheet) As String
    Q = "'" & ws.Name & "'!"
End Function

' Записать массив значений в строку, начиная с ячейки addr.
Private Sub PutRow(ByVal ws As Worksheet, ByVal addr As String, ByVal v As Variant)
    Dim k As Long
    For k = LBound(v) To UBound(v)
        ws.Range(addr).Offset(0, k - LBound(v)).Value = v(k)
    Next k
End Sub

Private Sub Note(ByVal rng As Range, ByVal txt As String)
    Dim c As Range
    For Each c In rng.Cells
        If Not c.Comment Is Nothing Then c.Comment.Delete
        c.AddComment txt
        c.Comment.Shape.Width = 220
        c.Comment.Shape.Height = 70
    Next c
End Sub

' Пометить ячейки, которые студентка заполняет сама (оранжевая заливка + примечание).
Private Sub Todo(ByVal rng As Range, ByVal txt As String)
    rng.Interior.Color = RGB(255, 199, 128)
    Note rng, R("DOSNYaT': ") & txt
End Sub

Private Sub Assume(ByVal rng As Range, ByVal txt As String)
    Note rng, R("(dopushchenie) ") & txt
End Sub

' Лист итогов NotaCode: пересоздаётся при каждом запуске.
Private Function ResultSheet() As Worksheet
    Dim ws As Worksheet, nm As String
    nm = R("{NotaCode}_Itog")
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = nm
    ws.Range("A1").Value = R("Itogi raschyota dlya {NotaCode} (LR2-3, Khadzhinova K.)")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    Set ResultSheet = ws
End Function

' Кластерная гистограмма по диапазону (первый столбец - подписи, первая строка - серии).
Private Sub AddChart(ByVal ws As Worksheet, ByVal src As Range, ByVal title As String, _
                     ByVal topCell As String, Optional ByVal logScale As Boolean = False)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(ws.Range(topCell).Left, ws.Range(topCell).Top, 520, 280)
    co.Chart.ChartType = xlColumnClustered
    co.Chart.SetSourceData Source:=src
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = title
    If logScale Then
        co.Chart.Axes(xlValue).ScaleType = xlScaleLogarithmic
    End If
End Sub

Private Sub Scen3(ByVal ws As Worksheet, ByVal row As Long)
    PutRow ws, "B" & row, Array(R("Ostorozhnyj"), R("Bazovyj"), R("Optimistichnyj"))
End Sub
