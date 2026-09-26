Attribute VB_Name = "NC4_Platezhesposobnost"
' NOTE (ASCII): comments are in Russian (UTF-8). Read them in VS Code / GitHub if VBE shows garbage.
' String literals are ASCII transliteration decoded by R() - works on any Windows ANSI code page.
'
' =====================================================================================
' ЛР2-3 «Оценка объёма рынка», дисциплина «Электронный бизнес». Проект NotaCode.
' Выполнила: Хаджинова К. Модуль NC4_Platezhesposobnost: проверка рынка через платёжеспособность
' Шаблон: Материалы\shablon_excel_metodika_proverki_rynka_cherez_platezhesposobnost.xlsx
'
' Как запустить:
'   1) скопируйте шаблон из папки Материалы (оригинал не трогайте), откройте копию в Excel;
'   2) Alt+F11 -> File -> Import File... -> выберите этот файл NC4_Platezhesposobnost.bas;
'   3) Alt+F8 -> NC4_Platezhesposobnost.Main -> Выполнить (или F5 внутри процедуры Main);
'   4) сохраните книгу как .xlsm (если нужен макрос) или .xlsx (только результат).
'
' Что делает макрос:
'   1) заполняет 01_Параметры, 02_Сегменты (те же S-01...S-04, что и в методе потенциальных клиентов);
'   2) 03_Ценовой_коридор: реальные цены Eraser (снято 17.09.2026) и цена NotaCode Pro 12 BYN/мес;
'      ячейка «Цена в модели» заменена ценой NotaCode (в шаблоне там медиана рынка);
'   3) 04_Экономика_покупки: экономия времени на построении диаграмм vs цена Pro (допущения);
'   4) 05_Платежеспособность: бюджет/мес. vs 12 BYN; 06_Сценарии: достижимая доля 1/3/7 %;
'   5) 07_Similarweb: реальные данные 4 доменов; удаляет лишние демо-строки; строит NotaCode_Итог.
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
    Dim p As Worksheet, s As Worksheet, k As Worksheet, e As Worksheet, b As Worksheet, sc As Worksheet
    Dim w As Worksheet, src As Worksheet, res As Worksheet
    Dim seg As Variant, ec As Variant, sw As Variant, i As Long
    Set p = SH("01_"): Set s = SH("02_"): Set k = SH("03_"): Set e = SH("04_"): Set b = SH("05_")
    Set sc = SH("06_"): Set w = SH("07_"): Set src = SH("09_")

    p.Range("B5").Value = R("{NotaCode} {Pro} - podpiska na {web}-{IDE} {diagram} {as} {code} dlya formal'nykh notatsij")
    p.Range("B6").Value = R("Respublika Belarus'")
    p.Range("B7").Value = "BYN": p.Range("B8").Value = R("god")
    p.Range("B9").Value = 0.0012: p.Range("B10").Value = 1: p.Range("B11").Value = 0.75
    p.Range("B12").Value = DateSerial(2026, 9, 25)
    Assume p.Range("B9"), R("vizit{\u2192}registratsiya 4 % {x} {Free\u2192Pro} 3 %")
    Assume p.Range("B11"), R("(12 - 3) / 12: sebestoimost' platyashchego pol'zovatelya 3 {BYN}/mes (khosting, {AI}-kvota)")

    seg = Array( _
        Array(R("{S}-01 Studenty IT-spetsial'nostej"), 19600, 0.8, 0.9, 0.03, 72, R("Uchebnye raboty s {UML}/{IDEF}/{DFD}/{ERD}"), R("{belarus}.{by} + dopushcheniya")), _
        Array(R("{S}-02 Analitiki"), 1600, 0.6, 0.9, 0.1, 120, R("{BPMN}, {IDEF}0, {DFD}, {ERD}"), R("PVT + dopushcheniya")), _
        Array(R("{S}-03 Arkhitektory i razrabotchiki"), 1800, 0.5, 0.9, 0.08, 120, R("{UML}, {ERD}, e'ksport {DDL}"), R("PVT + dopushcheniya")), _
        Array(R("{S}-04 Prepodavateli"), 252, 0.7, 0.9, 0.1, 120, R("Primery i shablony dlya zanyatij"), R("{belarus}.{by} + dopushcheniya")))
    For i = 0 To 3
        s.Cells(5 + i, 1).Value = seg(i)(0): s.Cells(5 + i, 2).Value = seg(i)(1): s.Cells(5 + i, 3).Value = seg(i)(2)
        s.Cells(5 + i, 4).Value = seg(i)(3): s.Cells(5 + i, 5).Value = seg(i)(4): s.Cells(5 + i, 6).Value = 1
        s.Cells(5 + i, 7).Value = seg(i)(5): s.Cells(5 + i, 11).Value = seg(i)(6): s.Cells(5 + i, 12).Value = seg(i)(7)
    Next i
    s.Range("A9:G9").ClearContents: s.Range("K9:L9").ClearContents
    Note s.Range("B5"), R("klientov v granitsakh rynka = 28 000 {x} 0,7 (sm. modul' {NC}3)")
    Assume s.Range("E5:E8"), R("dolya gotovykh platit' {Pro} (studenty - 3 % po {KPI}-05)")

    ' --- ценовой коридор, BYN/мес на пользователя
    k.Range("A5").Value = R("{Pro} pomesyachno (na pol'zovatelya)")
    k.Range("B5").Value = 12: k.Range("C5").Value = 52.5: k.Range("D5").Value = 180: k.Range("E5").Value = 3
    k.Range("G5").Value = 12
    k.Range("I5").Value = R("min. = {NotaCode} 4 {USD}; mediana = {Eraser} {Starter} (15+20)/2 {USD}; maks. = {Eraser} {Business} 60 {USD}")
    k.Range("J5").Value = R("{eraser}.{io}/{pricing}, 17.09.2026; kurs 3,00")
    k.Range("A6").Value = R("{Pro} pri godovoj oplate (v pereschyote na mesyats)")
    k.Range("B6").Value = 10: k.Range("C6").Value = 45: k.Range("D6").Value = 135: k.Range("E6").Value = 3
    k.Range("G6").Value = 10
    k.Range("I6").Value = R("min. = 40 {USD}/12; mediana = {Eraser} {Starter} 15 {USD}; maks. = {Eraser} {Business} 45 {USD} (godovaya oplata)")
    k.Range("J6").Value = R("{eraser}.{io}/{pricing}, 17.09.2026; kurs 3,00")
    k.Range("A7:E9").ClearContents: k.Range("I7:J9").ClearContents
    k.Range("A11").Value = R("Besplatnye zameniteli ({PlantUML}, {Mermaid}, {draw}.{io}) - 0 {BYN}: v koridor ne vklyucheny (delenie na 0), uchteny v vyvode.")
    Note k.Range("G5"), R("ZAMENENO: v shablone zdes' ={C} (mediana rynka); dlya proverki podstavlena tsena {NotaCode} {Pro}")
    Note k.Range("G6"), R("ZAMENENO: tsena {NotaCode} {Pro} pri oplate za god (40 {USD} / 12 {x} 3,00)")
    Assume k.Range("E5:E6"), R("sebestoimost' platyashchego pol'zovatelya 3 {BYN}/mes")

    ' --- экономика покупки: B = стоимость времени на диаграммы в мес., C = 1, D = доля экономии времени
    e.Range("B4").Value = R("Stoimost' vremeni na diagrammy, {BYN}/mes")
    e.Range("C4").Value = R("Dolya zatrat (1 = vsyo vremya - zatraty)")
    e.Range("D4").Value = R("Dolya e'konomii vremeni s {NotaCode}")
    e.Range("E4").Value = R("E'konomiya, {BYN}/mes")
    ec = Array( _
        Array(R("{S}-01 Studenty"), 100, 0.4, R("20 ch/mes {x} 5 {BYN}/ch (al'ternativnaya stoimost' vremeni)")), _
        Array(R("{S}-02 Analitiki"), 250, 0.4, R("10 ch/mes {x} 25 {BYN}/ch")), _
        Array(R("{S}-03 Arkhitektory i razrabotchiki"), 280, 0.3, R("8 ch/mes {x} 35 {BYN}/ch")), _
        Array(R("{S}-04 Prepodavateli"), 90, 0.3, R("6 ch/mes {x} 15 {BYN}/ch")))
    For i = 0 To 3
        e.Cells(5 + i, 1).Value = ec(i)(0): e.Cells(5 + i, 2).Value = ec(i)(1): e.Cells(5 + i, 3).Value = 1
        e.Cells(5 + i, 4).Value = ec(i)(2): e.Cells(5 + i, 6).Value = 12
        e.Cells(5 + i, 10).Value = ec(i)(3): e.Cells(5 + i, 11).Value = R("dopushchenie; proverit' oprosom")
    Next i
    e.Range("A9:D9").ClearContents: e.Range("F9").ClearContents: e.Range("J9:K9").ClearContents
    Assume e.Range("B5:D8"), R("chasy na diagrammy i stavka - dopushcheniya; 40 % e'konomii - tsel' {PR}-01 ({NotaCode} {vs} {draw}.{io}), ne izmereno")

    ' --- платёжеспособность
    ec = Array(Array(R("{S}-01 Studenty"), 15, 0.03, 72, R("lichnyj byudzhet na PO, dopushchenie")), _
               Array(R("{S}-02 Analitiki"), 60, 0.1, 120, R("byudzhet rabotodatelya na instrumenty, dopushchenie")), _
               Array(R("{S}-03 Arkhitektory i razrabotchiki"), 60, 0.08, 120, R("byudzhet rabotodatelya na instrumenty, dopushchenie")), _
               Array(R("{S}-04 Prepodavateli"), 20, 0.1, 120, R("lichnyj/kafedral'nyj byudzhet, dopushchenie")))
    For i = 0 To 3
        b.Cells(5 + i, 1).Value = ec(i)(0): b.Cells(5 + i, 2).Value = ec(i)(1): b.Cells(5 + i, 3).Value = 12
        b.Cells(5 + i, 6).Value = ec(i)(2): b.Cells(5 + i, 8).Value = ec(i)(3): b.Cells(5 + i, 10).Value = ec(i)(4)
    Next i
    b.Range("A9:C9").ClearContents: b.Range("F9").ClearContents: b.Range("H9").ClearContents: b.Range("J9").ClearContents
    Todo b.Range("B5"), R("proverit' po Belstatu (dokhody, stipendiya) ili oprosu studentov")
    Assume b.Range("B6:B8"), R("dostupnyj byudzhet na PO v mesyats; proverit' oprosom ({BO}-04: >= 30 respondentov)")

    PutRow sc, "B9", Array(0.01, 0.03, 0.07)
    ' ИСПРАВЛЕНИЕ: пустая строка сегмента давала 0 вместо "" в столбце A и #ЗНАЧ! в итогах
    For i = 13 To 20
        sc.Cells(i, 1).Formula = "=IF(" & Q(s) & "A" & (i - 8) & "="""",""""," & Q(s) & "A" & (i - 8) & ")"
    Next i
    Assume sc.Range("B9:D9"), R("dostizhimaya dolya 1/3/7 % (kak v metode potentsial'nykh klientov)")

    sw = Array( _
        Array("plantuml.com", 507000, 0.0907, "", R("17.09.2026")), _
        Array("eraser.io", 678500, 0.005, 0.4905, R("23.09.2026")), _
        Array("app.diagrams.net", 7800000, 0.002, 0.6865, R("23.09.2026")), _
        Array("mermaidchart.com", 20000, 0.007, "", R("17.09.2026")))
    For i = 0 To 3
        w.Cells(5 + i, 1).Value = sw(i)(0): w.Cells(5 + i, 2).Value = sw(i)(1): w.Cells(5 + i, 3).Value = sw(i)(2)
        w.Cells(5 + i, 5).Value = "": w.Cells(5 + i, 6).Value = sw(i)(3): w.Cells(5 + i, 7).Value = ""
        w.Cells(5 + i, 8).Value = 0.0012: w.Cells(5 + i, 10).Value = 120
        w.Cells(5 + i, 12).Value = R("{Similarweb}, vruchnuyu, ") & sw(i)(4)
    Next i
    w.Range("C4").Value = R("Dolya RF/RB {x} dolya kategorii")
    w.Range("H4").Value = R("Konversiya vizit {\u2192} {Pro}")
    Note w.Range("C4"), R("v shablone net stolbtsa relevantnosti - dolya kategorii (1 / 0,5 / 0,2 / 0,7) umnozhena na dolyu geografii")
    Todo w.Range("E5"), R("doli kanalov {Search}/{Direct}/{Social} dlya vsekh 4 domenov ({similarweb}.{com}, {Marketing} {channels})")

    src.Range("E5").Value = R("DOSNYaT'")
    src.Range("E6").Value = DateSerial(2026, 9, 23)
    src.Range("D9").Value = "https://www.eraser.io/pricing": src.Range("E9").Value = DateSerial(2026, 9, 17)
    src.Range("E10").Value = R("DOSNYaT' (opros {BO}-04)")

    Application.CalculateFull
    Set res = ResultSheet()
    PutRow res, "A3", Array(R("Segment"), R("Indeks platyozhesposobnosti"), R("{ROI} za god"))
    For i = 0 To 3
        res.Cells(4 + i, 1).Formula = "=" & Q(b) & "A" & (5 + i)
        res.Cells(4 + i, 2).Formula = "=" & Q(b) & "D" & (5 + i)
        res.Cells(4 + i, 3).Formula = "=" & Q(e) & "H" & (5 + i)
    Next i
    res.Range("A9").Value = R("Vyvod shablona"): res.Range("B9").Formula = "=" & Q(SH("08_")) & "A16"
    res.Range("A10").Value = R("Dostizhimaya vyruchka (baza), {BYN}/god"): res.Range("B10").Formula = "=" & Q(sc) & "G22"
    res.Range("B4:C7").NumberFormat = "0.00": res.Range("B10").NumberFormat = "# ##0"
    res.Columns("A:C").AutoFit
    AddChart res, res.Range("A3:C7"), R("Platyozhesposobnost' segmentov {NotaCode} (indeks i {ROI})"), "E3"
    If showMsg Then MsgBox R("Gotovo. Srednij indeks platyozhesposobnosti: ") & Format(SH("08_").Range("B7").Value, "0.00"), vbInformation, "NotaCode"
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
