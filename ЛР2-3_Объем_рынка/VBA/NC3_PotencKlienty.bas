Attribute VB_Name = "NC3_PotencKlienty"
' NOTE (ASCII): comments are in Russian (UTF-8). Read them in VS Code / GitHub if VBE shows garbage.
' String literals are ASCII transliteration decoded by R() - works on any Windows ANSI code page.
'
' =====================================================================================
' ЛР2-3 «Оценка объёма рынка», дисциплина «Электронный бизнес». Проект NotaCode.
' Выполнила: Хаджинова К. Модуль NC3_PotencKlienty: оценка объёма рынка через количество потенциальных клиентов (TAM/SAM/SOM)
' Шаблон: Материалы\shablon_excel_metodika_ocenki_rynka_cherez_kolichestvo_potencialnyh_klientov.xlsx
'
' Как запустить:
'   1) скопируйте шаблон из папки Материалы (оригинал не трогайте), откройте копию в Excel;
'   2) Alt+F11 -> File -> Import File... -> выберите этот файл NC3_PotencKlienty.bas;
'   3) Alt+F8 -> NC3_PotencKlienty.Main -> Выполнить (или F5 внутри процедуры Main);
'   4) сохраните книгу как .xlsm (если нужен макрос) или .xlsx (только результат).
'
' Что делает макрос:
'   1) заполняет 01_Параметры и 02_Сегменты сегментами S-01...S-04 NotaCode (реальные базы: ~7000 выпускников ИТ/год,
'      ПВТ >1000 компаний, 21 вуз; остальное - допущения с примечаниями);
'   2) задаёт достижимую долю 1 / 3 / 7 % (К5 методики) в 03_Сценарии;
'   3) заполняет 05_Similarweb реальными данными 4 доменов;
'   4) обновляет 07_Источники; строит NotaCode_Итог с диаграммой TAM/SAM/SOM (логарифмическая шкала).
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
    Dim p As Worksheet, s As Worksheet, sc As Worksheet, w As Worksheet, src As Worksheet, res As Worksheet
    Dim seg As Variant, sw As Variant, i As Long
    Set p = SH("01_"): Set s = SH("02_"): Set sc = SH("03_"): Set w = SH("05_"): Set src = SH("07_")

    p.Range("B4").Value = R("{NotaCode} - {web}-{IDE} {diagram} {as} {code} dlya formal'nykh notatsij")
    p.Range("B5").Value = R("Razrabotka PO, modelirovanie ({UML}, {BPMN}, {ERD}, {IDEF}0/1{X}/3, {DFD}, seti Petri)")
    p.Range("B6").Value = R("Podpiska {Pro} (4 {USD}/mes ili 40 {USD}/god): vse formaty e'ksporta, polnaya istoriya versij, {AI}-kvota")
    p.Range("B7").Value = R("Respublika Belarus' (RF/SNG - rasshirenie, dannykh net)")
    p.Range("B8").Value = R("{B}2{C} (studenty, prepodavateli) + {B}2{B}-pol'zovateli (analitiki, arkhitektory)")
    p.Range("B9").Value = "BYN"
    p.Range("B10").Value = R("1 god")
    p.Range("B11").Value = DateSerial(2026, 9, 25)
    p.Range("B12").Value = R("Khadzhinova K.")

    seg = Array( _
        Array(R("{S}-01 Studenty IT-spetsial'nostej RB"), "B2C", R("{belarus}.{by}: ~7000 vypusknikov/god {x} 4 kursa (dopushchenie)"), 28000, 0.7, 0.8, 0.9, 0.03, 72, R("Srednyaya")), _
        Array(R("{S}-02 Biznes- i sistemnye analitiki"), "B2B", R("PVT >1000 kompanij {x} 2 analitika (dopushchenie)"), 2000, 0.8, 0.6, 0.9, 0.1, 120, R("Nizkaya")), _
        Array(R("{S}-03 Arkhitektory PO i razrabotchiki"), "B2B", R("PVT >1000 kompanij {x} 3 spetsialista (dopushchenie)"), 3000, 0.6, 0.5, 0.9, 0.08, 120, R("Nizkaya")), _
        Array(R("{S}-04 Prepodavateli tekhn. distsiplin"), "B2C", R("21 vuz ({belarus}.{by}) {x} 15 prepodavatelej (dopushchenie)"), 315, 0.8, 0.7, 0.9, 0.1, 120, R("Nizkaya")))
    For i = 0 To 3
        s.Cells(4 + i, 1).Value = seg(i)(0): s.Cells(4 + i, 2).Value = seg(i)(1): s.Cells(4 + i, 3).Value = seg(i)(2)
        s.Cells(4 + i, 4).Value = seg(i)(3): s.Cells(4 + i, 5).Value = seg(i)(4): s.Cells(4 + i, 7).Value = seg(i)(5)
        s.Cells(4 + i, 8).Value = seg(i)(6): s.Cells(4 + i, 9).Value = seg(i)(7): s.Cells(4 + i, 11).Value = seg(i)(8)
        s.Cells(4 + i, 12).Value = 1: s.Cells(4 + i, 14).Value = 0.03: s.Cells(4 + i, 16).Value = seg(i)(9)
    Next i
    s.Range("Q4").Value = R("{Free} - osnovnoj tarif; platit men'shinstvo (3 %); chek 72 = semestr 6 mes. {x} 12 {BYN}")
    s.Range("Q5").Value = R("Rabota s {BPMN}/{IDEF}0/{DFD}/{ERD}; platit rabotodatel' ili sam spetsialist")
    s.Range("Q6").Value = R("{UML}, {ERD}, e'ksport v {DDL}/{PlantUML}/{Mermaid}")
    s.Range("Q7").Value = R("Shablony i primery dlya zanyatij")
    s.Range("A8").Value = R("Russkoyazychnyj segment RF/SNG")
    s.Range("B8").Value = R("smeshannyj"): s.Range("C8").Value = R("DOSNYaT': statistika RF ne sobiralas'")
    s.Range("D8:E8").Value = 0: s.Range("G8:I8").Value = 0: s.Range("K8:L8").Value = 0: s.Range("N8").Value = 0
    s.Range("P8").Value = R("Net dannykh"): s.Range("Q8").Value = R("uchityvaet|sya tol'ko kachestvenno (optimistichnyj stsenarij)")
    s.Range("A9").Value = R("Rezerv"): s.Range("B9").Value = "": s.Range("C9").Value = ""
    s.Range("D9:E9").Value = 0: s.Range("G9:I9").Value = 0: s.Range("K9:L9").Value = 0: s.Range("N9").Value = 0
    s.Range("P9").Value = R("Net dannykh"): s.Range("Q9").Value = ""
    Assume s.Range("E4:E7"), R("dolya v granitsakh: studenty, u kotorykh est' distsipliny s notatsiyami; spetsialisty, rabotayushchie s formal'nymi notatsiyami")
    Assume s.Range("I4:I7"), R("dolya gotovykh platit' {Pro}; dlya studentov = konversiya {Free\u2192Pro} 3 % ({KPI}-05)")
    Assume s.Range("N4:N7"), R("dostizhimaya dolya 3 % (K5 tema1.{docx}: bazovo 1-3 %)")
    Note s.Range("D4"), R("real'nye: ~7000 vypusknikov IT-spetsial'nostej v god, 21 vuz ({belarus}.{by}); {x}4 kursa - dopushchenie")
    Note s.Range("D5"), R("real'nye: bolee 1000 kompanij-rezidentov PVT ({opencompanyinbelarus}.{com}); 2 analitika na kompaniyu - dopushchenie")

    PutCol sc, "G4", Array(0.01, 0.03, 0.07)
    Assume sc.Range("G4:G6"), R("dostizhimaya dolya 1/3/7 % - verkhnie granitsy diapazonov K5 (unikal'naya nisha {IDEF}/{DFD}, otlichie {D}-01)")

    sw = Array( _
        Array("PlantUML", "https://www.similarweb.com/website/plantuml.com/", R("pryamoj konkurent"), 507000, 0.0907, 0.5, 0.35), _
        Array("Eraser", "https://www.similarweb.com/website/eraser.io/", R("pryamoj konkurent ({AI})"), 678500, 0.01, 0.5095, 0.175), _
        Array("draw.io", "https://www.similarweb.com/website/app.diagrams.net/", R("zamenitel' ({GUI})"), 7800000, 0.01, 0.3135, 0.07), _
        Array("Mermaid Chart", "https://www.similarweb.com/website/mermaidchart.com/", R("pryamoj konkurent"), 20000, 0.01, 0.5, 0.245))
    For i = 0 To 3
        w.Cells(4 + i, 1).Value = sw(i)(0): w.Cells(4 + i, 2).Value = sw(i)(1): w.Cells(4 + i, 3).Value = sw(i)(2)
        w.Cells(4 + i, 4).Value = sw(i)(3): w.Cells(4 + i, 5).Value = sw(i)(4): w.Cells(4 + i, 7).Value = sw(i)(5)
        w.Cells(4 + i, 8).Value = sw(i)(6)
        w.Cells(4 + i, 10).Value = 0.04: w.Cells(4 + i, 11).Value = 0.03: w.Cells(4 + i, 12).Value = 120
    Next i
    w.Range("A8").Value = R("DOSNYaT': 5-j konkurent"): w.Range("B8").Value = "": w.Range("C8").Value = ""
    w.Range("D8:E8").Value = 0: w.Range("G8:H8").Value = 0: w.Range("J8:L8").Value = 0
    w.Range("J3").Value = R("Konversiya v registratsiyu {Free}")
    w.Range("K3").Value = R("Konversiya {Free} {\u2192} {Pro}")
    Note w.Range("E4"), R("real'nye: Rossiya 9,07 % trafika {plantuml}.{com} (17.09.2026)")
    Assume w.Range("E5:E7"), R("RF/RB vne top-5 stran {\u2192} 1 %")
    Note w.Range("G5:G6"), R("real'nye: 1 - dolya {Direct} ({eraser} 49,05 %, {draw}.{io} 68,65 %)")
    Todo w.Range("G4"), R("kanaly {plantuml}.{com} v {Similarweb} ({Marketing} {channels}); poka 0,5")
    Todo w.Range("G7"), R("kanaly {mermaidchart}.{com} v {Similarweb}; poka 0,5")
    Assume w.Range("H4:H7"), R("0,35 {x} dolya kategorii ""formal'nye notatsii"" (1 / 0,5 / 0,2 / 0,7)")

    src.Range("C4").Value = R("{belarus}.{by} (vypuskniki IT, vuzy); {opencompanyinbelarus}.{com} (PVT)")
    src.Range("E4").Value = DateSerial(2026, 9, 17): src.Range("F4").Value = R("Fakt + dopushchenie")
    src.Range("C5").Value = R("e'kspertno; proverit' oprosom >= 30 respondentov ({BO}-04)")
    src.Range("C8").Value = R("Tsena {NotaCode} {Pro}: 4 {USD}/mes, 40 {USD}/god; kurs 3,00 (dopushchenie, {nbrb}.{by})")
    src.Range("E8").Value = DateSerial(2026, 9, 25): src.Range("F8").Value = R("E'kspertnoe dopushchenie")
    src.Range("E9").Value = DateSerial(2026, 9, 23)
    src.Range("H9").Value = R("{plantuml}.{com}, {mermaidchart}.{com} - 17.09.2026; {eraser}.{io}, {app}.{diagrams}.{net} - 23.09.2026")

    Application.CalculateFull
    Set res = ResultSheet()
    PutRow res, "A3", Array(R("Stsenarij"), "TAM", "SAM", "SOM")
    For i = 0 To 2
        res.Cells(4 + i, 1).Formula = "=" & Q(sc) & "A" & (4 + i)
        res.Cells(4 + i, 2).Formula = "=" & Q(sc) & "H" & (4 + i)
        res.Cells(4 + i, 3).Formula = "=" & Q(sc) & "I" & (4 + i)
        res.Cells(4 + i, 4).Formula = "=" & Q(sc) & "J" & (4 + i)
    Next i
    res.Range("A8").Value = R("Adresuemykh klientov (baza)"): res.Range("B8").Formula = "=" & Q(SH("04_")) & "C6"
    res.Range("A9").Value = R("{SOM} / vyruchka konkurentov po {SW}"): res.Range("B9").Formula = "=" & Q(SH("04_")) & "C16"
    res.Range("B4:D8").NumberFormat = "# ##0": res.Range("B9").NumberFormat = "0.0%"
    res.Columns("A:D").AutoFit
    AddChart res, res.Range("A3:D6"), R("{TAM} / {SAM} / {SOM} {NotaCode} po stsenariyam, {BYN}/god (log. shkala)"), "F3", True
    If showMsg Then MsgBox R("Gotovo. {SAM} bazovyj: ") & Format(res.Range("C5").Value, "0") & R(" {BYN}/god, {SOM}: ") & _
        Format(res.Range("D5").Value, "0") & R(" {BYN}/god."), vbInformation, "NotaCode"
End Sub

Private Sub PutCol(ByVal ws As Worksheet, ByVal addr As String, ByVal v As Variant)
    Dim k As Long
    For k = LBound(v) To UBound(v)
        ws.Range(addr).Offset(k - LBound(v), 0).Value = v(k)
    Next k
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
