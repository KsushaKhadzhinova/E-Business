Attribute VB_Name = "NC5_SverkhuSnizu"
' NOTE (ASCII): comments are in Russian (UTF-8). Read them in VS Code / GitHub if VBE shows garbage.
' String literals are ASCII transliteration decoded by R() - works on any Windows ANSI code page.
'
' =====================================================================================
' ЛР2-3 «Оценка объёма рынка», дисциплина «Электронный бизнес». Проект NotaCode.
' Выполнила: Хаджинова К. Модуль NC5_SverkhuSnizu: подходы «сверху вниз» и «снизу вверх» с Similarweb + сверка трёх методов
' Шаблон: Материалы\shablon_excel_metodika_sverhu_vniz_snizu_vverh_similarweb.xlsx
'
' Как запустить:
'   1) скопируйте шаблон из папки Материалы (оригинал не трогайте), откройте копию в Excel;
'   2) Alt+F11 -> File -> Import File... -> выберите этот файл NC5_SverkhuSnizu.bas;
'   3) Alt+F8 -> NC5_SverkhuSnizu.Main -> Выполнить (или F5 внутри процедуры Main);
'   4) сохраните книгу как .xlsm (если нужен макрос) или .xlsx (только результат).
'
' Что делает макрос:
'   1) 01_Параметры: годовой чек 120 BYN, конверсии Free/Pro, корректировка выборки 1/0,6;
'   2) 02_Сверху_вниз: денежная база «оборот ИТ-сектора РБ» = ВВП 78,59 млрд USD x 6,1 % x 3,00 BYN
'      (реальные данные) и 4 коэффициента сужения; строка 10 (чек) = 1, т.к. база уже в деньгах;
'   3) 03_Снизу_вверх: сегменты S-01...S-04 (C = потребность x доля готовых платить);
'   4) 04_Similarweb: реальные данные 4 доменов; лишняя демо-строка обнуляется;
'   5) сверка трёх методов (06_Сверка) и лист NotaCode_Итог с диаграммой (лог. шкала).
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
    Dim p As Worksheet, t As Worksheet, u As Worksheet, w As Worksheet, v As Worksheet, src As Worksheet, res As Worksheet
    Dim seg As Variant, sw As Variant, i As Long
    Set p = SH("01_"): Set t = SH("02_"): Set u = SH("03_"): Set w = SH("04_"): Set v = SH("06_"): Set src = SH("08_")

    p.Range("B5").Value = R("Instrumenty modelirovaniya v formal'nykh notatsiyakh ({diagram} {as} {code}) - {NotaCode}")
    p.Range("B6").Value = R("Respublika Belarus' (+ RF/SNG po trafiku {Similarweb})")
    p.Range("B7").Value = "BYN"
    p.Range("B8").Value = 120: p.Range("C8").Value = R("{BYN}/pol'zovatel'/god")
    p.Range("E8").Value = R("{Pro} 40 {USD}/god {x} 3,00 (kurs - dopushchenie)")
    p.Range("B9").Value = 1
    p.Range("B10").Value = 0.04: p.Range("B11").Value = 0.03: p.Range("B12").Value = 1
    p.Range("A10").Value = R("Konversiya vizita v registratsiyu {Free}")
    p.Range("A11").Value = R("Konversiya {Free} {\u2192} {Pro}")
    p.Range("B13").Formula = "=1/0.6"
    Assume p.Range("B13"), R("1 / koe'ffitsient okhvata vyborki 0,6 (4 domena, rynok fragmentirovan)")
    Assume p.Range("B10:B12"), R("{SaaS}-orientiry; {KPI}-05 {NotaCode} = 3 %; godovoj chek uzhe vklyuchaet prodleniya")

    ' --- 02_Сверху_вниз: денежная база (ВВП)
    t.Range("A5").Value = R("Oborot IT-sektora RB, {BYN}/god (VVP 78,59 mlrd {USD} {x} 6,1 % {x} 3,00)")
    t.Range("B5:D5").Formula = "=78590000000*0.061*3"
    t.Range("E5").Value = "BYN": t.Range("F5").Value = R("{macrotrends}.{net} (VVP 2024), {opencompanyinbelarus}.{com} (6,1 %), snyato 17.09.2026")
    t.Range("A6").Value = R("Dolya raskhodov na instrumenty razrabotki"): PutRow t, "B6", Array(0.015, 0.02, 0.025)
    t.Range("A7").Value = R("Dolya instrumentov modelirovaniya i diagramm"): PutRow t, "B7", Array(0.04, 0.05, 0.06)
    t.Range("A8").Value = R("Dolya {web} / {diagram}-{as}-{code} segmenta"): PutRow t, "B8", Array(0.05, 0.1, 0.15)
    t.Range("A9").Value = R("Dolya notatsij, kotorye pokryvaet {NotaCode}"): PutRow t, "B9", Array(0.5, 0.6, 0.7)
    t.Range("F6").Value = R("dopushchenie"): t.Range("F7").Value = R("dopushchenie")
    t.Range("F8").Value = R("dopushchenie"): t.Range("F9").Value = R("dopushchenie ({UML}/{BPMN}/{ERD}/{IDEF}/{DFD} {vs} {C}4/seti/oblako)")
    t.Range("A10").Value = R("Koe'ffitsient pereschyota (baza uzhe denezhnaya, chek ne primenyaet|sya)")
    t.Range("B10:D10").Value = 1: t.Range("E10").Value = R("koe'f.")
    t.Range("A11").Value = R("Denezhnyj ob''yom posle suzheniya ({SAM} sverkhu vniz)"): t.Range("E11").Value = R("{BYN}/god")
    Note t.Range("B10"), R("IZMENENO: v shablone ={B}8 (chek); baza stroki 5 - den'gi, poe'tomu mnozhitel' 1")
    Assume t.Range("B6:D9"), R("koe'ffitsienty suzheniya e'kspertnye - glavnoe uzkoe mesto metoda (pravilo soglasovaniya 1)")

    ' --- 03_Снизу_вверх
    seg = Array( _
        Array(R("{S}-01 Studenty IT-spetsial'nostej"), 19600, 0.024, 0.9, 72, R("{belarus}.{by} + dopushcheniya")), _
        Array(R("{S}-02 Analitiki"), 1600, 0.06, 0.9, 120, R("PVT + dopushcheniya")), _
        Array(R("{S}-03 Arkhitektory i razrabotchiki"), 1800, 0.04, 0.9, 120, R("PVT + dopushcheniya")), _
        Array(R("{S}-04 Prepodavateli"), 252, 0.07, 0.9, 120, R("21 vuz + dopushcheniya")))
    For i = 0 To 3
        u.Cells(5 + i, 1).Value = seg(i)(0): u.Cells(5 + i, 2).Value = seg(i)(1): u.Cells(5 + i, 3).Value = seg(i)(2)
        u.Cells(5 + i, 4).Value = seg(i)(3): u.Cells(5 + i, 5).Value = 1: u.Cells(5 + i, 6).Value = seg(i)(4)
        u.Cells(5 + i, 9).Value = seg(i)(5): u.Cells(5 + i, 10).Value = R("Zapolneno")
    Next i
    u.Range("A9").Value = R("RF/SNG (net dannykh)"): u.Range("B9:F9").Value = 0: u.Range("I9").Value = R("DOSNYaT'"): u.Range("J9").Value = R("net dannykh")
    u.Range("A10").Value = R("Rezerv"): u.Range("B10:F10").Value = 0: u.Range("I10").Value = "": u.Range("J10").Value = ""
    u.Range("C4").Value = R("Dolya s potrebnost'yu {x} dolya gotovykh platit'")
    Note u.Range("C5"), R("0,8 {x} 0,03: v shablone net stolbtsa platyozhesposobnosti, poe'tomu dolya gotovykh platit' vklyuchena syuda")

    ' --- 04_Similarweb
    sw = Array( _
        Array("plantuml.com", R("Pryamoj"), 507000, 0.0907, 0.5, 0.35, 0.8, R("Rossiya 9,07 %; kanaly - DOSNYaT'"), "2026-09-17"), _
        Array("eraser.io", R("Pryamoj"), 678500, 0.01, 0.5095, 0.175, 0.8, R("{Direct} 49,05 %; {bounce} 43,73 %; 3,57 str."), "2026-09-23"), _
        Array("app.diagrams.net", R("Zamenitel'"), 7800000, 0.01, 0.3135, 0.07, 0.7, R("{Direct} 68,65 %; {bounce} 55,5 %; 2,98 str."), "2026-09-23"), _
        Array("mermaidchart.com", R("Pryamoj"), 20000, 0.01, 0.5, 0.245, 0.7, R("<20 tys. vizitov; kanaly - DOSNYaT'"), "2026-09-17"))
    For i = 0 To 3
        w.Cells(5 + i, 1).Value = sw(i)(0): w.Cells(5 + i, 2).Value = sw(i)(1): w.Cells(5 + i, 3).Value = sw(i)(2)
        w.Cells(5 + i, 4).Value = sw(i)(3): w.Cells(5 + i, 5).Value = sw(i)(4): w.Cells(5 + i, 6).Value = sw(i)(5)
        w.Cells(5 + i, 7).Value = sw(i)(6)
        w.Cells(5 + i, 11).Value = "https://www.similarweb.com/website/" & sw(i)(0) & "/"
        w.Cells(5 + i, 12).Value = sw(i)(8): w.Cells(5 + i, 13).Value = sw(i)(7)
    Next i
    w.Range("A9:B9").ClearContents: w.Range("C9:G9").Value = 0: w.Range("K9:M9").ClearContents
    Todo w.Range("A9"), R("5-10 domenov dopolnitel'no ({mermaid}.{live}, {dbdiagram}.{io}, {lucidchart}.{com} ...) - metodika trebuet 10-20")
    Assume w.Range("D6:D8"), R("RF/RB vne top-5 stran {\u2192} 1 %")
    Assume w.Range("F5:F8"), R("0,35 {x} dolya kategorii ""formal'nye notatsii""")
    Assume w.Range("G5:G8"), R("koe'ffitsient vovlechyonnosti po {bounce} {rate}")

    src.Range("E5").Value = "2026-09-17 / 2026-09-23"
    src.Range("H5").Value = R("Zapolneno (4 domena)")
    src.Range("C8").Value = R("{macrotrends}.{net} (VVP 2024 = 78,59 mlrd {USD}); {opencompanyinbelarus}.{com} (IT = 6,1 % VVP, PVT > 1000)")
    src.Range("E8").Value = "2026-09-17": src.Range("H8").Value = R("Zapolneno")
    src.Range("D9").Value = "https://www.eraser.io/pricing": src.Range("E9").Value = "2026-09-17"
    src.Range("H9").Value = R("Zapolneno")

    Application.CalculateFull
    Set res = ResultSheet()
    PutRow res, "A3", Array(R("Metod"), R("Ostorozhnyj"), R("Bazovyj"), R("Optimistichnyj"))
    For i = 0 To 5
        res.Cells(4 + i, 1).Formula = "=" & Q(v) & "A" & (5 + i)
        res.Cells(4 + i, 2).Formula = "=" & Q(v) & "B" & (5 + i)
        res.Cells(4 + i, 3).Formula = "=" & Q(v) & "C" & (5 + i)
        res.Cells(4 + i, 4).Formula = "=" & Q(v) & "D" & (5 + i)
    Next i
    res.Range("B4:D9").NumberFormat = "# ##0"
    res.Columns("A:D").AutoFit
    AddChart res, res.Range("A3:D6"), R("Sverka tryokh metodov, {BYN}/god (log. shkala)"), "F3", True
    If showMsg Then MsgBox R("Gotovo. Nizhnyaya otsenka: ") & Format(res.Range("B7").Value, "0") & R(", verkhnyaya: ") & _
        Format(res.Range("D9").Value, "0") & R(" {BYN}/god."), vbInformation, "NotaCode"
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
