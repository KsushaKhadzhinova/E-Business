Attribute VB_Name = "NC7_ObemSW"
' NOTE (ASCII): comments are in Russian (UTF-8). Read them in VS Code / GitHub if VBE shows garbage.
' String literals are ASCII transliteration decoded by R() - works on any Windows ANSI code page.
'
' =====================================================================================
' ЛР2-3 «Оценка объёма рынка», дисциплина «Электронный бизнес». Проект NotaCode.
' Выполнила: Хаджинова К. Модуль NC7_ObemSW: оценка объёма рынка по данным Similarweb (ЛР2 шаблон + методики ЛР3)
' Шаблон: Материалы\metodika_ocenki_obema_rynka_po_similarweb.xlsx
'
' Как запустить:
'   1) скопируйте шаблон из папки Материалы (оригинал не трогайте), откройте копию в Excel;
'   2) Alt+F11 -> File -> Import File... -> выберите этот файл NC7_ObemSW.bas;
'   3) Alt+F8 -> NC7_ObemSW.Main -> Выполнить (или F5 внутри процедуры Main);
'   4) сохраните книгу как .xlsm (если нужен макрос) или .xlsx (только результат).
'
' Что делает макрос:
'   1) ИСПРАВЛЯЕТ ошибку шаблона: '01_Параметры'!B9/B10 суммировали G5:G24/H5:H24 и теряли 1-го конкурента
'      (строка 4) - теперь G4:G23/H4:H23;
'   2) заполняет 02_Конкуренты_SW реальными данными 4 доменов (визиты, bounce, pages/visit, Direct);
'   3) 06_Воронка_сценарии: покрытие 0,80/0,60/0,45, конверсии Free/Pro, чек 72/120/144 BYN, доля 1/3/7 %;
'   4) добавляет лист «NotaCode_ЛР3» с расчётами методик ЛР3: V_sample, V_geo, V_total = V_geo / охват,
'      CR3, CR5, HHI (пороги 1500 / 2500), TAM/SAM/SOM цифрового внимания;
'   5) NotaCode_Итог с диаграммой долей трафика.
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
    Dim p As Worksheet, c As Worksheet, g As Worksheet, q5 As Worksheet, f As Worksheet, src As Worksheet
    Dim res As Worksheet, l3 As Worksheet, sw As Variant, i As Long, qs As Variant, nm As String, cr As String
    Set p = SH("01_"): Set c = SH("02_"): Set g = SH("04_"): Set q5 = SH("05_"): Set f = SH("06_"): Set src = SH("08_")

    ' --- ИСПРАВЛЕНИЕ ОШИБКИ ШАБЛОНА
    p.Range("B9").Formula = "=SUM(" & Q(c) & "G4:G23)"
    p.Range("B10").Formula = "=SUM(" & Q(c) & "H4:H23)"
    Note p.Range("B9"), R("ISPRAVLENO: bylo {SUM}({G}5:{G}24) - teryalsya konkurent v stroke 4")
    Note p.Range("B10"), R("ISPRAVLENO: bylo {SUM}({H}5:{H}24)")

    p.Range("B4").Value = R("Instrumenty {diagram} {as} {code} dlya formal'nykh notatsij ({NotaCode})")
    p.Range("B5").Value = R("Russkoyazychnyj segment: RF + RB (dolya RB - DOSNYaT')")
    p.Range("B6").Value = R("poslednie 3 mesyatsa (okno besplatnogo {SW}); {Total} {Visits} za poslednij mesyats")
    p.Range("B7").Value = DateSerial(2026, 9, 23)
    p.Range("B11").Value = 0.6
    Assume p.Range("B11"), R("4 domena, rynok fragmentirovan (mnogo lokal'nykh {IDE}-plaginov {PlantUML}) {\u2192} 0,6")

    sw = Array( _
        Array("plantuml.com", "PlantUML", R("{DSL}-renderer {UML}"), R("pryamoj"), 507000, 0.0907, "", "", "", "", R("Rossiya 9,07 %; 17.09.2026; kanaly i vovlechyonnost' - DOSNYaT'")), _
        Array("eraser.io", "Eraser / DiagramGPT", R("{DSL} + {AI}-diagrammy"), R("pryamoj"), 678500, 0.005, 0.4373, 3.57, "00:01:39", 0.4905, R("Indiya 40,81 %, SShA 8,61 %; RF/RB 1 % {x} kategoriya 0,5; 23.09.2026")), _
        Array("app.diagrams.net", "draw.io", R("{GUI}-redaktor diagramm"), R("zamenitel'"), 7800000, 0.002, 0.555, 2.98, "", 0.6865, R("SShA 10,89 %, Indiya 8,29 %; RF/RB 1 % {x} kategoriya 0,2; 23.09.2026")), _
        Array("mermaidchart.com", "Mermaid Chart", R("platnyj {Mermaid}"), R("pryamoj"), 20000, 0.007, "", "", "", "", R("<20 tys.; V'etnam 44,98 %; RF/RB 1 % {x} 0,7; 17.09.2026")))
    For i = 0 To 3
        c.Cells(4 + i, 1).Value = sw(i)(0): c.Cells(4 + i, 2).Value = sw(i)(1): c.Cells(4 + i, 3).Value = sw(i)(2)
        c.Cells(4 + i, 4).Value = sw(i)(3): c.Cells(4 + i, 5).Value = sw(i)(4): c.Cells(4 + i, 6).Value = sw(i)(5)
        c.Cells(4 + i, 10).Value = sw(i)(6): c.Cells(4 + i, 11).Value = sw(i)(7): c.Cells(4 + i, 12).Value = sw(i)(8)
        c.Cells(4 + i, 13).Value = sw(i)(9)
        c.Range("N" & (4 + i) & ":R" & (4 + i)).ClearContents
        c.Cells(4 + i, 19).Value = sw(i)(10)
    Next i
    c.Range("A8:F8").ClearContents: c.Range("J8:S8").ClearContents
    c.Range("F3").Value = R("Dolya RF/RB {x} dolya kategorii")
    Note c.Range("F3"), R("v shablone net stolbtsa ""dolya relevantnosti"" (zamechanie 3) - ona peremnozhena s dolej geografii")
    Todo c.Range("M4"), R("doli kanalov {Direct}/{Organic}/{Referral}/{Social}/{Paid}/{Display} {plantuml}.{com}")
    Todo c.Range("N5"), R("ostal'nye kanaly {eraser}.{io} i {app}.{diagrams}.{net} (izvesten tol'ko {Direct})")
    Todo c.Range("M7"), R("kanaly {mermaidchart}.{com}")
    Todo c.Range("A8"), R("5-20 domenov: {mermaid}.{live}, {dbdiagram}.{io}, {lucidchart}.{com}, {gleek}.{io} ...")

    ' --- 04_География: доли снимает студентка
    PutRow g, "A4", Array(R("Belarus'")): PutRow g, "A5", Array(R("Minsk"))
    PutRow g, "A6", Array(R("Rossiya")): PutRow g, "A7", Array(R("Drugie strany SNG"))
    g.Range("B4:B7").ClearContents
    Todo g.Range("B4"), R("{Similarweb} {Geography}: doli RB/RF/SNG po kazhdomu domenu, zatem srednevzveshenno")
    g.Range("H6").Value = R("{plantuml}.{com}: Rossiya 9,07 % vsego trafika (real'nye)")

    ' --- 05_Запросы
    qs = Array(Array(R("{idef}0 onlajn"), R("kommercheskij"), 1, R("vklyuchit'")), Array(R("{uml} diagramma onlajn"), R("kommercheskij"), 1, R("vklyuchit'")), _
        Array(R("{dfd} diagramma"), R("instrumental'nyj"), 0.7, R("vklyuchit'")), Array("plantuml", R("brendovyj konkurenta"), 0.6, R("vklyuchit' otdel'no")), _
        Array("text to diagram ai", R("instrumental'nyj"), 0.7, R("vklyuchit'")), Array(R("{idef}0 primer kursovaya"), R("obuchayushchij"), 0.3, R("otdel'no")), _
        Array(R("chto takoe {uml}"), R("informatsionnyj"), 0.1, R("isklyuchit'")))
    For i = 0 To 6
        q5.Cells(4 + i, 1).Value = qs(i)(0): q5.Cells(4 + i, 2).Value = R("{Similarweb} / {Search} - DOSNYaT'")
        q5.Cells(4 + i, 3).ClearContents: q5.Cells(4 + i, 5).Value = qs(i)(1): q5.Cells(4 + i, 6).Value = qs(i)(2)
        q5.Cells(4 + i, 8).Value = R("notatsii"): q5.Cells(4 + i, 9).Value = qs(i)(3): q5.Cells(4 + i, 10).Value = ""
    Next i
    Todo q5.Range("C4"), R("doli poiskovogo trafika po zaprosam ({SW} {Search} / {Keywords}, besplatno vidna chast')")

    ' --- 06_Воронка_сценарии
    PutColumn f, "B5", Array(0.8, 0.6, 0.45)
    PutColumn f, "D5", Array(0.02, 0.04, 0.06)
    PutColumn f, "E5", Array(0.02, 0.03, 0.05)
    PutColumn f, "F5", Array(72, 120, 144)
    PutColumn f, "G5", Array(1, 1, 1)
    PutColumn f, "C12", Array(0.01, 0.03, 0.07)
    f.Range("D4").Value = R("Konversiya v registratsiyu {Free}")
    f.Range("E4").Value = R("Konversiya {Free} {\u2192} {Pro}")
    f.Range("F4").Value = R("Godovoj chek {Pro}, {BYN}")
    f.Range("H4").Value = R("Novykh {Pro}/mes.")
    Assume f.Range("B5:B7"), R("ostorozhnyj = plotnaya vyborka (0,8) {\u2192} men'shij rynok; sm. ZADANIE_LR3.{md} o raskhozhdenii shkal")
    f.Range("H19").Value = "BYN/" & R("god")

    ' --- 08_Источники: журнал ручного сбора
    src.Range("E4:E9").Value = DateSerial(2026, 9, 23)
    src.Range("A14:F20").ClearContents
    PutRow src, "A14", Array(DateSerial(2026, 9, 17), "plantuml.com", "Traffic & Engagement", "Total visits", 507000, R("snimok - DOSNYaT' ({scr}_05)"))
    PutRow src, "A15", Array(DateSerial(2026, 9, 17), "plantuml.com", "Geography", R("Rossiya"), 0.0907, R("top-strana"))
    PutRow src, "A16", Array(DateSerial(2026, 9, 17), "mermaidchart.com", "Traffic & Engagement", "Total visits", "<20000", R("malo dannykh"))
    PutRow src, "A17", Array(DateSerial(2026, 9, 23), "eraser.io", "Traffic & Engagement", "Total visits", 678500, R("{bounce} 43,73 %, 3,57 str."))
    PutRow src, "A18", Array(DateSerial(2026, 9, 23), "eraser.io", "Marketing channels", "Direct", 0.4905, R("organika - 2-j kanal"))
    PutRow src, "A19", Array(DateSerial(2026, 9, 23), "app.diagrams.net", "Traffic & Engagement", "Total visits", 7800000, R("{bounce} 55,5 %, 2,98 str."))
    PutRow src, "A20", Array(DateSerial(2026, 9, 23), "app.diagrams.net", "Marketing channels", "Direct", 0.6865, "")

    ' --- лист расчётов ЛР3
    nm = R("{NotaCode}_LR3")
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set l3 = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    l3.Name = nm
    cr = Q(c)
    l3.Range("A1").Value = R("Metodiki LR3: tsifrovoj ob''yom, kontsentratsiya, {TAM}/{SAM}/{SOM} tsifrovogo vnimaniya")
    l3.Range("A1").Font.Bold = True
    PutRow l3, "A3", Array(R("Pokazatel'"), R("Znachenie"), R("Formula / kommentarij"))
    PutRow l3, "A4", Array("V_sample", "", R("summa vizitov vyborki (vse strany) = {TAM} tsifrovogo vnimaniya"))
    l3.Range("B4").Formula = "=SUM(" & cr & "E4:E23)"
    PutRow l3, "A5", Array("V_geo", "", R("summa vizity {x} (dolya RF/RB {x} kategoriya) = relevantnyj trafik"))
    l3.Range("B5").Formula = "=" & Q(p) & "B9"
    PutRow l3, "A6", Array(R("Okhvat vyborki (baza)"), "", R("0,80-0,90 plotnaya; 0,60-0,75 tipichno; 0,40-0,55 fragmentirovan"))
    l3.Range("B6").Formula = "=" & Q(p) & "B11"
    PutRow l3, "A7", Array("V_total", "", "V_total = V_geo / " & R("okhvat"))
    l3.Range("B7").Formula = "=B5/B6"
    PutRow l3, "A8", Array(R("Kommercheski relevantnye vizity"), "", R("{V}_{total} {x} 35 % (LR3, razd. 16: 20/35/50 %)"))
    l3.Range("B8").Formula = "=B7*0.35"
    PutRow l3, "A9", Array(R("Rynok, {BYN}/god (baza)"), "", R("{x} 4 % {x} 3 % {x} 120 {BYN} {x} 12"))
    l3.Range("B9").Formula = "=B8*0.04*0.03*120*12"
    PutRow l3, "A10", Array(R("{SOM} (baza), {BYN}/god"), "", R("{x} dostizhimaya dolya 3 %"))
    l3.Range("B10").Formula = "=B9*0.03"
    PutRow l3, "A12", Array("CR3, %", "", R("summa 3 krupnejshikh dolej (po relevantnomu trafiku)"))
    l3.Range("B12").Formula = "=100*(LARGE(" & cr & "I4:I23,1)+LARGE(" & cr & "I4:I23,2)+LARGE(" & cr & "I4:I23,3))"
    PutRow l3, "A13", Array("CR5, %", "", R("v vyborke 4 domena {\u2192} {CR}5 = {CR}4"))
    l3.Range("B13").Formula = "=100*SUM(" & cr & "I4:I23)"
    PutRow l3, "A14", Array("HHI", "", R("summa kvadratov dolej v % (= doli^2 {x} 10 000)"))
    l3.Range("B14").Formula = "=SUMPRODUCT(" & cr & "I4:I23*100," & cr & "I4:I23*100)"
    PutRow l3, "A15", Array(R("Uroven' kontsentratsii"), "", R("do 1 500 - nizkaya; 1 500-2 500 - umerennaya; vyshe 2 500 - vysokaya"))
    l3.Range("B15").Formula = "=IF(B14<1500,""" & R("nizkaya") & """,IF(B14<=2500,""" & R("umerennaya") & """,""" & R("vysokaya") & """))"
    PutRow l3, "A17", Array(R("{SAM} po kanalu, vizitov/mes."), "", R("vizity {x} dolya RF/RB {x} dolya ne-pryamykh kanalov (bez uchyota kategorii)"))
    l3.Range("B17").Value = 51002
    Note l3.Range("B17"), R("raschyot v {Python}\{nc}_{model}.{py} ({lr}3_{similarweb}): 45 985 {x} 0,5 + 6 785 {x} 0,5095 + 78 000 {x} 0,3135 + 200 {x} 0,5")
    PutRow l3, "A18", Array(R("{SOM} po trafiku, vizitov/mes."), "", R("{SAM} po kanalu {x} 3 %"))
    l3.Range("B18").Formula = "=B17*0.03"
    PutRow l3, "A19", Array(R("{SOM} po vyruchke, {BYN}/god"), "", R("{x} 4 % {x} 3 % {x} 120 {x} 12"))
    l3.Range("B19").Formula = "=B18*0.04*0.03*120*12"
    l3.Range("B4:B10,B17:B19").NumberFormat = "# ##0": l3.Range("B12:B14").NumberFormat = "0.0"
    l3.Columns("A:C").AutoFit

    Application.CalculateFull
    Set res = ResultSheet()
    PutRow res, "A3", Array(R("Domen"), R("Dolya relevantnogo trafika"))
    For i = 0 To 3
        res.Cells(4 + i, 1).Formula = "=" & cr & "A" & (4 + i)
        res.Cells(4 + i, 2).Formula = "=" & cr & "I" & (4 + i)
    Next i
    res.Range("B4:B7").NumberFormat = "0.0%"
    res.Range("A9").Value = R("Rynok/god: ostorozhnyj - bazovyj - optimistichnyj")
    res.Range("B9").Formula = "=" & Q(f) & "B22"
    res.Range("A10").Value = "HHI": res.Range("B10").Formula = "=" & Q(l3) & "B14"
    res.Columns("A:B").AutoFit
    AddChart res, res.Range("A3:B7"), R("Doli relevantnogo trafika ({Similarweb}), {HHI} - sm. list {NotaCode}_LR3"), "D3"
    If showMsg Then MsgBox R("Gotovo. {HHI} = ") & Format(l3.Range("B14").Value, "0") & " (" & l3.Range("B15").Value & R("). Oshibka {B}9/{B}10 ispravlena."), vbInformation, "NotaCode"
End Sub

Private Sub PutColumn(ByVal ws As Worksheet, ByVal addr As String, ByVal v As Variant)
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
