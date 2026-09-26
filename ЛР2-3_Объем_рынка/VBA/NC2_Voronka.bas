Attribute VB_Name = "NC2_Voronka"
' NOTE (ASCII): comments are in Russian (UTF-8). Read them in VS Code / GitHub if VBE shows garbage.
' String literals are ASCII transliteration decoded by R() - works on any Windows ANSI code page.
'
' =====================================================================================
' ЛР2-3 «Оценка объёма рынка», дисциплина «Электронный бизнес». Проект NotaCode.
' Выполнила: Хаджинова К. Модуль NC2_Voronka: оценка рынка через цифровую воронку
' Шаблон: Материалы\shablon_excel_metodika_ocenki_rynka_cherez_cifrovuyu_voronku.xlsx
'
' Как запустить:
'   1) скопируйте шаблон из папки Материалы (оригинал не трогайте), откройте копию в Excel;
'   2) Alt+F11 -> File -> Import File... -> выберите этот файл NC2_Voronka.bas;
'   3) Alt+F8 -> NC2_Voronka.Main -> Выполнить (или F5 внутри процедуры Main);
'   4) сохраните книгу как .xlsm (если нужен макрос) или .xlsx (только результат).
'
' Что делает макрос:
'   1) вносит параметры NotaCode (чек 72/120/144 BYN, конверсии Free/Pro) в 01_Параметры;
'   2) заполняет 02_Трафик_охват планом каналов NotaCode на 1-й год (допущения);
'   3) заполняет 03_Конкуренты_SW реальными данными Similarweb (plantuml.com, eraser.io, app.diagrams.net, mermaidchart.com);
'   4) ИСПРАВЛЯЕТ ошибку шаблона: 03_Конкуренты_SW!C12 и C14 делили на '01_Параметры'!B12
'      (доля релевантного трафика) вместо B13 (коэффициент покрытия конкурентов);
'   5) обнуляет лист 07_Платформа_GMV (NotaCode - подписка, не платформа);
'   6) строит лист NotaCode_Итог с диаграммой сценариев.
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
    Dim p As Worksheet, t As Worksheet, c As Worksheet, g As Worksheet, src As Worksheet, res As Worksheet
    Dim ch As Variant, sw As Variant, i As Long
    Set p = SH("01_"): Set t = SH("02_"): Set c = SH("03_"): Set g = SH("07_"): Set src = SH("09_")

    ' --- 01_Параметры
    p.Range("B4").Value = R("{NotaCode} - {web}-{IDE} {diagram} {as} {code} dlya formal'nykh notatsij (tarif {Pro})")
    p.Range("B5").Value = R("Respublika Belarus' + russkoyazychnyj segment RF/SNG")
    p.Range("B6").Value = R("mesyats / god")
    p.Range("B7").Value = "BYN"
    p.Range("B8").Value = R("analitiki i arkhitektory ({Pro}), studenty ({Free\u2192Pro}), prepodavateli")
    p.Range("B9").Value = 120: p.Range("B10").Value = 72: p.Range("B11").Value = 144
    p.Range("C9:C11").Value = R("{BYN}/god")
    p.Range("D9").Value = R("{Pro} 40 {USD}/god {x} 3,00 {BYN}")
    p.Range("D10").Value = R("6 mes. pomesyachno {x} 12 {BYN}")
    p.Range("D11").Value = R("12 mes. pomesyachno {x} 12 {BYN}")
    p.Range("B12").Value = 0.35
    p.Range("B13").Value = 0.6
    p.Range("B14").Value = DateSerial(2026, 9, 25)
    Assume p.Range("B12"), R("kommercheski relevantnaya dolya trafika (LR3: 20/35/50 %)")
    Assume p.Range("B13"), R("v vyborke 4 domena iz fragmentirovannogo rynka {\u2192} pokrytie 0,6 (LR3: 0,60-0,75 tipichno)")
    PutRow p, "B17", Array(0.6, 0.02, 0.02)
    PutRow p, "B18", Array(0.8, 0.04, 0.03)
    PutRow p, "B19", Array(1, 0.06, 0.05)
    p.Range("F17").Value = 1: p.Range("F18").Value = 1: p.Range("F19").Value = 1.2
    p.Range("B16").Value = R("Dolya realizatsii plana okhvata")
    p.Range("C16").Value = R("Konversiya vizit {\u2192} registratsiya {Free}")
    p.Range("D16").Value = R("Konversiya {Free} {\u2192} {Pro}")
    p.Range("F16").Value = R("Prodleniya {Pro} (koe'f.)")
    Assume p.Range("B17:D19"), R("plan okhvata realizuet|sya na 60/80/100 %; {KPI}-05 {NotaCode}: {Free\u2192Pro} = 3 %")

    ' --- 02_Трафик_охват: план каналов NotaCode (1-й год)
    ch = Array( _
        Array(R("Organicheskij poisk"), R("Yandeks/{Google} po yadru (Vordstat - DOSNYaT')"), 30000, 0.05, 0.7, R("nizkoe")), _
        Array(R("Platnyj poisk"), R("Yandeks Direkt, testovaya kampaniya"), 10000, 0.04, 0.8, R("nizkoe")), _
        Array(R("Sotsseti i {Telegram}"), R("IT-soobshchestva, vuzovskie chaty"), 40000, 0.01, 0.5, R("nizkoe")), _
        Array(R("Referal'nye ploshchadki"), R("{GitHub} {README}, Khabr, katalogi"), 5000, 0.05, 0.7, R("nizkoe")), _
        Array(R("Pryamoj trafik"), R("ssylki iz metodichek, zakladki"), 300, 1, 0.8, R("nizkoe")), _
        Array(R("Katalogi instrumentov"), R("{AlternativeTo}, {Product} {Hunt}"), 3000, 0.03, 0.6, R("nizkoe")), _
        Array(R("Prochie (rekomendatsii prepodavatelej)"), R("e'kspertnaya otsenka"), 2000, 0.05, 0.9, R("nizkoe")))
    For i = 0 To 6
        t.Cells(4 + i, 1).Value = ch(i)(0): t.Cells(4 + i, 2).Value = ch(i)(1)
        t.Cells(4 + i, 3).Value = ch(i)(2): t.Cells(4 + i, 4).Value = ch(i)(3)
        t.Cells(4 + i, 6).Value = ch(i)(4): t.Cells(4 + i, 8).Value = ch(i)(5)
        t.Cells(4 + i, 9).Value = R("plan 1-go goda (dopushchenie); utochnit' posle zapuska po Metrike")
    Next i
    Assume t.Range("C4:C10"), R("produkt ne zapushchen, okhvat - plan po kanalam; dlya poiska zamenit' chastotnost'yu Vordstat")

    ' --- 03_Конкуренты_SW: реальные данные Similarweb (бесплатный доступ)
    sw = Array( _
        Array("plantuml.com", R("pryamoj ({DSL}-renderer)"), 507000, 0.0907, 1, R("17.09.2026; dolya RF 9,07 % (real'nye)")), _
        Array("eraser.io", R("pryamoj ({DSL} + {AI})"), 678500, 0.01, 0.5, R("23.09.2026; RF/RB vne top-5, 1 % - dopushchenie")), _
        Array("app.diagrams.net", R("zamenitel' ({GUI})"), 7800000, 0.01, 0.2, R("23.09.2026; RF/RB vne top-5, 1 % - dopushchenie")), _
        Array("mermaidchart.com", R("pryamoj (platnyj {Mermaid})"), 20000, 0.01, 0.7, R("17.09.2026; <20 tys. vizitov (verkhnyaya granitsa)")))
    For i = 0 To 3
        c.Cells(4 + i, 1).Value = sw(i)(0): c.Cells(4 + i, 2).Value = sw(i)(1)
        c.Cells(4 + i, 3).Value = sw(i)(2): c.Cells(4 + i, 4).Value = sw(i)(3): c.Cells(4 + i, 5).Value = sw(i)(4)
        c.Cells(4 + i, 7).Value = 0.0012: c.Cells(4 + i, 8).Value = 120
        c.Cells(4 + i, 10).Value = sw(i)(5)
    Next i
    c.Range("A8").Value = R("5-j sajt (napr., {mermaid}.{live})")
    c.Range("B8").Value = R("DOSNYaT'")
    c.Range("C8:E8").ClearContents: c.Range("G8:H8").ClearContents
    Todo c.Range("C8"), R("{similarweb}.{com}/{website}/{mermaid}.{live} - {Total} {Visits}, dolya RF/RB")
    c.Range("E3").Value = R("Dolya kategorii ""formal'nye notatsii""")
    c.Range("G3").Value = R("Konversiya vizit {\u2192} {Pro} (0,04 {x} 0,03)")
    Assume c.Range("E4:E7"), R("dolya trafika, otnosyashchayasya k formal'nym notatsiyam ({UML}/{BPMN}/{IDEF}/{ERD}/{DFD})")
    ' ИСПРАВЛЕНИЕ ОШИБКИ ШАБЛОНА: знаменатель - коэффициент покрытия (B13), а не доля трафика (B12)
    c.Range("C12").Formula = "=C11/" & Q(p) & "B13"
    c.Range("C14").Formula = "=C13/" & Q(p) & "B13*12"
    Note c.Range("C12"), R("ISPRAVLENO: bylo /{'}01_Parametry{'}!{B}12 (dolya trafika 0,7), stalo /{B}13 (koe'ffitsient pokrytiya)")
    Note c.Range("C14"), R("ISPRAVLENO: bylo /{'}01_Parametry{'}!{B}12, stalo /{B}13")

    ' --- 07_Платформа_GMV: не применимо
    g.Range("B4:D5").Value = 0: g.Range("B7:D7").Value = 0
    g.Range("G4").Value = R("Ne primenimo: {NotaCode} - podpisochnyj {SaaS}, a ne platforma sdelok")
    g.Range("G5").Value = R("{GMV} ne rasschityvaet|sya (metodika: {GMV} {\u2260} vyruchka)")

    ' --- 09_Источники
    src.Range("D4:D7").ClearContents
    src.Range("F4").Value = R("Produkt ne zapushchen - dannykh net; posle zapuska podklyuchit' Yandeks Metriku")
    src.Range("D8").Value = DateSerial(2026, 9, 23)
    src.Range("F8").Value = R("{plantuml}/{mermaidchart} - 17.09.2026, {eraser}/{app}.{diagrams}.{net} - 23.09.2026; besplatnyj dostup")
    src.Range("C10").Value = "https://www.eraser.io/pricing"
    src.Range("D10").Value = DateSerial(2026, 9, 17)
    src.Range("F10").Value = R("Tseny {Eraser}: {Starter} 15-20 {USD}, {Business} 45-60 {USD} za uchastnika v mesyats")

    Application.CalculateFull
    Set res = ResultSheet()
    res.Range("A3").Value = R("Stsenarij"): res.Range("B3").Value = R("Vyruchka/god s prodleniyami, {BYN}")
    res.Range("A4").Value = R("Ostorozhnyj"): res.Range("B4").Formula = "=" & Q(SH("06_")) & "E4"
    res.Range("A5").Value = R("Bazovyj"): res.Range("B5").Formula = "=" & Q(SH("06_")) & "E5"
    res.Range("A6").Value = R("Optimistichnyj"): res.Range("B6").Formula = "=" & Q(SH("06_")) & "E6"
    res.Range("A8").Value = R("Relevantnye poseshcheniya/mes."): res.Range("B8").Formula = "=" & Q(t) & "C14"
    res.Range("A9").Value = R("Registratsii/god (baza)"): res.Range("B9").Formula = "=" & Q(SH("04_")) & "C8*12"
    res.Range("A10").Value = R("Novye {Pro}/god (baza)"): res.Range("B10").Formula = "=" & Q(SH("04_")) & "C10*12"
    res.Range("A11").Value = R("Polnyj trafik segmenta po {SW}, mes. (ispravleno)"): res.Range("B11").Formula = "=" & Q(c) & "C12"
    res.Range("A12").Value = R("Vyruchka segmenta po {SW}, god (ispravleno)"): res.Range("B12").Formula = "=" & Q(c) & "C14"
    res.Range("B4:B12").NumberFormat = "# ##0"
    res.Columns("A:B").AutoFit
    AddChart res, res.Range("A3:B6"), R("Tsifrovaya voronka {NotaCode}: vyruchka po stsenariyam, {BYN}/god"), "D3"
    If showMsg Then MsgBox R("Gotovo. Bazovaya vyruchka voronki: ") & Format(res.Range("B5").Value, "0") & R(" {BYN}/god. Oshibka {C}12/{C}14 ispravlena."), vbInformation, "NotaCode"
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
