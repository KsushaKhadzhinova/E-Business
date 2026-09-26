Attribute VB_Name = "NC1_PoiskSpros"
' NOTE (ASCII): comments are in Russian (UTF-8). Read them in VS Code / GitHub if VBE shows garbage.
' String literals are ASCII transliteration decoded by R() - works on any Windows ANSI code page.
'
' =====================================================================================
' ЛР2-3 «Оценка объёма рынка», дисциплина «Электронный бизнес». Проект NotaCode.
' Выполнила: Хаджинова К. Модуль NC1_PoiskSpros: оценка рынка через поисковый спрос
' Шаблон: Материалы\shablon_excel_ocenka_rynka_cherez_poiskovyy_spros.xlsx
'
' Как запустить:
'   1) скопируйте шаблон из папки Материалы (оригинал не трогайте), откройте копию в Excel;
'   2) Alt+F11 -> File -> Import File... -> выберите этот файл NC1_PoiskSpros.bas;
'   3) Alt+F8 -> NC1_PoiskSpros.Main -> Выполнить (или F5 внутри процедуры Main);
'   4) сохраните книгу как .xlsm (если нужен макрос) или .xlsx (только результат).
'
' Что делает макрос:
'   1) создаёт (если нет) лист «Ввод_NotaCode» для частотности Вордстат - заполняете его сами;
'   2) вносит в 01_Параметры, 02_Семантика, 03_Google_Trends, 06_Воронка, 09_Источники данные NotaCode;
'   3) реальные данные Google Trends (PlantUML, text to diagram, 2024-01...2024-12, снято 17.09.2026);
'   4) удаляет демо-частотность «маркетинг» из 04_Wordstat и переносит туда ваши данные с листа ввода;
'   5) строит лист NotaCode_Итог с диаграммой сценариев.
'   Пока Вордстат не снят, выручка по методу = 0 (метод не рассчитан) - это честно, не ошибка.
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
    Dim p As Worksheet, s As Worksheet, gt As Worksheet, wsd As Worksheet, fv As Worksheet
    Dim src As Worksheet, inp As Worksheet, res As Worksheet
    Dim q As Variant, i As Long, r0 As Long, hasWS As Boolean
    Set p = SH("01_"): Set s = SH("02_"): Set gt = SH("03_"): Set wsd = SH("04_")
    Set fv = SH("06_"): Set src = SH("09_")
    Set inp = InputSheet()

    ' --- 01_Параметры
    p.Range("B4").Value = R("Respublika Belarus' + russkoyazychnyj segment (RF/SNG)")
    p.Range("B5").Value = R("12 mes.: {GT} 2024-01...2024-12 (snyato 17.09.2026); Vordstat - poslednie 12 mes.")
    p.Range("B6").Value = 120: p.Range("C6").Value = R("{BYN}/god")
    p.Range("D6").Value = R("Godovoj tarif {Pro} 40 {USD} {x} 3,00")
    p.Range("B7").Value = 0.06: p.Range("B8").Value = 0.04: p.Range("B9").Value = 0.03
    p.Range("B10").Value = 1.1: p.Range("B11").Value = 0.4: p.Range("B12").Value = 0.6: p.Range("B13").Value = 0.5
    p.Range("A8").Value = R("Konversiya posetitel' {\u2192} registratsiya {Free}")
    p.Range("A9").Value = R("Konversiya {Free} {\u2192} {Pro}")
    Assume p.Range("B7:B10"), R("orientiry {SaaS}-voronki; {KPI}-05 {NotaCode}: {Free\u2192Pro} = 3 %")

    ' --- 02_Семантика: семантическое ядро NotaCode
    q = Array( _
        Array(R("{idef}0 onlajn"), R("Instrumental'nyj spros"), R("kommercheskoe"), R("RB"), "", 1, 0.8), _
        Array(R("{uml} diagramma onlajn"), R("Instrumental'nyj spros"), R("kommercheskoe"), R("RB"), "", 0.9, 0.7), _
        Array(R("{dfd} diagramma onlajn"), R("Instrumental'nyj spros"), R("kommercheskoe"), R("RB"), "", 1, 0.7), _
        Array(R("{bpmn} onlajn redaktor"), R("Instrumental'nyj spros"), R("kommercheskoe"), R("RB"), "", 0.8, 0.7), _
        Array(R("{er} diagramma onlajn"), R("Instrumental'nyj spros"), R("smeshannoe"), R("RB"), "", 0.8, 0.6), _
        Array(R("{plantuml}"), R("Brendovyj (konkurent)"), R("smeshannoe"), "Worldwide", 41, 0.9, 0.5), _
        Array(R("{text} {to} {diagram}"), R("Instrumental'nyj ({AI})"), R("smeshannoe"), "Worldwide", 26, 0.8, 0.5), _
        Array(R("{diagram} {as} {code}"), R("Obshchij spros"), R("informatsionnoe"), "Worldwide", 3, 1, 0.3), _
        Array(R("{idef}0 primer kursovaya"), R("Uchebnyj spros"), R("informatsionno-problemnoe"), R("RB"), "", 0.7, 0.3), _
        Array(R("chto takoe {uml}"), R("Isklyuchaemyj"), R("informatsionnoe"), R("RB"), "", 0.4, 0.1))
    For i = 0 To 9
        r0 = 4 + i
        s.Cells(r0, 2).Value = q(i)(0): s.Cells(r0, 3).Value = q(i)(1): s.Cells(r0, 4).Value = q(i)(2)
        s.Cells(r0, 5).Value = q(i)(3)
        s.Cells(r0, 6).Value = inp.Cells(19 + i, 2).Value
        s.Cells(r0, 7).Value = q(i)(4): s.Cells(r0, 8).Value = q(i)(5): s.Cells(r0, 9).Value = q(i)(6)
        If IsEmpty(inp.Cells(19 + i, 2).Value) Then Todo s.Cells(r0, 6), R("chastotnost' Vordstat, region Belarus', list Vvod_{NotaCode}!{B}") & (19 + i)
        If VarType(q(i)(4)) = vbString Then Todo s.Cells(r0, 7), R("indeks {Google} {Trends} (Belarus', 12 mes.) - po zhelaniyu")
    Next i
    Note s.Range("G9"), R("real'nye: srednij indeks {GT} za 5 let, {Worldwide}, snyato 17.09.2026")
    Assume s.Range("H4:I13"), R("vesa relevantnosti i kommercheskoj blizosti po tablitse metodiki (razd. 4.1)")

    ' --- 03_Google_Trends: реальные данные (архив DiagramCode, ЛР1)
    gt.Range("B3").Value = R("{PlantUML} ({Worldwide})")
    gt.Range("C3").Value = R("{text} {to} {diagram} ({Worldwide})")
    gt.Range("D3").Value = R("{uml} diagramma (Belarus')")
    gt.Range("E3").Value = R("{idef}0 (Belarus')")
    PutColumn gt, "A4", Array(DateSerial(2024, 1, 1), DateSerial(2024, 2, 1), DateSerial(2024, 3, 1), _
        DateSerial(2024, 4, 1), DateSerial(2024, 5, 1), DateSerial(2024, 6, 1), DateSerial(2024, 7, 1), _
        DateSerial(2024, 8, 1), DateSerial(2024, 9, 1), DateSerial(2024, 10, 1), DateSerial(2024, 11, 1), _
        DateSerial(2024, 12, 1))
    PutColumn gt, "B4", Array(37, 35, 46, 48, 47, 49, 40, 37, 39, 50, 51, 53)
    PutColumn gt, "C4", Array(7, 9, 9, 11, 10, 6, 5, 6, 8, 10, 11, 12)
    gt.Range("D4:E15").ClearContents
    Todo gt.Range("D4"), R("{trends}.{google}.{com}: {uml} diagramma i {idef}0, Belarus', 12 mes., {CSV} {\u2192} {D}4:{E}15")
    gt.Range("G4").Value = R("Real'nye dannye {Google} {Trends}, 1-ya nedelya mesyatsa, snyato 17.09.2026 (arkhiv LR1)")

    ' --- 04_Wordstat: только ваши данные (демо «маркетинг» удаляется)
    hasWS = Application.WorksheetFunction.Count(inp.Range("B4:E15")) > 0
    For i = 0 To 11
        wsd.Cells(4 + i, 1).Value = inp.Cells(4 + i, 1).Value
        wsd.Cells(4 + i, 2).Resize(1, 4).Value = inp.Cells(4 + i, 2).Resize(1, 4).Value
        wsd.Cells(4 + i, 7).Value = inp.Cells(4 + i, 6).Value
    Next i
    wsd.Range("B3").Value = R("Kommercheskie ({idef}0/{uml}/{dfd} onlajn)")
    wsd.Range("C3").Value = R("Problemnye (kak narisovat' ...)")
    wsd.Range("D3").Value = R("Instrumental'nye ({bpmn}, {er}, {plantuml})")
    wsd.Range("E3").Value = R("Uchebnye/lokal'nye")
    If Not hasWS Then Todo wsd.Range("B4"), R("zapolnite list Vvod_{NotaCode} (Vordstat, region Belarus') i zapustite {Main} eshchyo raz")

    ' --- 06_Воронка: сценарии NotaCode
    PutRow fv, "B5", Array(0.04, 0.06, 0.1)
    PutRow fv, "B7", Array(0.02, 0.04, 0.06)
    PutRow fv, "B9", Array(0.02, 0.03, 0.05)
    PutRow fv, "B11", Array(72, 120, 144)
    PutRow fv, "B12", Array(1, 1.1, 1.2)
    fv.Range("A7").Value = R("Konversiya posetitel' {\u2192} registratsiya {Free}")
    fv.Range("A8").Value = R("Registratsii / mes.")
    fv.Range("A9").Value = R("Konversiya {Free} {\u2192} {Pro}")
    fv.Range("A10").Value = R("Novye podpiski {Pro} / mes.")
    fv.Range("A11").Value = R("Godovoj chek {Pro}, {BYN}")
    fv.Range("A12").Value = R("Koe'ffitsient prodlenij")
    fv.Range("E11").Value = R("72 = 6 mes. {x} 12; 120 = 40 {USD}/god; 144 = 12 mes. {x} 12 {BYN}")
    Assume fv.Range("B5:D12"), R("dolya perekhodov 4/6/10 %, vizit{\u2192}registratsiya 2/4/6 %, {Free\u2192Pro} 2/3/5 % ({KPI}-05 = 3 %)")

    ' --- 09_Источники
    src.Range("C4").Value = R("{Worldwide}, 5 let, vse kategorii; mesyachnaya vyborka 2021-09...2024-12")
    src.Range("D4").Value = DateSerial(2026, 9, 17)
    src.Range("C6").Value = R("DOSNYaT': Belarus', poslednie 12 mes., vse ustrojstva")
    src.Range("D6").ClearContents
    Todo src.Range("D6"), R("data snyatiya Vordstat")
    src.Range("C8").Value = R("ne ispol'zovalis' (produkt ne zapushchen)")
    src.Range("C9").Value = R("net ({MVP} k 06.12.2026)")

    Application.CalculateFull
    ' --- Итоги
    Set res = ResultSheet()
    res.Range("A3").Value = R("Stsenarij"): res.Range("B3").Value = R("Vyruchka/god, {BYN}")
    res.Range("A4").Value = R("Ostorozhnyj"): res.Range("B4").Formula = "=" & Q(fv) & "B14"
    res.Range("A5").Value = R("Bazovyj"): res.Range("B5").Formula = "=" & Q(fv) & "C14"
    res.Range("A6").Value = R("Optimistichnyj"): res.Range("B6").Formula = "=" & Q(fv) & "D14"
    res.Range("A8").Value = R("Srednij svodnyj indeks sprosa")
    res.Range("B8").Formula = "=" & Q(SH("08_")) & "B4"
    res.Range("A9").Value = R("Vordstat vnesyon?")
    If hasWS Then res.Range("B9").Value = R("da") Else res.Range("B9").Value = R("net - metod ne rasschitan (DOSNYaT')")
    res.Range("B4:B6").NumberFormat = "# ##0"
    res.Columns("A:B").AutoFit
    AddChart res, res.Range("A3:B6"), R("Poiskovyj spros: vyruchka {NotaCode} po stsenariyam, {BYN}/god"), "D3"
    Application.CalculateFull
    If showMsg Then MsgBox R("Gotovo. Bazovaya vyruchka po poiskovomu sprosu: ") & Format(res.Range("B5").Value, "0") & _
        R(" {BYN}/god. Vordstat vnesyon: ") & res.Range("B9").Value, vbInformation, "NotaCode"
End Sub

' Лист «Ввод_NotaCode»: создаётся один раз и не перезаписывается (ваши данные сохраняются).
Private Function InputSheet() As Worksheet
    Dim ws As Worksheet, nm As String, i As Long, qs As Variant
    nm = R("Vvod_{NotaCode}")
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set InputSheet = ws
        Exit Function
    End If
    Set ws = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1))
    ws.Name = nm
    ws.Range("A1").Value = R("Vvod dannykh Yandeks Vordstat ({wordstat}.{yandex}.{ru}, vkhod po Yandeks {ID}), region Belarus'")
    ws.Range("A1").Font.Bold = True
    PutRow ws, "A3", Array(R("Mesyats"), R("Kommercheskie"), R("Problemnye"), R("Instrumental'nye"), R("Uchebnye/lokal'nye"), R("Kommercheskij koe'ffitsient"))
    For i = 0 To 11
        ws.Cells(4 + i, 1).Value = DateSerial(2025, 10 + i, 1)
        ws.Cells(4 + i, 1).NumberFormat = "mmm yyyy"
        ws.Cells(4 + i, 6).Value = 0.6
    Next i
    ws.Range("B4:E15").Interior.Color = RGB(255, 242, 204)
    Assume ws.Range("F4"), R("dolya kommercheskikh namerenij v summe grupp; utochnit' po razmetke yadra")
    ws.Range("A17").Value = R("Chastotnost' otdel'nykh zaprosov (pokazov v mesyats, Belarus') {\u2192} 02_Semantika!{F}")
    ws.Range("A17").Font.Bold = True
    PutRow ws, "A18", Array(R("Zapros"), R("Pokazov v mesyats"))
    qs = Array(R("{idef}0 onlajn"), R("{uml} diagramma onlajn"), R("{dfd} diagramma onlajn"), R("{bpmn} onlajn redaktor"), _
               R("{er} diagramma onlajn"), "plantuml", "text to diagram", "diagram as code", R("{idef}0 primer kursovaya"), R("chto takoe {uml}"))
    For i = 0 To 9
        ws.Cells(19 + i, 1).Value = qs(i)
    Next i
    ws.Range("B19:B28").Interior.Color = RGB(255, 242, 204)
    ws.Range("A30").Value = R("Gruppy (stolbtsy {B}-{E}): summa pokazov zaprosov gruppy za mesyats iz vkladki Dinamika.")
    ws.Columns("A:F").AutoFit
    Set InputSheet = ws
End Function

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
