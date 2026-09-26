Attribute VB_Name = "LR4_NotaCode"
Option Explicit
'==============================================================================
' ЛР4. Пять сил Портера и барьеры входа (Similarweb) — проект NotaCode
' Выполнила: Хаджинова К.   Дисциплина «Электронный бизнес»
'
' Что делает модуль (макрос Main):
'  1. Создаёт лист «Ввод_NotaCode»: экспертные баллы факторов, для которых нет
'     данных (допущения), значения платформенной зависимости и заменителей,
'     исходные и НОРМИРОВАННЫЕ веса (в шаблоне сумма весов = 1,04 -> приводится к 1,00).
'  2. Чинит 01_Параметры: формулы в колонке B съехали на строку
'     (B8 «Дата» содержала COUNTA, B9 «Число конкурентов» — SUM, B10 была пуста).
'  3. Удаляет демо-строки (agency-a.by и т.п.) и вносит 18 конкурентов NotaCode
'     в листы 02, 05, 06 и 10 запросов в лист 04. Трафик SimilarWeb — только
'     реально снятый в ЛР2-3 (17.09.2026 и 23.09.2026) по 4 доменам; остальное —
'     оранжевые ячейки с примечанием «ДОСНЯТЬ».
'  4. Защищает расчётные формулы 04/05/06 от пустых строк (раньше пустая строка
'     давала индекс 0 и балл 1, занижая средние), чинит 05!B38 (формула массива)
'     и 06!B38 (сумма двух долей могла превышать 100 %).
'  5. Лист 07: если по фактору нет данных, балл берётся с листа Ввод_NotaCode
'     (колонка J показывает источник балла: «расчёт» или «экспертно»).
'  6. Добавляет лист 10_Портер_NotaCode (матрица 5 сил, цифровые усилители,
'     8 критериев методики Similarweb) и заполняет журнал 09_Источники.
'
' Как запустить:
'  1. ВАЖНО: исходный шаблон из Материалы\ Excel НЕ ОТКРЫВАЕТ (в 08_Дашборд!A15
'     строковая константа формулы длиннее 255 символов). Используйте исправленную
'     копию Excel\shablon_ocenka_barerov_ISPRAVLEN.xlsx: откройте её и сохраните как
'     ЛР4_NotaCode_барьеры_входа.xlsm («Книга Excel с поддержкой макросов»).
'  2. Alt+F11 -> File -> Import File -> VBA\LR4_NotaCode.bas
'  3. Закройте редактор, Alt+F8 -> Main -> Выполнить.
'  4. После досъёма данных впишите их в оранжевые ячейки — формулы пересчитаются
'     сами; экспертный балл на листе 07 автоматически заменится расчётным.
'  Повторный запуск Main безопасен: он заново перезапишет листы данными NotaCode
'  (вписанные вами досъёмные значения при этом будут стёрты — сохраните копию).
'==============================================================================

Private Const SH_PAR As String = "01_Параметры"
Private Const SH_SW As String = "02_Конкуренты_SW"
Private Const SH_CH As String = "03_Каналы_трафика"
Private Const SH_SR As String = "04_Поиск_реклама"
Private Const SH_RV As String = "05_Отзывы_рейтинги"
Private Const SH_TF As String = "06_Технологии_финансы"
Private Const SH_BR As String = "07_Барьеры_входа"
Private Const SH_DB As String = "08_Дашборд"
Private Const SH_SRC As String = "09_Источники"
Private Const SH_IN As String = "Ввод_NotaCode"
Private Const SH_PT As String = "10_Портер_NotaCode"

Private Const TODO_TXT As String = "ДОСНЯТЬ"

Public Sub Main()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim nm As Variant
    For Each nm In Array(SH_PAR, SH_SW, SH_CH, SH_SR, SH_RV, SH_TF, SH_BR, SH_DB, SH_SRC)
        If Not SheetExists(wb, CStr(nm)) Then
            MsgBox "В активной книге нет листа «" & nm & "». Откройте копию шаблона " & _
                   "shablon_excel_ocenka_konkurencii_i_barerov_vhoda_similarweb.xlsx.", vbCritical
            Exit Sub
        End If
    Next nm

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    BuildInputSheet wb
    FixParams wb.Worksheets(SH_PAR)
    FillCompetitors wb.Worksheets(SH_SW)
    MarkChannels wb.Worksheets(SH_CH)
    FillSearch wb.Worksheets(SH_SR)
    FillReviews wb.Worksheets(SH_RV)
    FillTech wb.Worksheets(SH_TF)
    FixBarriers wb.Worksheets(SH_BR)
    BuildPorterSheet wb
    FixDashboard wb.Worksheets(SH_DB)
    FillSources wb.Worksheets(SH_SRC)

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.ScreenUpdating = True

    Dim b As Worksheet
    Set b = wb.Worksheets(SH_BR)
    MsgBox "Готово." & vbCrLf & _
           "CR3 = " & Format(wb.Worksheets(SH_SW).Range("B38").Value, "0.0%") & _
           ", HHI = " & Format(wb.Worksheets(SH_SW).Range("B40").Value, "0") & vbCrLf & _
           "Сумма весов = " & Format(b.Range("E13").Value, "0.00") & vbCrLf & _
           "Индекс барьеров входа = " & Format(b.Range("F13").Value, "0.00") & " (" & b.Range("G13").Value & ")" & vbCrLf & _
           "Оранжевые ячейки — данные, которые нужно доснять.", vbInformation, "ЛР4 NotaCode"
    wb.Worksheets(SH_DB).Activate
End Sub

'------------------------------------------------------------------ служебные
Private Function SheetExists(wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function GetOrAddSheet(wb As Workbook, ByVal nm As String, ByVal afterName As String) As Worksheet
    If SheetExists(wb, nm) Then
        Set GetOrAddSheet = wb.Worksheets(nm)
        GetOrAddSheet.Cells.Clear
    Else
        Set GetOrAddSheet = wb.Worksheets.Add(After:=wb.Worksheets(afterName))
        GetOrAddSheet.Name = nm
    End If
End Function

' Пометить ячейки как «ДОСНЯТЬ»: оранжевая заливка + примечание
Private Sub MarkTodo(rng As Range, ByVal note As String)
    rng.Interior.Color = RGB(255, 204, 153)
    Dim c As Range
    Set c = rng.Cells(1, 1)
    On Error Resume Next
    c.ClearComments
    c.AddComment TODO_TXT & ": " & note
    c.Comment.Shape.Width = 260
    c.Comment.Shape.Height = 80
    On Error GoTo 0
End Sub

Private Sub PutRow(ws As Worksheet, ByVal r As Long, ByVal firstCol As Long, vals As Variant)
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        If Not IsEmpty(vals(i)) Then ws.Cells(r, firstCol + i - LBound(vals)).Value = vals(i)
    Next i
End Sub

' Список конкурентов NotaCode (одинаковый порядок во всех листах)
Private Function Competitors() As Variant
    Competitors = Array( _
        Array("app.diagrams.net (draw.io)", "GUI-редактор (заменитель)"), _
        Array("eraser.io", "прямой (DSL + AI)"), _
        Array("plantuml.com", "прямой (DSL, open-source)"), _
        Array("mermaidchart.com", "прямой (DSL, коммерческий)"), _
        Array("mermaid.live", "прямой (DSL, open-source)"), _
        Array("d2lang.com", "прямой (DSL)"), _
        Array("graphviz.org", "прямой (DSL-движок)"), _
        Array("structurizr.com", "нишевой (C4 DSL)"), _
        Array("gleek.io", "прямой (DSL + AI)"), _
        Array("lucidchart.com", "GUI SaaS (частичный)"), _
        Array("miro.com", "платформа-доска (частичный)"), _
        Array("camunda.com (Modeler)", "нишевой (BPMN)"), _
        Array("bpmn.io", "нишевой (BPMN, open-source)"), _
        Array("ramussoftware.com (Ramus)", "нишевой (IDEF0/DFD)"), _
        Array("sparxsystems.com (Enterprise Architect)", "CASE enterprise"), _
        Array("visual-paradigm.com", "CASE enterprise"), _
        Array("dbdiagram.io", "нишевой (ERD DSL)"), _
        Array("kroki.io", "агрегатор-рендер"))
End Function

'------------------------------------------------------------------ Ввод_NotaCode
Private Sub BuildInputSheet(wb As Workbook)
    Dim ws As Worksheet
    Set ws = GetOrAddSheet(wb, SH_IN, SH_SRC)
    ws.Range("A1").Value = "Ввод NotaCode: экспертные баллы (допущения) и веса факторов барьеров входа"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Экспертный балл используется на листе 07 ТОЛЬКО пока по фактору нет данных. " & _
                           "Строки совпадают со строками листа 07_Барьеры_входа."
    ws.Range("A3:G3").Value = Array("Фактор", "Строка листа 07", "Экспертный балл 1/3/5", "Обоснование (допущение)", _
                                    "Вес шаблона", "Вес нормированный", "Комментарий")
    ws.Range("A3:G3").Font.Bold = True

    Dim f As Variant, i As Long
    f = Array( _
        Array("Концентрация трафика", 5, "Не используется: считается по 02 (CR3/HHI).", 0.16), _
        Array("Органическое SEO-давление", 3, "Выдача по англоязычным запросам занята сильными доменами; по русским IDEF/DFD-запросам — слабее.", 0.12), _
        Array("Платное рекламное давление", 3, "Ниша инструментов разработчика: реклама есть, но не доминирует (нет данных Ad Library).", 0.12), _
        Array("Брендовая сила конкурентов", 5, "Используется, только если в 02 нет долей Direct.", 0.12), _
        Array("Поисково-рекламная конкуренция", 3, "Нет CPC/SEO difficulty — средний уровень до досъёма.", 0.12), _
        Array("Репутационный барьер", 5, "PlantUML с 2009 г., draw.io, Lucidchart, Miro — массовые отзывы и сообщества.", 0.12), _
        Array("Технологический и ресурсный барьер", 3, "DSL-инструменты — небольшие open-source команды; Lucidchart/Miro/Sparx — крупные компании.", 0.12), _
        Array("Платформенная зависимость", 3, "Балл считается из значения B15.", 0.08), _
        Array("Угроза заменителей", 5, "Балл считается из значения B16.", 0.08))
    For i = 0 To 8
        ws.Cells(4 + i, 1).Value = f(i)(0)
        ws.Cells(4 + i, 2).Value = 4 + i
        ws.Cells(4 + i, 3).Value = f(i)(1)
        ws.Cells(4 + i, 4).Value = f(i)(2)
        ws.Cells(4 + i, 5).Value = f(i)(3)
        ws.Cells(4 + i, 6).Formula = "=E" & (4 + i) & "/SUM($E$4:$E$12)"
        ws.Cells(4 + i, 3).Interior.Color = RGB(255, 242, 204)
    Next i
    ws.Range("A13").Value = "Итого"
    ws.Range("E13").Formula = "=SUM(E4:E12)"
    ws.Range("F13").Formula = "=SUM(F4:F12)"
    ws.Range("G13").Value = "В шаблоне сумма весов 1,04 -> индекс завышается (макс. 5,2). Нормирование: w/1,04."
    ws.Range("F4:F13").NumberFormat = "0.0000"

    ws.Range("A15").Value = "Платформенная зависимость (0–1)"
    ws.Range("B15").Value = 0.5
    ws.Range("D15").Value = "Допущение: поиск Google/Яндекс, GitHub, VK/Telegram-сообщества вузов дают около половины доступа к клиенту."
    ws.Range("A16").Value = "Угроза заменителей (0–1)"
    ws.Range("B16").Value = 0.8
    ws.Range("D16").Value = "Допущение: бесплатный draw.io, LLM (генерация PlantUML/Mermaid-кода), Visio/Word, Ramus в вузах, рисование от руки."
    ws.Range("B15:B16").Interior.Color = RGB(255, 242, 204)
    ws.Columns("A").ColumnWidth = 36
    ws.Columns("D").ColumnWidth = 80
    ws.Columns("G").ColumnWidth = 40
End Sub

'------------------------------------------------------------------ 01_Параметры
Private Sub FixParams(ws As Worksheet)
    ws.Range("B5").Value = "Web-IDE «diagram as code» для формальных нотаций (UML, BPMN, ERD, IDEF0/IDEF1X/IDEF3, DFD, сети Петри)"
    ws.Range("B6").Value = "Республика Беларусь + РФ/СНГ (русскоязычный сегмент)"
    ws.Range("B7").Value = "последние 3 месяца (SimilarWeb, снято 17.09.2026 и 23.09.2026)"
    ' Исправление сдвига формул на строку
    ws.Range("B8").Value = DateSerial(2026, 9, 25)
    ws.Range("B8").NumberFormat = "dd.mm.yyyy"
    ws.Range("B9").Formula = "=COUNTA('" & SH_SW & "'!A4:A33)"
    ws.Range("B10").Formula = "=SUM('" & SH_SW & "'!E4:E33)"
    ws.Range("B10").NumberFormat = "#,##0"
    ws.Range("F8").Value = "ИСПРАВЛЕНО: была формула COUNTA (число конкурентов)"
    ws.Range("F9").Value = "02_Конкуренты_SW (ИСПРАВЛЕНО: была SUM трафика)"
    ws.Range("F10").Value = "02_Конкуренты_SW (ИСПРАВЛЕНО: ячейка была пустой)"
End Sub

'------------------------------------------------------------------ 02_Конкуренты_SW
Private Sub FillCompetitors(ws As Worksheet)
    Dim r As Long, comp As Variant, i As Long
    ws.Range("A4:D33").ClearContents
    ws.Range("G4:Q33").ClearContents
    ws.Range("A4:D33,G4:Q33").Interior.Color = RGB(255, 242, 204)
    ws.Range("A4:Q33").ClearComments
    comp = Competitors()
    For i = 0 To UBound(comp)
        r = 4 + i
        ws.Cells(r, 1).Value = comp(i)(0)
        ws.Cells(r, 2).Value = comp(i)(1)
        ws.Cells(r, 5).Formula = "=IFERROR(C" & r & "*D" & r & ",0)"
        ws.Cells(r, 6).Formula = "=IFERROR(E" & r & "/SUM($E$4:$E$33),0)"
    Next i

    ' Реальные данные SimilarWeb (ЛР2-3 старой версии). Доля геогр. = 1,0 — допущение:
    ' доли RU/BY сняты только для plantuml.com (Россия 9,07 %).
    ' Столбцы: C визиты, D доля геогр., G длит., H стр/визит, I отказы, J direct
    PutRow ws, 4, 3, Array(7800000, 1, Empty, Empty, Empty, 2.98, 0.555, 0.6865)
    ws.Range("Q4").Value = "SimilarWeb 23.09.2026: 7,8 млн визитов, США 10,89 %; direct 68,65 %. Длительность и каналы кроме direct — ДОСНЯТЬ"
    PutRow ws, 5, 3, Array(678500, 1, Empty, Empty, 99, 3.57, 0.4373, 0.4905)
    ws.Range("Q5").Value = "SimilarWeb 23.09.2026: Индия 40,81 %, США 8,61 %; direct 49,05 %, далее органика (доля — ДОСНЯТЬ)"
    PutRow ws, 6, 3, Array(507000, 1)
    ws.Range("Q6").Value = "SimilarWeb 17.09.2026: Россия 9,07 % (топ-1 страна); вовлечённость и каналы — ДОСНЯТЬ"
    PutRow ws, 7, 3, Array(20000, 1)
    ws.Range("Q7").Value = "SimilarWeb 17.09.2026: «< 20 тыс.» (верхняя граница); Вьетнам 44,98 %, Индонезия 26,27 %, США 15,1 %"

    For r = 4 To 7
        MarkTodo ws.Range("D" & r), "доля RU+BY из Similarweb -> Geography. Пока 1,0 (глобальный трафик, допущение)."
    Next r
    MarkTodo ws.Range("G4"), "длительность визита app.diagrams.net"
    MarkTodo ws.Range("K4:P5"), "доли каналов Organic/Paid/Social/Referrals/Display и брендовый поиск (Similarweb -> Marketing channels)"
    MarkTodo ws.Range("G6:P7"), "вовлечённость и каналы plantuml.com / mermaidchart.com"
    MarkTodo ws.Range("C8:D21"), "визиты и доля RU+BY: similarweb.com/website/<домен>, период «последние 3 месяца», одна дата для всех"
    MarkTodo ws.Range("G8:P21"), "вовлечённость и каналы (Similarweb, та же дата)"
    For r = 8 To 21
        ws.Cells(r, 17).Value = "SimilarWeb не снимался — ДОСНЯТЬ"
    Next r
    ws.Range("D4:D33,F4:F33,I4:P33").NumberFormat = "0.00%"
    ws.Range("C4:C33,E4:E33").NumberFormat = "#,##0"
End Sub

Private Sub MarkChannels(ws As Worksheet)
    MarkTodo ws.Range("C5"), "органика = 0, пока в 02 не внесены доли Organic; на листе 07 до досъёма используется экспертный балл"
    MarkTodo ws.Range("C6"), "платный поиск = 0, пока в 02 не внесены доли Paid"
    ws.Range("F10").Value = "Нижняя граница: Direct известен только у app.diagrams.net и eraser.io, брендовый поиск не снят."
End Sub

'------------------------------------------------------------------ 04_Поиск_реклама
Private Sub FillSearch(ws As Worksheet)
    Dim q As Variant, r As Long, i As Long
    ws.Range("A4:J33").ClearContents
    ws.Range("M4:M33").ClearContents
    ws.Range("A4:M33").ClearComments
    q = Array( _
        Array("plantuml онлайн", "инструментальный"), _
        Array("uml диаграмма онлайн", "инструментальный"), _
        Array("idef0 онлайн / программа для idef0", "инструментальный"), _
        Array("dfd диаграмма онлайн", "инструментальный"), _
        Array("bpmn редактор онлайн", "инструментальный"), _
        Array("er диаграмма онлайн", "инструментальный"), _
        Array("сети петри онлайн", "инструментальный"), _
        Array("diagram as code", "проблемный"), _
        Array("text to diagram / ai diagram generator", "коммерческий"), _
        Array("draw.io аналог", "коммерческий"))
    For i = 0 To UBound(q)
        r = 4 + i
        ws.Cells(r, 1).Value = q(i)(0)
        ws.Cells(r, 2).Value = q(i)(1)
        ws.Cells(r, 3).Value = "Вордстат (РБ+РФ) / Keyword Planner / выдача / Ads Transparency"
        ws.Cells(r, 13).Value = "Метрики — ДОСНЯТЬ"
    Next i
    MarkTodo ws.Range("D4:J13"), "частотность (Вордстат, Яндекс ID), CPC и PPC 0–1 (Keyword Planner), SEO difficulty (Semrush/Ahrefs), число сильных доменов в топ-10, рекламодатели и объявления (Ads Transparency / Meta Ad Library)"
    For r = 4 To 33
        ws.Cells(r, 11).Formula = "=IF(OR($A" & r & "="""",COUNT($E" & r & ":$H" & r & ",$J" & r & ")<5),""""," & _
            "MIN(1,AVERAGE(MIN(E" & r & "/2,1),F" & r & ",G" & r & "/100,MIN(H" & r & "/10,1),MIN(J" & r & "/50,1))))"
        ws.Cells(r, 12).Formula = "=IF(K" & r & "="""","""",IF(K" & r & ">=0.7,5,IF(K" & r & ">=0.4,3,1)))"
    Next r
    ws.Range("B37").Formula = "=IFERROR(AVERAGE(K4:K33),""нет данных"")"
    ws.Range("C37").Formula = "=IF(ISNUMBER(B37),IF(B37>=0.7,""высокое"",IF(B37>=0.4,""среднее"",""низкое"")),""" & TODO_TXT & """)"
    ws.Range("B38").Formula = "=IFERROR(COUNTIF(L4:L33,5)/COUNT(L4:L33),""нет данных"")"
    ws.Range("C38").Formula = "=IF(ISNUMBER(B38),IF(B38>=0.5,""высокая доля"",IF(B38>=0.25,""средняя доля"",""низкая доля"")),""" & TODO_TXT & """)"
End Sub

'------------------------------------------------------------------ 05_Отзывы_рейтинги
Private Sub FillReviews(ws As Worksheet)
    Dim comp As Variant, r As Long, i As Long
    ws.Range("A4:K33").ClearContents
    ws.Range("N4:N33").ClearContents
    ws.Range("A4:N33").ClearComments
    comp = Competitors()
    For i = 0 To UBound(comp)
        ws.Cells(4 + i, 1).Value = comp(i)(0)
        ws.Cells(4 + i, 14).Value = "Рейтинги/отзывы — ДОСНЯТЬ"
    Next i
    MarkTodo ws.Range("B4:K21"), "G2 / Capterra / Trustpilot / Google: рейтинг и число отзывов; кейсы на сайте; возраст бренда (год запуска — из «About» или Википедии)"
    For r = 4 To 33
        ws.Cells(r, 12).Formula = "=IF(OR($A" & r & "="""",COUNT($J" & r & ":$K" & r & ")<2,COUNT($B" & r & ",$D" & r & ",$F" & r & ",$H" & r & ")=0),""""," & _
            "MIN(1,AVERAGE(IFERROR(AVERAGE(B" & r & ",D" & r & ",F" & r & ",H" & r & ")/5,0),MIN((C" & r & "+E" & r & "+G" & r & "+I" & r & ")/300,1),MIN(J" & r & "/30,1),MIN(K" & r & "/10,1))))"
        ws.Cells(r, 13).Formula = "=IF(L" & r & "="""","""",IF(L" & r & ">=0.7,5,IF(L" & r & ">=0.4,3,1)))"
    Next r
    ws.Range("B37").Formula = "=IFERROR(AVERAGE(L4:L33),""нет данных"")"
    ws.Range("C37").Formula = "=IF(ISNUMBER(B37),IF(B37>=0.7,""высокий"",IF(B37>=0.4,""средний"",""низкий"")),""" & TODO_TXT & """)"
    ' ИСПРАВЛЕНО: исходная =AVERAGE(C4:C33+E4:E33+...) — формула массива, в обычном вводе даёт #VALUE!
    ws.Range("B38").Formula = "=IF(COUNT(C4:C33,E4:E33,G4:G33,I4:I33)=0,""нет данных"",SUM(C4:C33,E4:E33,G4:G33,I4:I33)/COUNTA(A4:A33))"
    ws.Range("C38").Formula = "=IF(ISNUMBER(B38),IF(B38>='" & SH_PAR & "'!B17,""высокий барьер"",IF(B38>='" & SH_PAR & "'!B17*0.5,""средний барьер"",""низкий барьер"")),""" & TODO_TXT & """)"
End Sub

'------------------------------------------------------------------ 06_Технологии_финансы
Private Sub FillTech(ws As Worksheet)
    Dim comp As Variant, r As Long, i As Long
    ' F — мобильное приложение, G — личный кабинет/аккаунт (по публичным сайтам, проверить)
    Dim appCab As Variant
    appCab = Array( _
        Array("нет", "нет"), Array("нет", "да"), Array("нет", "нет"), Array("нет", "да"), _
        Array("нет", "нет"), Array("нет", "нет"), Array("нет", "нет"), Array("нет", "да"), _
        Array("нет", "да"), Array("да", "да"), Array("да", "да"), Array("нет", "да"), _
        Array("нет", "нет"), Array("нет", "нет"), Array("нет", "нет"), Array("нет", "да"), _
        Array("нет", "да"), Array("нет", "нет"))
    ws.Range("A4:J33").ClearContents
    ws.Range("M4:M33").ClearContents
    ws.Range("A4:M33").ClearComments
    comp = Competitors()
    For i = 0 To UBound(comp)
        r = 4 + i
        ws.Cells(r, 1).Value = comp(i)(0)
        ws.Cells(r, 6).Value = appCab(i)(0)
        ws.Cells(r, 7).Value = appCab(i)(1)
        ws.Cells(r, 13).Value = "F/G — по сайту (проверить); остальное — ДОСНЯТЬ"
    Next i
    MarkTodo ws.Range("B4:E21"), "финансирование (Crunchbase), сотрудники (LinkedIn), стек и CRM 1–5 (BuiltWith/Wappalyzer)"
    MarkTodo ws.Range("H4:J21"), "активные объявления (Ads Transparency / Meta Ad Library), вакансии (сайт, LinkedIn, hh/rabota.by), патенты/уникальные активы (0/1)"
    For r = 4 To 33
        ws.Cells(r, 11).Formula = "=IF(OR($A" & r & "="""",COUNT($B" & r & ":$E" & r & ",$H" & r & ":$J" & r & ")<7),""""," & _
            "MIN(1,AVERAGE(MIN(B" & r & "/500000,1),MIN(C" & r & "/60,1),D" & r & "/5,E" & r & "/5,IF(F" & r & "=""да"",1,0),IF(G" & r & "=""да"",1,0),MIN(H" & r & "/50,1),MIN(I" & r & "/6,1),IF(J" & r & ">0,1,0))))"
        ws.Cells(r, 12).Formula = "=IF(K" & r & "="""","""",IF(K" & r & ">=0.7,5,IF(K" & r & ">=0.4,3,1)))"
    Next r
    ws.Range("B37").Formula = "=IFERROR(AVERAGE(K4:K33),""нет данных"")"
    ws.Range("C37").Formula = "=IF(ISNUMBER(B37),IF(B37>=0.7,""высокая"",IF(B37>=0.4,""средняя"",""низкая"")),""" & TODO_TXT & """)"
    ' ИСПРАВЛЕНО: исходная формула складывала две доли (могла дать > 100 %); теперь — доля игроков, у которых есть кабинет ИЛИ приложение
    ws.Range("B38").Formula = "=IFERROR(SUMPRODUCT((A4:A33<>"""")*(((F4:F33=""да"")+(G4:G33=""да""))>0))/COUNTA(A4:A33),0)"
End Sub

'------------------------------------------------------------------ 07_Барьеры_входа
Private Sub FixBarriers(ws As Worksheet)
    Dim r As Long
    Dim q As String
    q = "'" & SH_IN & "'!"
    ws.Range("J3").Value = "Источник балла"
    ws.Range("J3").Font.Bold = True

    ' D4 — формула шаблона (концентрация всегда считается по 02)
    ws.Range("J4").Value = "расчёт (02: CR3/HHI)"
    ws.Range("D5").Formula = "=IF(COUNT('" & SH_SW & "'!$K$4:$K$33)>0,'" & SH_CH & "'!E5," & q & "$C$5)"
    ws.Range("J5").Formula = "=IF(COUNT('" & SH_SW & "'!$K$4:$K$33)>0,""расчёт"",""экспертно (Ввод_NotaCode)"")"
    ws.Range("D6").Formula = "=IF(AND(COUNT('" & SH_SW & "'!$L$4:$L$33)>0,ISNUMBER('" & SH_SR & "'!$B$37)),IF(C6>=0.6,5,IF(C6>=0.35,3,1))," & q & "$C$6)"
    ws.Range("J6").Formula = "=IF(AND(COUNT('" & SH_SW & "'!$L$4:$L$33)>0,ISNUMBER('" & SH_SR & "'!$B$37)),""расчёт"",""экспертно (Ввод_NotaCode)"")"
    ws.Range("D7").Formula = "=IF(COUNT('" & SH_SW & "'!$J$4:$J$33)>0,'" & SH_CH & "'!E10," & q & "$C$7)"
    ws.Range("J7").Formula = "=IF(COUNT('" & SH_SW & "'!$J$4:$J$33)>0,""расчёт (нижняя граница)"",""экспертно (Ввод_NotaCode)"")"
    For r = 8 To 10
        ws.Range("D" & r).Formula = "=IF(ISNUMBER(C" & r & "),IF(C" & r & ">=0.7,5,IF(C" & r & ">=0.4,3,1))," & q & "$C$" & r & ")"
        ws.Range("J" & r).Formula = "=IF(ISNUMBER(C" & r & "),""расчёт"",""экспертно (Ввод_NotaCode)"")"
    Next r
    ws.Range("C11").Formula = "=" & q & "$B$15"
    ws.Range("C12").Formula = "=" & q & "$B$16"
    ws.Range("J11").Value = "экспертно (Ввод_NotaCode, B15)"
    ws.Range("J12").Value = "экспертно (Ввод_NotaCode, B16)"

    ' ИСПРАВЛЕНО: веса 0,16 + 6*0,12 + 2*0,08 = 1,04 -> нормированы к 1,00
    For r = 4 To 12
        ws.Range("E" & r).Formula = "=" & q & "F" & r
    Next r
    ws.Range("E13").Formula = "=SUM(E4:E12)"
    ws.Range("E4:E13").NumberFormat = "0.000"
    ws.Range("C13").NumberFormat = "0.00"
    ws.Range("F4:F13").NumberFormat = "0.00"
    ws.Range("H13").Value = "Сумма весов приведена к 1,00 (в шаблоне было 1,04 при подписи «1»)."
    ws.Columns("J").ColumnWidth = 30
End Sub

'------------------------------------------------------------------ 10_Портер_NotaCode
Private Sub BuildPorterSheet(wb As Workbook)
    Dim ws As Worksheet, i As Long, v As Variant
    Set ws = GetOrAddSheet(wb, SH_PT, SH_BR)
    ws.Range("A1").Value = "Пять сил Портера с цифровой адаптацией — NotaCode (шкала 1/3/5: 1 низкое, 3 среднее, 5 высокое давление)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:C3").Value = Array("Сила", "Оценка 1/3/5", "Ключевые доказательства")
    v = Array( _
        Array("Конкуренция между игроками", 5, "18 игроков (прямые DSL, GUI, CASE, платформы); CR3 = 99,8 %, HHI = 7590 по 4 доменам Similarweb"), _
        Array("Угроза новых участников", 3, "Запуск дешёвый (open-source парсеры, elkjs, LLM API), но доверие, валидация нотаций и сообщество требуют времени"), _
        Array("Сила покупателей", 5, "Бесплатные PlantUML/Mermaid/draw.io, прозрачные цены, нулевые издержки переключения, студенты чувствительны к цене"), _
        Array("Сила поставщиков", 3, "Google/Яндекс/GitHub (трафик), LLM-провайдеры (смягчено BYOK/заглушкой), бесплатные хостинги, платёжный провайдер РБ"), _
        Array("Угроза заменителей", 5, "draw.io, LLM (ChatGPT пишет PlantUML/Mermaid), Visio/Word, Ramus в вузах, рисование от руки"), _
        Array("Цифровые усилители (сводно)", 3, "Среднее 6 факторов = 2,33 -> ближайшее значение шкалы 3"))
    For i = 0 To 5
        ws.Cells(4 + i, 1).Value = v(i)(0)
        ws.Cells(4 + i, 2).Value = v(i)(1)
        ws.Cells(4 + i, 3).Value = v(i)(2)
    Next i
    ws.Range("A10").Value = "Среднее (5 сил + ЦУ)"
    ws.Range("B10").Formula = "=AVERAGE(B4:B9)"
    ws.Range("C10").Formula = "=IF(B10<=2,""низкое давление"",IF(B10<=3.5,""умеренное давление"",""высокое давление""))"
    ws.Range("A11").Value = "Среднее (только 5 сил)"
    ws.Range("B11").Formula = "=AVERAGE(B4:B8)"
    ws.Range("C11").Formula = "=IF(B11<=2,""низкое давление"",IF(B11<=3.5,""умеренное давление"",""высокое давление""))"

    ws.Range("A13:C13").Value = Array("Цифровой фактор", "Оценка 1/3/5", "Обоснование")
    v = Array( _
        Array("Сетевые эффекты", 3, "Mermaid нативно в GitHub/GitLab/Notion, плагины PlantUML — де-факто стандарты; совместная работа у Miro/Lucid"), _
        Array("Данные", 1, "Данные не дают лидерам решающего преимущества; LLM доступны всем"), _
        Array("Алгоритмическая видимость", 3, "Открытие продукта через выдачу Google/Яндекс, GitHub, маркетплейсы расширений"), _
        Array("Платформенная зависимость", 3, "GitHub, Google Drive, платёжный провайдер; альтернативы есть"), _
        Array("Издержки переключения", 1, "Текстовые форматы и экспорт — переключение лёгкое"), _
        Array("Экосистемная связанность", 3, "Конкуренция с экосистемами: Atlassian+draw.io, GitHub+Mermaid, Microsoft+Visio, Miro"))
    For i = 0 To 5
        ws.Cells(14 + i, 1).Value = v(i)(0)
        ws.Cells(14 + i, 2).Value = v(i)(1)
        ws.Cells(14 + i, 3).Value = v(i)(2)
    Next i
    ws.Range("A20").Value = "Среднее цифровых факторов"
    ws.Range("B20").Formula = "=AVERAGE(B14:B19)"

    ws.Range("A22:C22").Value = Array("Критерий методики Similarweb (1–5)", "Балл", "Обоснование")
    v = Array( _
        Array("Количество релевантных конкурентов", 5, "Более 15 активных игроков, включая сильные бренды"), _
        Array("Концентрация трафика", 5, "Топ-3 = 99,8 % измеренного трафика"), _
        Array("Стоимость входа в каналы", 3, "Смешанная модель: органика, GitHub, вузы + реклама (экспертно)"), _
        Array("Поисковая конкуренция", 4, "Англоязычная выдача занята сильными доменами; русская IDEF/DFD — проверить (диапазон 3–5)"), _
        Array("Рекламная конкуренция", 3, "Нет данных рекламных библиотек — ДОСНЯТЬ"), _
        Array("Репутационный барьер", 5, "Многолетние бренды и сообщества"), _
        Array("Технологический барьер", 4, "Парсер DSL, валидация нотаций, раскладка, версии, экспорт (диапазон 3–5)"), _
        Array("Издержки переключения клиента", 2, "Низкие: текстовые форматы (диапазон 1–3)"))
    For i = 0 To 7
        ws.Cells(23 + i, 1).Value = v(i)(0)
        ws.Cells(23 + i, 2).Value = v(i)(1)
        ws.Cells(23 + i, 3).Value = v(i)(2)
    Next i
    ws.Range("A31").Value = "Среднее 8 критериев"
    ws.Range("B31").Formula = "=AVERAGE(B23:B30)"
    ws.Range("C31").Formula = "=IF(B31<=2,""низкий"",IF(B31<=3,""умеренный"",IF(B31<=4,""высокий"",""очень высокий"")))"
    ws.Range("B4:B31").Interior.Color = RGB(255, 242, 204)
    ws.Range("B10:B11,B20,B31").Interior.Color = RGB(226, 240, 217)
    ws.Range("B10:B31").NumberFormat = "0.00"
    ws.Range("A3:C3,A13:C13,A22:C22").Font.Bold = True
    ws.Columns("A").ColumnWidth = 38
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 110
End Sub

'------------------------------------------------------------------ 08_Дашборд
Private Sub FixDashboard(ws As Worksheet)
    ws.Range("D8").Formula = "=IF(B8>='" & SH_PAR & "'!B13,""Высокая концентрация (HHI)"",IF(B8>='" & SH_PAR & "'!B13*0.65,""Средняя концентрация"",""Низкая концентрация""))"
    ' ИСПРАВЛЕНО: в исходном A15 строковая константа длиннее 255 символов — Excel не открывает такой файл
    ws.Range("A15").Formula = "=""По результатам оценки цифровой конкуренции индекс барьеров входа составляет ""&TEXT(B5,""0.00"")&" & _
        """ из 5, что соответствует уровню: ""&B6&"". Наиболее значимые факторы давления — на листе 07_Барьеры_входа (колонка J — источник балла).""&" & _
        """ Данные Similarweb и альтернативных источников — оценочная база и требуют проверки фактическими коммерческими данными."""
    ws.Range("A12").Value = "Среднее давление 5 сил + ЦУ"
    ws.Range("B12").Formula = "='" & SH_PT & "'!B10"
    ws.Range("C12").Value = "1-5"
    ws.Range("D12").Formula = "='" & SH_PT & "'!C10"
    ws.Range("A13").Value = "Сумма весов факторов"
    ws.Range("B13").Formula = "='" & SH_BR & "'!E13"
    ws.Range("A17").Value = "Вывод NotaCode: давление высокое, барьеры средние (у верхней границы); вход — нишевой: " & _
        "строгие нотации (IDEF0/IDEF3/DFD/сети Петри) с проверкой правил, русскоязычные студенты и преподаватели РБ, " & _
        "совместимость (экспорт в PlantUML/Mermaid/draw.io), контент и вузовские каналы вместо платной рекламы."
    ws.Range("A17").WrapText = False
End Sub

'------------------------------------------------------------------ 09_Источники
Private Sub FillSources(ws As Worksheet)
    Dim r As Long
    ws.Range("F4").Value = "17.09.2026; 23.09.2026"
    For r = 5 To 19
        ws.Cells(r, 6).Value = TODO_TXT & " (не использовался)"
        ws.Cells(r, 6).Interior.Color = RGB(255, 204, 153)
    Next r
    ws.Range("A21:G21").Value = Array("Источник (доп., NotaCode)", "URL", "Что взято", "Для какой силы/фактора", "Формат фиксации", "Дата обращения", "Ограничения")
    ws.Range("A21:G21").Font.Bold = True
    ws.Range("A22:G22").Value = Array("Google Trends", "https://trends.google.com/", "PlantUML ср. 41, text to diagram ср. 26, Беларусь #5 по PlantUML", "покупатели, заменители, спрос", "CSV + скриншот (ЛР1)", "17.09.2026", "относительный индекс 0–100")
    ws.Range("A23:G23").Value = Array("Официальные сайты конкурентов", "plantuml.com, eraser.io, d2lang.com, mermaidchart.com", "цены, модель доходов, функции", "конкуренция, покупатели", "таблица (ЛР5)", "17.09.2026; 23.09.2026", "декларации компаний")
    ws.Range("A24:G24").Value = Array("Документация NotaCode", "E:\NotaCode\docs\business", "сегменты, тарифы Free/Pro, ограничения", "поставщики, границы рынка", "md", "25.09.2026", "внутренние допущения")
    ws.Range("A25:G25").Value = Array("GitHub", "https://github.com/", "звёзды/активность PlantUML, Mermaid, D2, bpmn.io", "заменители, новые участники", "таблица", TODO_TXT, "не равно числу пользователей")
    ws.Range("A26:G26").Value = Array("Яндекс Вордстат", "https://wordstat.yandex.ru/", "частотность запросов листа 04 (РБ, РФ)", "покупатели, поиск/реклама", "скриншот/таблица", TODO_TXT, "нужна авторизация Яндекс ID")
    ws.Range("F25:F26").Interior.Color = RGB(255, 204, 153)
End Sub
