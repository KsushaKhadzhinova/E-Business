Attribute VB_Name = "Oformlenie_BGUIR"
' =====================================================================
'  ЛР7 «Итоговый отчёт по бизнес-анализу» (NotaCode) — Word-макрос
'  оформления документа по требованиям БГУИР (СТП 01-2024, упрощённо).
'
'  Что делает макрос Main для АКТИВНОГО документа .docx:
'    1) поля страницы: левое 30 мм, правое 15 мм, верхнее 20 мм, нижнее 20 мм;
'       формат A4, книжная ориентация;
'    2) основной текст: Times New Roman 14 пт, выравнивание по ширине,
'       абзацный отступ 12,5 мм, интервал 1,0 или 1,5 (спрашивает при запуске),
'       интервалы до/после абзаца 0;
'    3) стили заголовков 1-3: Times New Roman 14 пт полужирный, чёрный;
'       Заголовок 1 — с новой страницы, ПРОПИСНЫМИ; Заголовок 2-3 — с отступа;
'    4) подписи: абзацы, начинающиеся с «Рисунок N», — по центру,
'       «Таблица N» — по левому краю без отступа (над таблицей);
'       абзацы, содержащие только картинку, — по центру;
'    5) таблицы: шрифт 12 пт, интервал 1,0, без абзацного отступа, границы,
'       повтор строки заголовка на новой странице;
'    6) нумерация страниц: внизу справа, на титульном листе не показывается;
'    7) автооглавление «СОДЕРЖАНИЕ»: вставляется перед первым «Заголовком 1»
'       (или обновляется, если уже есть), уровни 1-3, с точками-заполнителями.
'
'  Как запустить:
'    1) Сконвертируйте ОТЧЕТ.md в .docx, например:
'         pandoc ОТЧЕТ.md -o ОТЧЕТ.docx --resource-path=.
'       (или откройте md через Word и сохраните как .docx);
'    2) откройте ОТЧЕТ.docx в Word -> Alt+F11 -> File -> Import File ->
'       выберите этот файл Оформление_БГУИР.bas;
'    3) вернитесь в документ -> Alt+F8 -> Main -> «Выполнить»;
'    4) проверьте титульный лист и оглавление, сохраните документ.
'  Файл сохранён в кодировке Windows-1251 (так его читает редактор VBA).
' =====================================================================
Option Explicit

Private Const FONT_NAME As String = "Times New Roman"
Private Const FONT_SIZE As Single = 14
Private Const TABLE_FONT_SIZE As Single = 12
Private Const MARGIN_LEFT_MM As Single = 30
Private Const MARGIN_RIGHT_MM As Single = 15
Private Const MARGIN_TOP_MM As Single = 20
Private Const MARGIN_BOTTOM_MM As Single = 20
Private Const FIRST_INDENT_MM As Single = 12.5

Public Sub Main()
    Dim doc As Document
    Dim spacing As Single
    Dim answer As VbMsgBoxResult

    If Documents.Count = 0 Then
        MsgBox "Откройте документ .docx перед запуском макроса.", vbExclamation
        Exit Sub
    End If
    Set doc = ActiveDocument

    answer = MsgBox("Межстрочный интервал 1,5?" & vbCrLf & _
                    "Да — 1,5;  Нет — 1,0 (одинарный).", vbYesNoCancel + vbQuestion, _
                    "Оформление БГУИР")
    If answer = vbCancel Then Exit Sub
    If answer = vbYes Then spacing = 1.5 Else spacing = 1

    Application.ScreenUpdating = False
    SetupPage doc
    SetupBodyStyles doc, spacing
    SetupHeadingStyles doc, spacing
    ApplyDirectBodyFormat doc, spacing
    FormatCaptionsAndPictures doc
    FormatTables doc
    AddPageNumbers doc
    InsertOrUpdateToc doc
    Application.ScreenUpdating = True

    MsgBox "Оформление применено: поля 30/15/20/20 мм, " & FONT_NAME & " " & FONT_SIZE & _
           " пт, интервал " & Format(spacing, "0.0") & ", нумерация страниц, оглавление.", _
           vbInformation, "Оформление БГУИР"
End Sub

Private Function LineRule(ByVal spacing As Single) As WdLineSpacing
    If spacing = 1.5 Then
        LineRule = wdLineSpace1pt5
    Else
        LineRule = wdLineSpaceSingle
    End If
End Function

Private Sub SetupPage(ByVal doc As Document)
    Dim sec As Section
    For Each sec In doc.Sections
        With sec.PageSetup
            .PaperSize = wdPaperA4
            .Orientation = wdOrientPortrait
            .LeftMargin = MillimetersToPoints(MARGIN_LEFT_MM)
            .RightMargin = MillimetersToPoints(MARGIN_RIGHT_MM)
            .TopMargin = MillimetersToPoints(MARGIN_TOP_MM)
            .BottomMargin = MillimetersToPoints(MARGIN_BOTTOM_MM)
            .HeaderDistance = MillimetersToPoints(10)
            .FooterDistance = MillimetersToPoints(10)
            .DifferentFirstPageHeaderFooter = True
        End With
    Next sec
End Sub

Private Sub StyleFont(ByVal st As Style)
    On Error Resume Next
    With st.Font
        .Name = FONT_NAME
        .NameAscii = FONT_NAME
        .NameOther = FONT_NAME
        .Size = FONT_SIZE
        .Color = wdColorAutomatic
    End With
    On Error GoTo 0
End Sub

Private Sub StyleParagraph(ByVal st As Style, ByVal spacing As Single)
    On Error Resume Next
    With st.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = MillimetersToPoints(FIRST_INDENT_MM)
        .LeftIndent = 0
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = LineRule(spacing)
    End With
    On Error GoTo 0
End Sub

Private Sub SetupBodyStyles(ByVal doc As Document, ByVal spacing As Single)
    Dim names As Variant
    Dim i As Long
    Dim st As Style

    StyleFont doc.Styles(wdStyleNormal)
    StyleParagraph doc.Styles(wdStyleNormal), spacing

    ' стили, которые создаёт pandoc при конвертации из markdown
    names = Array("Body Text", "First Paragraph", "Compact", "Block Text", _
                  "Основной текст", "List Paragraph", "Абзац списка")
    For i = LBound(names) To UBound(names)
        Set st = Nothing
        On Error Resume Next
        Set st = doc.Styles(names(i))
        On Error GoTo 0
        If Not st Is Nothing Then
            StyleFont st
            StyleParagraph st, spacing
        End If
    Next i
End Sub

Private Sub SetupHeadingStyles(ByVal doc As Document, ByVal spacing As Single)
    Dim lvl As Long
    Dim st As Style
    For lvl = 1 To 3
        Select Case lvl
            Case 1: Set st = doc.Styles(wdStyleHeading1)
            Case 2: Set st = doc.Styles(wdStyleHeading2)
            Case Else: Set st = doc.Styles(wdStyleHeading3)
        End Select
        StyleFont st
        With st.Font
            .Bold = True
            .Italic = False
            .AllCaps = (lvl = 1)
        End With
        With st.ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .FirstLineIndent = MillimetersToPoints(FIRST_INDENT_MM)
            .LeftIndent = 0
            If lvl = 1 Then .SpaceBefore = 0 Else .SpaceBefore = 12
            .SpaceAfter = 12
            .LineSpacingRule = LineRule(spacing)
            .KeepWithNext = True
            .PageBreakBefore = (lvl = 1)
        End With
    Next lvl
End Sub

Private Function IsHeading(ByVal p As Paragraph) As Boolean
    IsHeading = (p.OutlineLevel <> wdOutlineLevelBodyText)
End Function

Private Sub ApplyDirectBodyFormat(ByVal doc As Document, ByVal spacing As Single)
    ' снимает «ручное» оформление, оставшееся от конвертации (другие шрифты, размеры)
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        If Not p.Range.Information(wdWithInTable) Then
            If Not IsHeading(p) Then
                p.Range.Font.Name = FONT_NAME
                p.Range.Font.Size = FONT_SIZE
                p.LineSpacingRule = LineRule(spacing)
                p.SpaceBefore = 0
                p.SpaceAfter = 0
            End If
        End If
    Next p
End Sub

Private Sub FormatCaptionsAndPictures(ByVal doc As Document)
    Dim p As Paragraph
    Dim t As String
    For Each p In doc.Paragraphs
        If Not p.Range.Information(wdWithInTable) Then
            t = Trim$(Replace(p.Range.Text, vbCr, ""))
            If p.Range.InlineShapes.Count > 0 And Len(t) <= 1 Then
                p.Alignment = wdAlignParagraphCenter
                p.FirstLineIndent = 0
                p.KeepWithNext = True
                p.SpaceBefore = 6
            ElseIf t Like "Рисунок [0-9]*" Then
                p.Alignment = wdAlignParagraphCenter
                p.FirstLineIndent = 0
                p.SpaceAfter = 12
                p.Range.Font.Bold = False
            ElseIf t Like "Таблица [0-9]*" Then
                p.Alignment = wdAlignParagraphLeft
                p.FirstLineIndent = 0
                p.KeepWithNext = True
                p.SpaceBefore = 6
                p.Range.Font.Bold = False
            End If
        End If
    Next p
End Sub

Private Sub FormatTables(ByVal doc As Document)
    Dim tbl As Table
    For Each tbl In doc.Tables
        With tbl.Range
            .Font.Name = FONT_NAME
            .Font.Size = TABLE_FONT_SIZE
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With
        tbl.Borders.Enable = True
        tbl.PreferredWidthType = wdPreferredWidthPercent
        tbl.PreferredWidth = 100
        On Error Resume Next
        tbl.Rows(1).HeadingFormat = True
        tbl.Rows(1).Range.Font.Bold = True
        tbl.Rows.AllowBreakAcrossPages = True
        On Error GoTo 0
    Next tbl
End Sub

Private Sub AddPageNumbers(ByVal doc As Document)
    Dim sec As Section
    Dim rng As Range
    For Each sec In doc.Sections
        Set rng = sec.Footers(wdHeaderFooterPrimary).Range
        rng.Text = ""
        rng.Fields.Add rng, wdFieldPage
        With sec.Footers(wdHeaderFooterPrimary).Range
            .ParagraphFormat.Alignment = wdAlignParagraphRight
            .ParagraphFormat.FirstLineIndent = 0
            .Font.Name = FONT_NAME
            .Font.Size = FONT_SIZE
        End With
        sec.Footers(wdHeaderFooterFirstPage).Range.Text = ""
    Next sec
End Sub

Private Sub InsertOrUpdateToc(ByVal doc As Document)
    Dim p As Paragraph
    Dim rng As Range
    Dim tocRng As Range

    If doc.TablesOfContents.Count > 0 Then
        doc.TablesOfContents(1).Update
        Exit Sub
    End If

    For Each p In doc.Paragraphs
        If p.OutlineLevel = wdOutlineLevel1 Then
            Set rng = p.Range
            Exit For
        End If
    Next p
    If rng Is Nothing Then Exit Sub

    ' два новых абзаца перед первым заголовком: «СОДЕРЖАНИЕ» и место под оглавление
    rng.InsertParagraphBefore
    rng.InsertParagraphBefore
    Set rng = rng.Paragraphs(1).Range
    rng.Style = doc.Styles(wdStyleNormal)
    rng.InsertBefore "СОДЕРЖАНИЕ"
    With rng.Paragraphs(1)
        .Alignment = wdAlignParagraphCenter
        .FirstLineIndent = 0
        .PageBreakBefore = True
        .SpaceAfter = 12
    End With
    rng.Paragraphs(1).Range.Font.Bold = True

    Set tocRng = rng.Paragraphs(1).Next.Range
    tocRng.Style = doc.Styles(wdStyleNormal)
    tocRng.Collapse wdCollapseStart
    doc.TablesOfContents.Add Range:=tocRng, UseHeadingStyles:=True, _
        UpperHeadingLevel:=1, LowerHeadingLevel:=3, IncludePageNumbers:=True, _
        RightAlignPageNumbers:=True, UseHyperlinks:=True
    doc.TablesOfContents(1).TabLeader = wdTabLeaderDots
    doc.TablesOfContents(1).Update
End Sub
