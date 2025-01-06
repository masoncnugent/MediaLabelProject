Imports Microsoft.Office.Interop.Word
Imports Word = Microsoft.Office.Interop.Word

Module Module1

    Public IntPageInd As Integer = 1
    Public IntColumnInd As Integer = 1

    Dim StCurrentText As String
    Dim StCurrentMediaSpe As String
    Dim DecCurrentLabelFontSize As Decimal

    'Public for Form1.Instance_Load()
    Public WithEvents WordApp As Word.Application = New Word.Application()
    Public WordDoc As Word.Document = WordApp.Documents.Add()

    Private WithEvents WordDocClosedTimer As System.Timers.Timer = New System.Timers.Timer()

    Dim EndOfDocRange As Word.Range

    'If you have a printer that isn't 'Microsoft Print To PDF' it gets pinged when margins are changed to make sure what's displayed is accurate with what will be printed. For this reason, I save the default printer as OriginalPrinter, work with 'Microsoft Print To PDF,' and set it back to the default printer by the end.
    Public OriginalPrinter As String = WordApp.ActivePrinter

    Public HiddenTemplateWordApp As Word.Application = New Word.Application()
    Public HiddenTemplateWordDoc As Word.Document = HiddenTemplateWordApp.Documents.Add()

    Public WordTableUserFirst As Table = Nothing
    Public HiddenWordTableTemplate As Table = Nothing



    'Handles when the Word Document is closed by freeing up all resources once closing has finished. The reason for waiting for the Word document to close is that if you try to close it during WordApp.DocumentBeforeClose everything breaks. So if you close Form1 and leave both Word Docs they should both close via CleanupWordApp(), but if you close the visible WordApp it closes after WordApp_Closed, while HiddenTemplateWordApp closes in CleanupWordApp()
    Public Sub WordApp_Closed(ByVal Doc As Word.Document, ByRef Cancel As Boolean) Handles WordApp.DocumentBeforeClose

        'To avoid prompts of saving the documents
        WordDoc.Saved = True
        HiddenTemplateWordDoc.Saved = True

        WordApp.ActivePrinter = OriginalPrinter

        WordDocClosedTimer.Interval = 1000
        'Test
        WordDocClosedTimer.AutoReset = False
        WordDocClosedTimer.Start()
    End Sub



    'The 'ticking' of the timer should allow WordApp to close fully before we attempt to run CleanupWordApp(), so that it only attempts closing HiddenTemplateWordApp
    Private Sub WordDocClosedTimer_Tick(sender As Object, e As EventArgs) Handles WordDocClosedTimer.Elapsed
        WordDocClosedTimer.Stop()

        Form1.Instance.Invoke(Sub() Form1.Instance.Close())

    End Sub



    'This is used to prepare the Word document used by the user as well as the template document. When pages are added to WordApp the tables in them are pasted from the template document's table
    'Why not just paste a saved table that's empty but not in any WordApp directly? Word doesn't allow this... so yes, this is a dumb workaround, but so long as HiddenTemplateWordApp is closed every time it is at least functional
    Public Sub PrepareWordDocument(ByRef WdApp As Word.Application, ByRef WdDoc As Word.Document, ByRef WdTable As Word.Table)

        WdApp.Options.SaveInterval = 0

        WdApp.ActivePrinter = "Microsoft Print to PDF"
        WdDoc.PageSetup.HeaderDistance = 0
        WdDoc.PageSetup.FooterDistance = 0
        WdDoc.PageSetup.TopMargin = WdApp.InchesToPoints(0.5)
        WdDoc.PageSetup.LeftMargin = WdApp.InchesToPoints(0.28)
        WdDoc.PageSetup.BottomMargin = 0
        WdDoc.PageSetup.RightMargin = WdApp.InchesToPoints(0.28)

        '20 rows, 7 columns (3 are going to be for the blank space in between labels)
        WdTable = WdDoc.Tables.Add(WdDoc.Range(0, 0), 20, 7)

        With WdTable
            .TopPadding = 0
            .BottomPadding = 0
            .LeftPadding = 0
            .RightPadding = 0
            .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Range.ParagraphFormat.LeftIndent = WordApp.InchesToPoints(0)
            .Range.ParagraphFormat.RightIndent = WordApp.InchesToPoints(0)
            .Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle
            .Range.ParagraphFormat.SpaceAfter = 0
            .Range.Font.Name = "Times New Roman"
            .Range.Font.Bold = True
            .Columns(1).SetWidth(WdApp.InchesToPoints(1.75), Word.WdRulerStyle.wdAdjustNone)
            .Columns(2).SetWidth(WdApp.InchesToPoints(0.3), Word.WdRulerStyle.wdAdjustNone)
            .Columns(3).SetWidth(WdApp.InchesToPoints(1.75), Word.WdRulerStyle.wdAdjustNone)
            .Columns(4).SetWidth(WdApp.InchesToPoints(0.3), Word.WdRulerStyle.wdAdjustNone)
            .Columns(5).SetWidth(WdApp.InchesToPoints(1.75), Word.WdRulerStyle.wdAdjustNone)
            .Columns(6).SetWidth(WdApp.InchesToPoints(0.3), Word.WdRulerStyle.wdAdjustNone)
            .Columns(7).SetWidth(WdApp.InchesToPoints(1.75), Word.WdRulerStyle.wdAdjustNone)
            .Rows.SetHeight(WdApp.InchesToPoints(0.5), Word.WdRowHeightRule.wdRowHeightExactly)
        End With

        'Keeps the table from going across pages
        WdTable.AllowPageBreaks = False

    End Sub



    'Adds the details of LabelMaker() to the Word document for the specific columns
    Public Sub FinalizeWordDocument(TplLabelInfo)

        StCurrentText = TplLabelInfo.Item1
        DecCurrentLabelFontSize = TplLabelInfo.Item2
        StCurrentMediaSpe = TplLabelInfo.Item3
        Dim IntColumnCount As Integer = CInt(TplLabelInfo.Item4)


        Dim TplColumnRange = DetermineLabelColumnRange(IntColumnCount)

        ChooseColumns(TplColumnRange.Item1, TplColumnRange.Item2, TplColumnRange.Item3)

        'Disables the text of "Your margins are pretty small. Some of your content might be cut off when you print. Do you still want to print?"
        WordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone

        WordApp.ActivePrinter = OriginalPrinter

        'Reenables warnings in case any legitimate ones should arise following printing
        WordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll
    End Sub



    'Uses simple logic to increment IntColumnEndInd and IntPageEndId based on IntColumnCount
    Private Function DetermineLabelColumnRange(IntColumnCount As Integer) As (Integer, Integer, Integer)

        Dim IntColumnStartInd As Integer = IntColumnInd
        Dim IntColumnEndInd As Integer = IntColumnStartInd
        Dim IntPageStartId As Integer = IntPageInd
        Dim IntPageEndId As Integer = IntPageStartId

        If IntColumnEndInd > 7 Then
            AddWordPage()
            IntColumnStartInd = 1
            IntColumnEndInd = 1
            IntPageStartId += 1
            IntPageEndId += 1
        End If

        ' -1 Because you're starting on the first new column by default
        For i As Integer = 1 To IntColumnCount - 1

            IntColumnEndInd += 2
            If IntColumnEndInd > 7 Then
                'This will make a blank page after 2 2000s lots
                AddWordPage()
                IntColumnEndInd = 1
                IntPageEndId += 1
            End If
        Next


        IntPageInd = IntPageEndId
        'This allows the next labels made to start on a different column
        IntColumnInd = IntColumnEndInd + 2

        Return (IntColumnStartInd, IntPageStartId, IntColumnCount)

    End Function



    'Feeds FillOutColumnCells() with data for which tables and columns should be filled out
    Private Sub ChooseColumns(IntColumnStartInd As Integer, IntPageStartId As Integer, IntColumnCount As Integer)

        SafeUpdateProgressBarMaxixmum(IntColumnCount)

        Dim IntCurrentPageId As Integer = IntPageStartId
        Dim IntCurrentColumnId As Integer = IntColumnStartInd

        For i As Integer = 1 To IntColumnCount

            FillOutColumnCells(IntCurrentPageId, IntCurrentColumnId)

            SafeUpdateProgressBarValue(i)

            IntCurrentColumnId += 2
            If IntCurrentColumnId > 7 Then
                IntCurrentColumnId = 1
                IntCurrentPageId += 1
            End If

        Next



    End Sub



    'This and the following function invoke the UI thread to make changes to the progress bar, as threads other than the UI thread cannot do so
    Public Sub SafeUpdateProgressBarMaxixmum(Maximum As Integer)
        'Checks if Form1.Instance exists, is not disposed, and has a handle
        If Form1.Instance IsNot Nothing AndAlso Not Form1.Instance.IsDisposed AndAlso Form1.Instance.IsHandleCreated Then
            'Uses BeginInvoke to perform the UI update asynchronously on the UI thread. The thread created to run Module1 code can't do this unless the main UI thread is invoked
            Form1.Instance.BeginInvoke(Sub()
                                           'Ensures the progress bar can be accessed and updated
                                           If Form1.Instance.ProgressBarLabels IsNot Nothing AndAlso Not Form1.Instance.ProgressBarLabels.IsDisposed Then
                                               Form1.Instance.ProgressBarLabels.Maximum = Maximum
                                           End If
                                       End Sub)
        End If
    End Sub



    Public Sub SafeUpdateProgressBarValue(Value As Integer)
        'Checks if Form1.Instance exists, is not disposed, and has a handle
        If Form1.Instance IsNot Nothing AndAlso Not Form1.Instance.IsDisposed AndAlso Form1.Instance.IsHandleCreated Then
            'Uses BeginInvoke to perform the UI update asynchronously on the UI thread. The thread created to run Module1 code can't do this unless the main UI thread is invoked
            Form1.Instance.BeginInvoke(Sub()
                                           'Ensures the progress bar can be accessed and updated
                                           If Form1.Instance.ProgressBarLabels IsNot Nothing AndAlso Not Form1.Instance.ProgressBarLabels.IsDisposed Then
                                               Form1.Instance.ProgressBarLabels.Value = Value
                                           End If
                                       End Sub)
        End If
    End Sub



    'Adds a new Word page and copies the template table to the new page
    Public Sub AddWordPage()

        HiddenWordTableTemplate.Range.Copy()

        EndOfDocRange = WordDoc.Range
        EndOfDocRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        EndOfDocRange.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
        EndOfDocRange.Paste()

    End Sub



    'Fills out the cells of individual columns, fed by ChooseColumns()
    Public Sub FillOutColumnCells(IntTableId As Integer, IntColIndex As Integer)

        Dim CurrentWordTable As Table = WordDoc.Tables(IntTableId)

        Dim WordColumn As Word.Column = CurrentWordTable.Columns(IntColIndex)

        With WordColumn.Borders
            .InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            .OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle
            .OutsideLineWidth = Word.WdLineWidth.wdLineWidth075pt
            .OutsideColor = Word.WdColor.wdColorBlack
        End With

        For Each Cell As Word.Cell In WordColumn.Cells
            Cell.Range.Text = StCurrentText
            Cell.Range.Font.Size = DecCurrentLabelFontSize
            Cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
        Next

        'Applies shading based on StMediaSpeTemp
        Select Case StCurrentMediaSpe
            Case "S"
                WordColumn.Shading.BackgroundPatternColor = RGB(255, 255, 0)
            Case "M"
                WordColumn.Shading.BackgroundPatternColor = RGB(197, 224, 179)
            Case "CaCl2"
                WordColumn.Shading.BackgroundPatternColor = RGB(156, 194, 229)
            Case Else
                WordColumn.Shading.BackgroundPatternColor = RGB(255, 255, 255)
        End Select
    End Sub


End Module
