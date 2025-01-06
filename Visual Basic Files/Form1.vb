Imports Word = Microsoft.Office.Interop.Word

'Variable prefix conventions
'Str = String
'Int = Integer
'Sin = Single
'Tpl = Tuple
'Dec = Decimal

'Non-VB data types have custom names with no prefixes


Public Class Form1

    'Singleton instance of Form1, used so Form1 can be accessed from background threads properly
    Private Shared _instance As Form1

    'Getter for the singleton instance (idk how this works, it just does)
    Public Shared ReadOnly Property Instance As Form1
        Get
            Return _instance
        End Get
    End Property

    Dim StAutoclaveID As String = ""
    Dim StRunNumber As String = ""
    Dim StVolume As String = ""
    Dim StMediaType As String = ""
    Dim StMediaSpe As String = ""
    Dim StRepeat As String = ""

    Dim StLabelTextL1 As String = ""
    Dim StLabelTextL2 As String = ""
    Dim StLabelTextL3 As String = ""

    Dim BoolUserEditingColumnCount As Boolean = False

    Dim DateToday = DateTime.Now
    Dim DateExp = DateTime.Now

    Dim LoadWordTask As Threading.Tasks.Task
    Dim LoadHiddenTemplateWordTask As Threading.Tasks.Task
    Dim ButtonSubmitTask As Threading.Tasks.Task



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Made so Module1 can access Form1 properly on its non-UI ButtonSubmitTask thread
        _instance = Me

        ButtonSubmit.Enabled = False
        LoadWordTask = Threading.Tasks.Task.Run(Sub() Module1.PrepareWordDocument(WordApp, WordDoc, WordTableUserFirst))

        LoadHiddenTemplateWordTask = Threading.Tasks.Task.Run(Sub() Module1.PrepareWordDocument(HiddenTemplateWordApp, HiddenTemplateWordDoc, HiddenWordTableTemplate))

    End Sub



    'Special code for if the Form is closed directly
    Public Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        CleanupWordApp()

    End Sub



    'This should get rid of excess accumulating Word Documents, along with the setting of saveinterval for each Word App to 0 in Module1.PrepareWordDocument
    Public Sub CleanupWordApp()
        Try
            If WordApp IsNot Nothing AndAlso HiddenTemplateWordApp IsNot Nothing Then
                WordApp.ActivePrinter = OriginalPrinter
                WordApp.Quit(False) 'False = Don't save changes
                HiddenTemplateWordApp.Quit(False) 'False = Don't save changes
                System.Runtime.InteropServices.Marshal.ReleaseComObject(WordApp)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(HiddenTemplateWordApp)
                WordApp = Nothing
                HiddenTemplateWordApp = Nothing
            End If


            'This 'Catch' could still fail LOL
        Catch ex As Exception
            If HiddenTemplateWordApp IsNot Nothing Then
                HiddenTemplateWordApp.Quit(False) 'False = Don't save changes
                System.Runtime.InteropServices.Marshal.ReleaseComObject(HiddenTemplateWordApp)
                HiddenTemplateWordApp = Nothing
            End If
        End Try
    End Sub


    'Words with Module1 to fill out the Word Document
    Private Async Sub ButtonSubmit_Click(sender As Object, e As EventArgs) Handles ButtonSubmit.Click

        ButtonSubmit.Enabled = False

        WordApp.Visible = False
        WordApp.ScreenUpdating = False

        'This is bundled as a tuple ahead of time so that ButtonSubmitTask doesn't find the form reset before it gets to it, causing errors when it would read from it
        Dim TplLabelInfo = LabelMaker()

        ButtonSubmitTask = Threading.Tasks.Task.Run(Sub() Module1.FinalizeWordDocument(TplLabelInfo))

        If CheckBoxMultiple.Checked Then
            ResetForm1()
            Await ButtonSubmitTask
            ProgressBarLabels.Value = 0
            ButtonSubmit.Enabled = True
        Else
            DisableForm1Input()
            Await ButtonSubmitTask
            ProgressBarLabels.Value = 0
            Me.Visible = False
            WordApp.ActivePrinter = OriginalPrinter
            WordApp.ScreenUpdating = True
            WordApp.Visible = True
            If CheckBoxPrint.Checked Then
                WordApp.PrintOut()
            End If
        End If

    End Sub



    'Disables all controls on Form1, but does not reset what has been inputted into each field
    Private Sub DisableForm1Input()
        CheckBoxGravityRun.Enabled = False
        CheckBoxMultiple.Enabled = False
        TextBoxRunNumber.Enabled = False
        ComboBoxVolume.Enabled = False
        ComboBoxMedia.Enabled = False
        ComboBoxMediaType.Enabled = False
        TextBoxRepeat.Enabled = False
        CheckedListBoxAutoclaveIDs.Enabled = False
        TextBoxColumns.Enabled = False
        CheckBoxPrint.Enabled = False

        LabelMaker()

    End Sub





    'Clears all information entered into Form1
    Private Sub ResetForm1()
        CheckBoxGravityRun.Checked = False
        CheckBoxMultiple.Checked = False
        TextBoxRunNumber.Text = ""
        ComboBoxVolume.Text = ""
        ComboBoxMedia.Text = ""
        ComboBoxMediaType.Text = ""
        TextBoxRepeat.Text = ""

        For i As Integer = 0 To CheckedListBoxAutoclaveIDs.Items.Count - 1
            CheckedListBoxAutoclaveIDs.SetItemChecked(i, False)
        Next

        CheckedListBoxAutoclaveIDs.ClearSelected()

        LabelMaker()

    End Sub



    'Determines how many labels each media volume / media volume + type needs
    Public Sub LabelColumnCountDeterminer()

        If Not BoolUserEditingColumnCount Then
            Select Case StVolume
                Case Is = "3000 ML"
                    TextBoxColumns.Text = 2
                Case Is = "2000 ML"
                    TextBoxColumns.Text = 2
                Case Is = "600 ML"
                    TextBoxColumns.Text = 4
                Case Is = "400 ML"
                    TextBoxColumns.Text = 7
                Case Is = "100 ML"
                    TextBoxColumns.Text = 15
            '1000s could be jars or bottles
                Case Is = "1000 ML"
                    Select Case StMediaType
                        Case Is = "SCDB"
                            TextBoxColumns.Text = 2
                        Case Is = "FTM"
                            TextBoxColumns.Text = 2
                            'All other types should be bottles
                        Case Is = "SMA"
                            TextBoxColumns.Text = 1
                        Case Is = "PBS"
                            TextBoxColumns.Text = 1
                        Case Is = "TR-X"
                            TextBoxColumns.Text = 1
                        Case Is = "MSA"
                            TextBoxColumns.Text = 1
                        Case Is = "SP80"
                            TextBoxColumns.Text = 1
                        Case Is = ""
                            TextBoxColumns.Text = ""
                        Case Else
                            TextBoxColumns.Text = 3
                    End Select
                    'Should make no labels if its a invalid media
                Case Else
                    TextBoxColumns.Text = ""
            End Select
        End If

    End Sub



    'Makes the label Preview window
    Public Function LabelMaker() As (String, Decimal, String, String)

        'This will be overwritten by CenterTextInLabel()
        LabelPreview.Padding = New Padding(0, 0, 0, 0)

        'This will resize the label if it would have been made smaller by a previous input. 15 is likely generous but this isn't too much extra code being run
        LabelPreview.Font = New System.Drawing.Font(LabelPreview.Font.FontFamily, 15, LabelPreview.Font.Style)

        'If the Preview autosizes to be larger than a label would be (based on DPI) we can lower the font until the size of the measured label text is at or below the size of a label
        LabelPreview.AutoSize = True

        Dim StDateTodayTemp As String = DateTimeToString(DateToday)
        Dim StDateExpTemp As String = DateTimeToString(DateExp)

        If StDateExpTemp = StDateTodayTemp Then
            StDateExpTemp = "XXXXXX"
        End If

        Dim StAutoclaveIDTemp As String
        Dim StRunNumberTemp As String
        Dim StVolumeTemp As String
        Dim StMediaTypeTemp As String

        'Adds a space if the media type is special
        Dim StMediaSpeSpace As String = ""
        If StMediaSpe IsNot "" Then
            StMediaSpeSpace = " "
        End If

        'Gives all the local variables 'X' Values if there is nothing applicable for them
        If StAutoclaveID = "" Then
            StAutoclaveIDTemp = "XXXXXX"
        Else
            StAutoclaveIDTemp = StAutoclaveID
        End If

        If StRunNumber = "" Then
            StRunNumberTemp = "-X"
        Else
            StRunNumberTemp = StRunNumber
        End If

        If StVolume = "" Then
            StVolumeTemp = "XXXX ML"
        Else
            StVolumeTemp = StVolume
        End If

        If StMediaType = "" Then
            StMediaTypeTemp = "XXXXXXX"
        Else
            StMediaTypeTemp = StMediaType
        End If

        'Gravity run formatting
        If CheckBoxGravityRun.Checked Then
            StLabelTextL1 = StAutoclaveIDTemp & StRunNumberTemp & " STERILIZED"
            StLabelTextL2 = "Preparation: " & StDateTodayTemp
            StLabelTextL3 = "Expiration: " & StDateExpTemp

            'Formatting for all other liquid runs
        Else
            StLabelTextL1 = StAutoclaveIDTemp & StRunNumberTemp & " " & StVolumeTemp & " " & StMediaSpe & StMediaSpeSpace & StMediaTypeTemp
            StLabelTextL2 = "Preparation: " & StDateTodayTemp & StRepeat & StMediaSpe
            StLabelTextL3 = "Expiration: " & StDateExpTemp
        End If

        LabelPreview.Text = StLabelTextL1 & Environment.NewLine & StLabelTextL2 & Environment.NewLine & StLabelTextL3

        LabelPreview.TextAlign = ContentAlignment.MiddleCenter

        Dim DecLabelFontSize As Decimal = LabelPreview.Font.Size

        Do
            If LabelPreview.Width > 168 - 4 Or LabelPreview.Height > 48 - 2 Then
                DecLabelFontSize -= 0.5
                LabelPreview.Font = New System.Drawing.Font(LabelPreview.Font.FontFamily, DecLabelFontSize, LabelPreview.Font.Style)
            Else
                Exit Do
            End If

        Loop While DecLabelFontSize > 1

        'Should color the Label for S, M, and CaCl2, and also back to white if anything changed it
        If StMediaSpe = "S" Then
            LabelPreview.BackColor = Color.FromArgb(255, 255, 0)
        ElseIf StMediaSpe = "M" Then
            LabelPreview.BackColor = Color.FromArgb(197, 224, 179)
        ElseIf StMediaSpe = "CaCl2" Then
            LabelPreview.BackColor = Color.FromArgb(156, 194, 229)
        Else
            LabelPreview.BackColor = Color.FromArgb(255, 255, 255)
        End If

        'Resizes the Label to the size it should be once the font has been shrunk to fit these dimensions
        LabelPreview.AutoSize = False
        LabelPreview.Width = 168 '1.75 inches at 96 DPI
        LabelPreview.Height = 48 '0.5 inches at 96 DPI

        CenterTextInLabel(LabelPreview)

        'This return is only used by ButtonSubmit_Click(). Global variables are included here because they could be reset before ButtonSubmitTask needs them in Module1
        Return (LabelPreview.Text, DecLabelFontSize, StMediaSpe, TextBoxColumns.Text)

    End Function



    'This should simulate the vertical cell alignment of cells in Word via a padding that programmatically adjusts to be the right size
    Private Sub CenterTextInLabel(ByVal Label As Label)
        Dim TextSize As SizeF
        Using LabelGraphics As Graphics = Label.CreateGraphics

            TextSize = TextRenderer.MeasureText(Label.Text, Label.Font)

        End Using

        Dim IntVerticalPadding As Integer = CInt(Math.Floor((Label.Height - TextSize.Height) / 2))
        Label.Padding = New Padding(0, IntVerticalPadding, 0, IntVerticalPadding)
    End Sub



    'Returns a date as a string in the format used for labels
    'This could easily be used for a custom date with a default parameter of today's date
    Private Function DateTimeToString(DateTimeNowObj As DateTime) As String
        Dim formattedDate As String = DateTimeNowObj.ToString("MMddyy")
        Return formattedDate
    End Function



    Private Sub CheckBoxGravityRun_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxGravityRun.CheckedChanged
        If CheckBoxGravityRun.Checked Then
            'Disables the liquid-relevant fields (graying them out), while also removing any text. Changing the text triggers the event handlers that updates the global variables as well
            ComboBoxVolume.Enabled = False
            ComboBoxVolume.Text = ""

            ComboBoxMedia.Enabled = False
            ComboBoxMedia.Text = ""

            ComboBoxMediaType.Enabled = False
            ComboBoxMediaType.Text = ""

            TextBoxRepeat.Enabled = False
            TextBoxRepeat.Text = ""

            DateExp = DateToday.AddMonths(6)

        Else
            'Enables the other fields if the gravity button is unchecked
            ComboBoxVolume.Enabled = True
            ComboBoxMedia.Enabled = True
            ComboBoxMediaType.Enabled = True
            TextBoxRepeat.Enabled = True

            DateExp = DateToday
        End If
        LabelMaker()

    End Sub



    'This feels like I'm re-writing some basic functionality that most definitely exists on its own
    Private Sub CheckedListBoxAutoclaveIDs_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBoxAutoclaveIDs.ItemCheck

        If e.NewValue = CheckState.Checked Then
            For i As Integer = 0 To CheckedListBoxAutoclaveIDs.Items.Count - 1
                If i <> e.Index Then
                    'If this unchecks a previous box to check a new one, i.e. checking 200859 after 200539 is checked, the unchecking of 200539 counts as an ItemCheck, and runs this function once more, whereby 200539's New Value is unchecked, setting StAutoclaveID to "" before the first function call continues and sets it to "200859
                    CheckedListBoxAutoclaveIDs.SetItemChecked(i, False)
                End If
            Next
            StAutoclaveID = CheckedListBoxAutoclaveIDs.SelectedItem
        ElseIf e.NewValue = CheckState.Unchecked Then
            StAutoclaveID = ""
        End If
        LabelMaker()

    End Sub



    Private Sub TextBoxRunNumber_TextChanged(sender As Object, e As EventArgs) Handles TextBoxRunNumber.TextChanged
        If TextBoxRunNumber.Text = "" Then
            StRunNumber = TextBoxRunNumber.Text
        Else
            StRunNumber = "-" & TextBoxRunNumber.Text
        End If
        LabelMaker()
    End Sub



    Private Sub ComboBoxVolume_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxVolume.TextChanged
        If ComboBoxVolume.Text = "" Then
            StVolume = ComboBoxVolume.Text
        Else
            StVolume = ComboBoxVolume.Text & " ML"
        End If
        LabelColumnCountDeterminer()
        LabelMaker()
    End Sub



    Private Sub ComboBoxMediaType_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxMediaType.TextChanged
        StMediaSpe = ComboBoxMediaType.Text
        LabelMaker()
    End Sub



    Private Sub ComboBoxMedia_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxMedia.TextChanged
        StMediaType = ComboBoxMedia.Text

        For i As Integer = 0 To ComboBoxMedia.Items.Count - 1
            If ComboBoxMedia.Items(i) = StMediaType Then

                Select Case StMediaType
                'Make sure all these expiration dates are accurate
                    Case Is = "SCDB"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "FTM"
                        DateExp = DateToday.AddMonths(1)
                    Case Is = "DE BRO"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "DFD"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "LET BRO"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "MSA"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "PBS"
                        DateExp = DateToday.AddMonths(1)
                    Case Is = "PBW"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "R2A"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "SAL 0.9%"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "SDA"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "SDI"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "SMA"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "SP80"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "TR-X"
                        DateExp = DateToday.AddMonths(3)
                    Case Is = "TSA"
                        DateExp = DateToday.AddMonths(3)
                End Select

                'This should exit only if the media has been found
                Exit For

            Else
                'When this is made into a string DateExp will be replaced with 'XXXXXX', since the expiration date should never be today's date
                DateExp = DateToday
            End If
        Next
        LabelColumnCountDeterminer()
        LabelMaker()
    End Sub



    Private Sub TextBoxRepeat_TextChanged(sender As Object, e As EventArgs) Handles TextBoxRepeat.TextChanged
        StRepeat = TextBoxRepeat.Text
        LabelMaker()
    End Sub



    Private Async Sub TextBoxColumns_TextChanged(sender As Object, e As EventArgs) Handles TextBoxColumns.TextChanged
        Dim IntTempColumnInd As Integer = IntColumnInd

        Dim IntTempPageInd As Integer = IntPageInd


        If TextBoxColumns.Text = "" Then
            ButtonSubmit.Enabled = False
            BoolUserEditingColumnCount = False
        Else
            Try
                Dim IntTempColumnCount As Integer = CInt(TextBoxColumns.Text)

                For i As Integer = 1 To IntTempColumnCount - 1
                    IntTempColumnInd += 2
                    If IntTempColumnInd > 7 Then
                        IntTempColumnInd = 1
                        IntTempPageInd += 1
                    End If

                Next

                TextBoxPages.Text = IntTempPageInd

                'If the user is manually editing the column count we don't want the column count to be overwritten by the programmatically determined column count from the volume
                If TextBoxColumns.Focused = True Then
                    BoolUserEditingColumnCount = True
                End If
                If ButtonSubmitTask IsNot Nothing AndAlso ButtonSubmitTask.Status = TaskStatus.Running Then
                    Await ButtonSubmitTask
                    ButtonSubmit.Enabled = True
                Else
                    ButtonSubmit.Enabled = True
                End If

            Catch ex As Exception
                ButtonSubmit.Enabled = False
                TextBoxPages.Text = IntPageInd
            End Try
        End If

    End Sub



    Private Sub CheckBoxMultiple_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxMultiple.CheckedChanged
        If CheckBoxPrint.Enabled Then
            CheckBoxPrint.Enabled = False
        Else
            CheckBoxPrint.Enabled = True
        End If
    End Sub
End Class


'Current Features Lacking:
'Custom dates
'Manual entry of each line
'Stuff to make it pretty or fun, such as something happening during the loading time