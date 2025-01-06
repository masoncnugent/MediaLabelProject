<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        ButtonSubmit = New Button()
        TextBoxRunNumber = New TextBox()
        ComboBoxMediaType = New ComboBox()
        LabelRunNumber = New Label()
        LabelVolume = New Label()
        LabelMediaType = New Label()
        CheckBoxGravityRun = New CheckBox()
        ComboBoxVolume = New ComboBox()
        ColorDialog1 = New ColorDialog()
        CheckedListBoxAutoclaveIDs = New CheckedListBox()
        LabelAutoclaveID = New Label()
        LabelMedia = New Label()
        ComboBoxMedia = New ComboBox()
        LabelPreview = New Label()
        TextBoxRepeat = New TextBox()
        LabelRepeat = New Label()
        CheckBoxMultiple = New CheckBox()
        LabelColumns = New Label()
        ProgressBarLabels = New ProgressBar()
        TextBoxColumns = New TextBox()
        TextBoxPages = New TextBox()
        LabelPages = New Label()
        CheckBoxPrint = New CheckBox()
        SuspendLayout()
        ' 
        ' ButtonSubmit
        ' 
        ButtonSubmit.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        ButtonSubmit.Location = New Point(496, 96)
        ButtonSubmit.Name = "ButtonSubmit"
        ButtonSubmit.Size = New Size(64, 48)
        ButtonSubmit.TabIndex = 7
        ButtonSubmit.Text = "Submit"
        ButtonSubmit.UseVisualStyleBackColor = True
        ' 
        ' TextBoxRunNumber
        ' 
        TextBoxRunNumber.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        TextBoxRunNumber.Location = New Point(144, 40)
        TextBoxRunNumber.Name = "TextBoxRunNumber"
        TextBoxRunNumber.Size = New Size(64, 21)
        TextBoxRunNumber.TabIndex = 9
        ' 
        ' ComboBoxMediaType
        ' 
        ComboBoxMediaType.AutoCompleteCustomSource.AddRange(New String() {"S", "M", "CaCl2"})
        ComboBoxMediaType.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        ComboBoxMediaType.FormattingEnabled = True
        ComboBoxMediaType.Items.AddRange(New Object() {"S", "M", "CaCl2"})
        ComboBoxMediaType.Location = New Point(408, 40)
        ComboBoxMediaType.Name = "ComboBoxMediaType"
        ComboBoxMediaType.Size = New Size(64, 23)
        ComboBoxMediaType.TabIndex = 11
        ' 
        ' LabelRunNumber
        ' 
        LabelRunNumber.AutoSize = True
        LabelRunNumber.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        LabelRunNumber.Location = New Point(144, 24)
        LabelRunNumber.Name = "LabelRunNumber"
        LabelRunNumber.Size = New Size(48, 15)
        LabelRunNumber.TabIndex = 12
        LabelRunNumber.Text = "Run No."
        ' 
        ' LabelVolume
        ' 
        LabelVolume.AutoSize = True
        LabelVolume.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        LabelVolume.Location = New Point(232, 24)
        LabelVolume.Name = "LabelVolume"
        LabelVolume.Size = New Size(42, 15)
        LabelVolume.TabIndex = 13
        LabelVolume.Text = "Volume"
        ' 
        ' LabelMediaType
        ' 
        LabelMediaType.AutoSize = True
        LabelMediaType.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        LabelMediaType.Location = New Point(408, 24)
        LabelMediaType.Name = "LabelMediaType"
        LabelMediaType.Size = New Size(67, 15)
        LabelMediaType.TabIndex = 14
        LabelMediaType.Text = "Media Type"
        ' 
        ' CheckBoxGravityRun
        ' 
        CheckBoxGravityRun.AutoSize = True
        CheckBoxGravityRun.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        CheckBoxGravityRun.Location = New Point(24, 152)
        CheckBoxGravityRun.Name = "CheckBoxGravityRun"
        CheckBoxGravityRun.Size = New Size(87, 19)
        CheckBoxGravityRun.TabIndex = 15
        CheckBoxGravityRun.Text = "Gravity Run"
        CheckBoxGravityRun.UseVisualStyleBackColor = True
        ' 
        ' ComboBoxVolume
        ' 
        ComboBoxVolume.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        ComboBoxVolume.FormattingEnabled = True
        ComboBoxVolume.Items.AddRange(New Object() {"3000", "2000", "1000", "600", "400", "100"})
        ComboBoxVolume.Location = New Point(232, 40)
        ComboBoxVolume.Name = "ComboBoxVolume"
        ComboBoxVolume.Size = New Size(64, 23)
        ComboBoxVolume.TabIndex = 16
        ' 
        ' CheckedListBoxAutoclaveIDs
        ' 
        CheckedListBoxAutoclaveIDs.CheckOnClick = True
        CheckedListBoxAutoclaveIDs.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        CheckedListBoxAutoclaveIDs.FormattingEnabled = True
        CheckedListBoxAutoclaveIDs.Items.AddRange(New Object() {"200539", "200859", "201738", "200359", "200475"})
        CheckedListBoxAutoclaveIDs.Location = New Point(24, 40)
        CheckedListBoxAutoclaveIDs.Name = "CheckedListBoxAutoclaveIDs"
        CheckedListBoxAutoclaveIDs.Size = New Size(96, 84)
        CheckedListBoxAutoclaveIDs.TabIndex = 20
        ' 
        ' LabelAutoclaveID
        ' 
        LabelAutoclaveID.AutoSize = True
        LabelAutoclaveID.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        LabelAutoclaveID.Location = New Point(24, 24)
        LabelAutoclaveID.Name = "LabelAutoclaveID"
        LabelAutoclaveID.Size = New Size(72, 15)
        LabelAutoclaveID.TabIndex = 21
        LabelAutoclaveID.Text = "Autoclave ID"
        ' 
        ' LabelMedia
        ' 
        LabelMedia.AutoSize = True
        LabelMedia.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        LabelMedia.Location = New Point(320, 24)
        LabelMedia.Name = "LabelMedia"
        LabelMedia.Size = New Size(38, 15)
        LabelMedia.TabIndex = 22
        LabelMedia.Text = "Media"
        ' 
        ' ComboBoxMedia
        ' 
        ComboBoxMedia.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        ComboBoxMedia.FormattingEnabled = True
        ComboBoxMedia.Items.AddRange(New Object() {"SCDB", "FTM", "DE BRO", "DFD", "LET BRO", "MSA", "PBS", "PBW", "R2A", "SAL 0.9%", "SDA", "SDI", "SMA", "SP80", "TR-X", "TSA"})
        ComboBoxMedia.Location = New Point(320, 40)
        ComboBoxMedia.Name = "ComboBoxMedia"
        ComboBoxMedia.Size = New Size(64, 23)
        ComboBoxMedia.TabIndex = 23
        ' 
        ' LabelPreview
        ' 
        LabelPreview.BackColor = SystemColors.Window
        LabelPreview.Font = New Font("Times New Roman", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        LabelPreview.Location = New Point(304, 120)
        LabelPreview.Name = "LabelPreview"
        LabelPreview.Size = New Size(168, 48)
        LabelPreview.TabIndex = 24
        ' 
        ' TextBoxRepeat
        ' 
        TextBoxRepeat.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        TextBoxRepeat.Location = New Point(496, 40)
        TextBoxRepeat.Name = "TextBoxRepeat"
        TextBoxRepeat.Size = New Size(64, 21)
        TextBoxRepeat.TabIndex = 25
        ' 
        ' LabelRepeat
        ' 
        LabelRepeat.AutoSize = True
        LabelRepeat.Font = New Font("Times New Roman", 9F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        LabelRepeat.Location = New Point(496, 24)
        LabelRepeat.Name = "LabelRepeat"
        LabelRepeat.Size = New Size(41, 15)
        LabelRepeat.TabIndex = 26
        LabelRepeat.Text = "Repeat"
        LabelRepeat.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' CheckBoxMultiple
        ' 
        CheckBoxMultiple.AutoSize = True
        CheckBoxMultiple.BackColor = SystemColors.Control
        CheckBoxMultiple.Location = New Point(320, 88)
        CheckBoxMultiple.Name = "CheckBoxMultiple"
        CheckBoxMultiple.Size = New Size(75, 19)
        CheckBoxMultiple.TabIndex = 27
        CheckBoxMultiple.Text = "Multiple?"
        CheckBoxMultiple.UseVisualStyleBackColor = False
        ' 
        ' LabelColumns
        ' 
        LabelColumns.AutoSize = True
        LabelColumns.Location = New Point(144, 72)
        LabelColumns.Name = "LabelColumns"
        LabelColumns.Size = New Size(55, 15)
        LabelColumns.TabIndex = 28
        LabelColumns.Text = "Columns"
        ' 
        ' ProgressBarLabels
        ' 
        ProgressBarLabels.Location = New Point(496, 144)
        ProgressBarLabels.Name = "ProgressBarLabels"
        ProgressBarLabels.Size = New Size(64, 23)
        ProgressBarLabels.Step = 1
        ProgressBarLabels.TabIndex = 31
        ' 
        ' TextBoxColumns
        ' 
        TextBoxColumns.Location = New Point(144, 88)
        TextBoxColumns.Name = "TextBoxColumns"
        TextBoxColumns.Size = New Size(64, 23)
        TextBoxColumns.TabIndex = 32
        ' 
        ' TextBoxPages
        ' 
        TextBoxPages.Enabled = False
        TextBoxPages.Location = New Point(232, 88)
        TextBoxPages.Name = "TextBoxPages"
        TextBoxPages.Size = New Size(64, 23)
        TextBoxPages.TabIndex = 33
        ' 
        ' LabelPages
        ' 
        LabelPages.AutoSize = True
        LabelPages.Location = New Point(232, 72)
        LabelPages.Name = "LabelPages"
        LabelPages.Size = New Size(38, 15)
        LabelPages.TabIndex = 34
        LabelPages.Text = "Pages"
        ' 
        ' CheckBoxPrint
        ' 
        CheckBoxPrint.AutoSize = True
        CheckBoxPrint.Location = New Point(408, 88)
        CheckBoxPrint.Name = "CheckBoxPrint"
        CheckBoxPrint.Size = New Size(56, 19)
        CheckBoxPrint.TabIndex = 35
        CheckBoxPrint.Text = "Print?"
        CheckBoxPrint.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(577, 193)
        Controls.Add(CheckBoxPrint)
        Controls.Add(LabelPages)
        Controls.Add(TextBoxPages)
        Controls.Add(TextBoxColumns)
        Controls.Add(ProgressBarLabels)
        Controls.Add(LabelColumns)
        Controls.Add(CheckBoxMultiple)
        Controls.Add(LabelRepeat)
        Controls.Add(TextBoxRepeat)
        Controls.Add(LabelPreview)
        Controls.Add(ComboBoxMedia)
        Controls.Add(LabelMedia)
        Controls.Add(LabelAutoclaveID)
        Controls.Add(CheckedListBoxAutoclaveIDs)
        Controls.Add(ComboBoxVolume)
        Controls.Add(CheckBoxGravityRun)
        Controls.Add(LabelMediaType)
        Controls.Add(LabelVolume)
        Controls.Add(LabelRunNumber)
        Controls.Add(ComboBoxMediaType)
        Controls.Add(TextBoxRunNumber)
        Controls.Add(ButtonSubmit)
        Name = "Form1"
        Text = "Media Prep Label Maker"
        ResumeLayout(False)
        PerformLayout()
    End Sub
    Friend WithEvents ButtonSubmit As Button
    Friend WithEvents TextBoxRunNumber As TextBox
    Friend WithEvents ComboBoxMediaType As ComboBox
    Friend WithEvents LabelRunNumber As Label
    Friend WithEvents LabelVolume As Label
    Friend WithEvents LabelMediaType As Label
    Friend WithEvents CheckBoxGravityRun As CheckBox
    Friend WithEvents ComboBoxVolume As ComboBox
    Friend WithEvents ColorDialog1 As ColorDialog
    Friend WithEvents CheckedListBoxAutoclaveIDs As CheckedListBox
    Friend WithEvents LabelAutoclaveID As Label
    Friend WithEvents LabelMedia As Label
    Friend WithEvents ComboBoxMedia As ComboBox
    Friend WithEvents LabelPreview As Label
    Friend WithEvents TextBoxRepeat As TextBox
    Friend WithEvents LabelRepeat As Label
    Friend WithEvents CheckBoxMultiple As CheckBox
    Friend WithEvents LabelColumns As Label
    Friend WithEvents ProgressBarLabels As ProgressBar
    Friend WithEvents TextBoxColumns As TextBox
    Friend WithEvents TextBoxPages As TextBox
    Friend WithEvents LabelPages As Label
    Friend WithEvents CheckBoxPrint As CheckBox

End Class
