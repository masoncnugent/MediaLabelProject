<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        LabelChoice = New Label()
        CheckBoxMediaPrep = New CheckBox()
        CheckBoxBioburden = New CheckBox()
        SuspendLayout()
        ' 
        ' LabelChoice
        ' 
        LabelChoice.AutoSize = True
        LabelChoice.Location = New Point(12, 9)
        LabelChoice.Name = "LabelChoice"
        LabelChoice.Size = New Size(140, 15)
        LabelChoice.TabIndex = 0
        LabelChoice.Text = "Choose A Label Template"
        ' 
        ' CheckBoxMediaPrep
        ' 
        CheckBoxMediaPrep.AutoSize = True
        CheckBoxMediaPrep.Location = New Point(12, 27)
        CheckBoxMediaPrep.Name = "CheckBoxMediaPrep"
        CheckBoxMediaPrep.Size = New Size(86, 19)
        CheckBoxMediaPrep.TabIndex = 1
        CheckBoxMediaPrep.Text = "Media Prep"
        CheckBoxMediaPrep.UseVisualStyleBackColor = True
        ' 
        ' CheckBoxBioburden
        ' 
        CheckBoxBioburden.AutoSize = True
        CheckBoxBioburden.Location = New Point(12, 52)
        CheckBoxBioburden.Name = "CheckBoxBioburden"
        CheckBoxBioburden.Size = New Size(81, 19)
        CheckBoxBioburden.TabIndex = 2
        CheckBoxBioburden.Text = "Bioburden"
        CheckBoxBioburden.UseVisualStyleBackColor = True
        ' 
        ' Form2
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(163, 72)
        Controls.Add(CheckBoxBioburden)
        Controls.Add(CheckBoxMediaPrep)
        Controls.Add(LabelChoice)
        Name = "Form2"
        Text = "Form2"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents LabelChoice As Label
    Friend WithEvents CheckBoxMediaPrep As CheckBox
    Friend WithEvents CheckBoxBioburden As CheckBox
End Class
