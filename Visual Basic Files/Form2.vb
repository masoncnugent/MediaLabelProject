Public Class Form2
    Private Sub CheckBoxMediaPrep_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxMediaPrep.CheckedChanged

        Dim MediaPrepForm As New Form1()
        MediaPrepForm.Show()
        Me.Close()

    End Sub

    Private Sub CheckBoxBioburden_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxBioburden.CheckedChanged

    End Sub
End Class