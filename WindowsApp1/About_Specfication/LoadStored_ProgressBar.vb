Public Class LoadStored_ProgressBar_Form
    Private Sub Done_Button_Click(sender As Object, e As EventArgs) Handles Done_Button.Click
        MsgBox($"{TextBox1.SelectionStart}/{TextBox1.SelectedText}")
        Dim curPos As Integer = TextBox1.SelectionStart
        With TextBox1
            .SelectionStart = .TextLength
            .ScrollToCaret()
        End With
    End Sub
End Class