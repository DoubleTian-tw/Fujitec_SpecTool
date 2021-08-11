Public Class About
    Dim chalink As ChangeLink = New ChangeLink()
    Dim ReadAbout_dat As String
    Dim About_path As String = Application.StartupPath & "\dat\About.dat"

    Private Sub About_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Try
        '    chalink.Initialization_ini()
        '    chalink.formPositionOnScreen_Setting(Me, chalink.sKeyValueScr.ToString, chalink.sKeyValuePos.ToString)

        '    chalink.Topmost_setting(Me, False)
        '    ReadAbout_dat = IO.File.ReadAllText(About_path)
        '    If ReadAbout_dat Is Nothing Then
        '        MsgBox("About.dat檔案遺失，請重新匯入")
        '    End If
        '    About_TextBox.Text = ReadAbout_dat
        'Catch ex As Exception
        '    MsgBox("開啟失敗")
        'End Try
    End Sub

    Private Sub About_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        'MagicTool.Show()
    End Sub

    Private Sub ConfirmSave_Button_Click(sender As Object, e As EventArgs) Handles ConfirmSave_Button.Click
        'Dim myvalue As Object = InputBox("輸入密碼 : ", "About Form Update")
        'If myvalue = "1111" Then
        '    MsgBox("正確")
        '    IO.File.WriteAllText(About_path, About_TextBox.Text)
        'Else
        '    MsgBox("密碼錯誤")
        'End If
    End Sub
End Class