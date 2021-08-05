Imports Microsoft.VisualBasic
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.SystemInformation
Public Class SetPosition

    Public Sub ScreenPosChange(Scr As String, Pos As String)

        If Scr = "主螢幕" Then

            If Pos = "左上" Then
                settingFormPos(0, 0)
            ElseIf Pos = "右上" Then
                settingFormPos(WorkingArea.Width - Size.Width, 0)
            ElseIf Pos = "左下" Then
                settingFormPos(0, WorkingArea.Height - Size.Height)
            ElseIf Pos = "右下" Then
                settingFormPos(WorkingArea.Width - Size.Width, WorkingArea.Height - Size.Height)
            End If
        ElseIf Scr = "副螢幕" Then

            If Pos = "左上" Then
                settingFormPos(-1920, 0)
            ElseIf Pos = "右上" Then
                settingFormPos(-Size.Width, 0)
            ElseIf Pos = "左下" Then
                settingFormPos(-1920, WorkingArea.Height - Size.Height)
            ElseIf Pos = "右下" Then
                settingFormPos(-Size.Width, WorkingArea.Height - Size.Height)
            End If
        End If

    End Sub

    Private Function settingFormPos(Form As String, Form_Position_X As Integer, Form_Position_Y As Integer)
        Form.Location = New System.Drawing.Point(Form_Position_X, Form_Position_Y)
        MagicTool.Location = New System.Drawing.Point(Form_Position_X, Form_Position_Y)
    End Function

End Class
