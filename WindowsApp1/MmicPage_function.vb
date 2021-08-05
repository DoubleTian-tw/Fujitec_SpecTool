Imports Microsoft.VisualBasic
Imports Microsoft.Office.Interop

Public Class MmicPage_function

    Sub liftControlNumber(tb1 As textbox, tb2 As TextBox, inLiftTb As TextBox, mpanel As Panel, auto_tbname As String, mtype As String) '手動填入已輸入好的數量後，新增對應數量在mmic的car no.上
        Dim ContainWidth, StartContainDis2 As Integer()
        Dim Contain As Control()

        'Dim StartContainDis As Integer = Num_TextBox.Location.X - Name_TextBox.Location.X '25 '設定每隔間距
        Dim StartPos_x As Integer = tb1.Location.X '10 '起始第一格left寬度
        Const StartPos_y As Integer = 10 '起始top高度
        Dim ConNum_tb As TextBox
        Dim errorTb As Label
        errorTb = New Label


        Dim mmic_Panel_num As String = (mpanel.Controls.Count - 2) / 2 '目前panel中的control行的數量

        '顯示內容數量及文字
        Contain = {tb1, tb2}
        ContainNum = Contain.Count
        'ContainWidth = {60, 60, 60, 60, 60, 60, 60, 60, 60, 100}
        ContainWidth = {tb1.Width, tb2.Width}
        StartContainDis2 = {tb1.Left, tb2.Left}
        '嘗試得到電梯輸入之總數
        Resize_JMForm(JMForm_size.re_size) '重新變大小
        Try
            LiftNum = inLiftTb.Text
            'mpanel.Controls.Clear()

            Dim type_result As MsgBoxResult = MsgBox("是否一併更改Object Name?", vbYesNo, "提醒")

            mpanel.Controls.Remove(ErrorLabel_temp(0))
            If mmic_Panel_num < LiftNum Then
                '動態生成add
                ReDim conNum_tb_temp(LiftNum - 1, ContainNum - 1)

                For i = Int(mmic_Panel_num) + 1 To Int(LiftNum)
                    For j = 1 To ContainNum
                        ConNum_tb = New TextBox()
                        With ConNum_tb
                            If type_result = MsgBoxResult.Yes Then
                                If j = 1 Then
                                    .Text = "L#" & i
                                Else
                                    .Text = mtype
                                End If
                            End If
                            .Width = ContainWidth(j - 1)
                            .Left = StartContainDis2(j - 1)
                            .Top = StartPos_y + (i - 1) * 30
                            .Visible = True
                            .TextAlign = HorizontalAlignment.Center
                            .Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
                            .Name = ($"{auto_tbname}_{i}_{j}")
                        End With
                        mpanel.Controls.Add(ConNum_tb)
                        conNum_tb_temp(i - 1, j - 1) = ConNum_tb
                        ResultOutput_TextBox.Text += conNum_tb_temp(i - 1, j - 1).Name & " 被新增，陣列長度目前為:" & conNum_tb_temp.Length & vbCrLf
                    Next j
                Next i
            ElseIf Int(mmic_Panel_num) = Int(LiftNum) Then
                'DO NOTHING
            Else
                ResultOutput_TextBox.Text += "刪除前陣列長度目前為:" & conNum_tb_temp.Length & vbCrLf
                '動態生成sub
                For i As Integer = Int(mmic_Panel_num) To Int(LiftNum) + 1 Step -1
                    For j As Integer = Int(ContainNum) To 1 Step -1
                        ResultOutput_TextBox.Text += conNum_tb_temp(i - 1, j - 1).Name & " 被刪除，陣列長度目前為:" & conNum_tb_temp.Length & vbCrLf
                        'MsgBox(conNum_tb_temp(i - 1, j - 1).Name)
                        'If conNum_tb_temp(i - 1, j - 1).Name <> tb1.Name Or conNum_tb_temp(i - 1, j - 1).Name <> tb2.Name Then
                        mpanel.Controls.Remove(conNum_tb_temp(i - 1, j - 1))
                        'End If
                    Next
                Next
                ReDim conNum_tb_temp(LiftNum - 1, ContainNum - 1)
                Dim a, b As Integer
                a = 0
                b = 0
                For Each i In mpanel.Controls
                    If i.name <> tb1.Name And i.name <> tb2.Name Then
                        conNum_tb_temp(a, b) = i
                        b = b + 1
                    End If
                    If b > 1 Then
                        b = 0
                        a = a + 1
                    End If
                Next
                ResultOutput_TextBox.Text += "刪除完畢陣列長度目前為:" & conNum_tb_temp.Length & vbCrLf


            End If
        Catch
            mpanel.Controls.Clear()
            With errorTb
                .Width = mpanel.Width - 10
                .Height = 20
                .Left = (mpanel.Width / 2) - (.Width / 2)
                .Top = (mpanel.Height / 2) - (.Height / 2)
                .Text = "輸入數量單位錯誤，請輸入正整數"
                .Name = "ErrorLabel"
                .Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
                .TextAlign = ContentAlignment.MiddleCenter
                .ForeColor = Color.Yellow
                .BackColor = Color.Red
            End With
            mpanel.Controls.Add(errorTb)
            'ReDim ErrorLabel_temp(0)
            ErrorLabel_temp(0) = errorTb
        End Try

    End Sub
End Class
