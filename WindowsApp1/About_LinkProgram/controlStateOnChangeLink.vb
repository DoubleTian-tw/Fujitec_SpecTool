Public Class controlStateOnChangeLink

    Overloads Function combobox_when_add(count_CB As Integer, ini_name As String, ini_childname As String _
                                         , ini_fatername As String, skey As System.Text.StringBuilder _
                                         , ChaLink_myCB As ComboBox, Form_myCB As ComboBox, nSize As UInt32, sinifilename As String) '當combobx加一時防止一次加過多情形
        '使用chalink
        Dim a As Integer = 0
        Form_myCB.Items.Clear()
        While a <= count_CB  'a從0開始與count_allFolder比較 相等時執行
            ini_name = ini_childname & a + 1 'folderPath_1開始
            GetPrivateProfileString(ini_fatername, ini_name, "", skey, nSize, sinifilename) '取得數值
            If skey.ToString <> "" Then
                ChaLink_myCB.Items.Add(skey.ToString) '如果取得值不為空則加入combox

                '加過後就給true 就不會造成一值add便很多的狀況
                Form_myCB.Items.Add(skey.ToString)
            Else
                Exit While '否則跳出迴圈
            End If
            a += 1
            count_CB = ChaLink_myCB.Items.Count '重新數combox的數量傳回去，反覆進行
        End While
        Return count_CB
    End Function
    Overloads Function combobox_when_add(count_CB As Integer, ini_name As String, ini_childname As String _
                                         , ini_fatername As String, skey As System.Text.StringBuilder _
                                         , Form_myCB As ComboBox, nSize As UInt32, sinifilename As String) '當combobx加一時防止一次加過多情形
        '不使用chalink
        Dim a As Integer = 0
        Form_myCB.Items.Clear()
        While a <= count_CB  'a從0開始與count_allFolder比較 相等時執行
            ini_name = ini_childname & a + 1 'folderPath_1開始
            GetPrivateProfileString(ini_fatername, ini_name, "", skey, nSize, sinifilename) '取得數值
            If skey.ToString <> "" Then
                'ChaLink_myCB.Items.Add(skey.ToString) '如果取得值不為空則加入combox

                '加過後就給true 就不會造成一值add便很多的狀況
                Form_myCB.Items.Add(skey.ToString)
            Else
                Exit While '否則跳出迴圈
            End If
            a += 1
            count_CB = Form_myCB.Items.Count '重新數combox的數量傳回去，反覆進行
        End While
        Return count_CB
    End Function
    Overloads Function combobox_when_add(count_CB As Integer, ini_name As String, ini_childname As String _
                                         , ini_fatername As String, skey As System.Text.StringBuilder _
                                         , myCB As ComboBox, magicTool_myCkList As CheckedListBox, nSize As UInt32, sinifilename As String) '當combobx加一時防止一次加過多情形
        Dim a As Integer = 0
        magicTool_myCkList.Items.Clear()
        While a <= count_CB  'a從0開始與count_allFolder比較 相等時執行
            ini_name = ini_childname & a + 1 'folderPath_1開始
            GetPrivateProfileString(ini_fatername, ini_name, "", skey, nSize, sinifilename) '取得數值
            If skey.ToString <> "" Then
                If myCB.Text <> skey.ToString Then
                    myCB.Items.Add(skey.ToString) '如果取得值不為空則加入combox
                End If
                '加過後就給true 就不會造成一值add便很多的狀況
                'If magicTool_ifadd = False Then
                If magicTool_myCkList.Text <> skey.ToString Then
                    magicTool_myCkList.Items.Add(skey.ToString)
                End If
            Else
                Exit While '否則跳出迴圈
            End If
            a += 1
            count_CB = myCB.Items.Count '重新數combox的數量傳回去，反覆進行
        End While
        Return count_CB
    End Function
    Sub combobox_when_save(count_CB As Integer, ini_name As String, ini_childname As String, ini_fathername As String _
                           , skey As System.Text.StringBuilder, cb As ComboBox _
                           , nSize As UInt32, sinifilename As String)
        Dim a As Integer
        For a = 0 To count_CB
            ini_name = ini_childname & a + 1

            If skey.ToString IsNot "" Then
                skey.Clear()
            End If

            Try   '如果all folder path 有刪除動作時，在跑combox.item時會少1，造成error，使用try catch讓刪除的值達成清空的步驟
                skey = skey.Append(cb.Items(a).ToString())
                WritePrivateProfileString(ini_fathername, ini_name, skey, sinifilename)
            Catch ex As Exception
                skey.Clear()
                WritePrivateProfileString(ini_fathername, ini_name, skey, sinifilename)
            End Try
        Next a

    End Sub







    Sub Add_button(tb As TextBox, cb As ComboBox) '新增鈕用在基本設定內
        If tb.Text <> "" Then
            If tb.Text = cb.Text Then
                Exit Sub
            End If
            cb.Items.Add(tb.Text)
        End If
    End Sub

    Sub Sub_button(tb As TextBox, cb As ComboBox) '刪除鈕用在基本設定內
        If tb.Text = cb.Text Then
            cb.Items.Remove(tb.Text)
        End If
    End Sub
End Class
