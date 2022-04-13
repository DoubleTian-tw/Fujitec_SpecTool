Module replaceControllerName
    'Public Const ctrlTypeName_Panel As String = "Panel"
    'Public Const ctrlTypeName_ComboBox As String = "ComboBox"
    'Public Const ctrlTypeName_Label As String = "Label"
    'Public Const ctrlTypeName_TextBox As String = "TextBox"
    'Public Const ctrlTypeName_CheckBox As String = "CheckBox"


    'Public Function repalce_replaceName_to_myCtrlType_inMyCtrl(myCtrl As Control, myCtrlType As String, replaceName As String) As String
    '    Return Replace(myCtrl.Name, myCtrlType, replaceName)
    'End Function

    '''' <summary>
    '''' 在myPanel中，取代之後的名字取得該控制項的Enabled    
    '''' </summary>
    '''' <param name="replace_name"></param>
    '''' <param name="myPanel"></param>
    '''' <returns></returns>
    'Public Function getRelace_Enable_onPanel(replace_name As String, myPanel As Control) As Boolean
    '        Try
    '            For Each mCtrl As Control In myPanel.Controls
    '                If mCtrl.Name = replace_name Then
    '                    getRelace_Enable_onPanel = mCtrl.Enabled
    '                    Return mCtrl.Enabled
    '                    Exit For
    '                Else
    '                    getRelace_Enable_onPanel = False
    '                End If
    '            Next

    '            'If getRelace_Enable_onPanel = False Then
    '            '    Return False
    '            'End If
    '        Catch ex As Exception
    '            MsgBox($"getRelace_Enable_onPanel function error : {ex.ToString}")
    '        End Try
    '    End Function

    '    ''' <summary>
    '    ''' 在myPanel中，取代之後的名字(replace_name)取得該控制項的文字，例如:A_Panel > A_Label 取得A_Label.Text
    '    ''' </summary>
    '    ''' <param name="replace_name"></param>
    '    ''' <param name="myPanel"></param>
    '    ''' <returns></returns>
    '    Public Function getRelace_NameText_onPanel(replace_name As String, myPanel As Control) As String
    '        Try
    '            For Each mCtrl As Control In myPanel.Controls
    '                If mCtrl.Name = replace_name Then
    '                    Return mCtrl.Text
    '                    Exit For
    '                End If
    '            Next
    '        Catch ex As Exception
    '            MsgBox($"getRelace_NameText_onPanel function error : {ex.ToString}")
    '        End Try
    '    End Function

    '''' <summary>
    '''' 在myPanel中，取代之後的名字取得控制項的Checked狀態
    '''' </summary>
    '''' <param name="replace_name"></param>
    '''' <param name="myPanel"></param>
    '''' <returns></returns>
    'Public Function getRelace_ChkBoxState_onPanel(replace_name As String, myPanel As Control) As Boolean
    '    Try
    '        For Each mCtrl As Control In myPanel.Controls
    '            If TypeOf (mCtrl) Is CheckBox And mCtrl.Name = replace_name Then
    '                Dim mChkBox As CheckBox
    '                mChkBox = mCtrl
    '                Return mChkBox.Checked
    '                Exit For
    '            Else
    '                getRelace_ChkBoxState_onPanel = False
    '            End If
    '        Next
    '        'If getRelace_ChkBoxState_onPanel = False Then
    '        '    Return False
    '        'End If
    '    Catch ex As Exception
    '        MsgBox($"getRelace_ChkBoxState_onPanel function error : {ex.ToString}")
    '    End Try
    'End Function

End Module
