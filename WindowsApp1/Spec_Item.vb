Public Class Spec_Item
    'Public specTW_cmbbox() As ComboBox
    Public specTW_panel() As Panel
    'Public specTW_label() As Label

    Public Sub ini_specTW_AllControler()
        'specTW_cmbbox = {JobMaker_Form.Spec_DRAuto_ComboBox, JobMaker_Form.Spec_CancellCall_ComboBox,
        '                 JobMaker_Form.Spec_AutoFan_ComboBox, JobMaker_Form.Spec_AutoPass_ComboBox,
        '                 JobMaker_Form.Spec_Indep_ComboBox,
        '                 JobMaker_Form.Spec_HinCpi_ComboBox, JobMaker_Form.Spec_Fire_ComboBox,
        '                 JobMaker_Form.Spec_Fireman_ComboBox, JobMaker_Form.Spec_Parking_ComboBox,
        '                 JobMaker_Form.Spec_Seismic_ComboBox, JobMaker_Form.Spec_CPI_ComboBox,
        '                 JobMaker_Form.Spec_HallGong_ComboBox, JobMaker_Form.Spec_HPIMsg_ComboBox,
        '                 JobMaker_Form.Spec_CarGong_ComboBox, JobMaker_Form.Spec_CRD_ComboBox,
        '                 JobMaker_Form.Spec_VonicBz_ComboBox, JobMaker_Form.Spec_DrHold_ComboBox,
        '                 JobMaker_Form.Spec_Landic_ComboBox, JobMaker_Form.Spec_MFLReturn_ComboBox,
        '                 JobMaker_Form.Spec_Vonic_ComboBox, JobMaker_Form.Spec_Emer_ComboBox,
        '                 JobMaker_Form.Spec_Elvic_ComboBox, JobMaker_Form.Spec_WCOB_ComboBox,
        '                 JobMaker_Form.Spec_HLL_ComboBox, JobMaker_Form.Spec_ATT_ComboBox,
        '                 JobMaker_Form.Spec_Flood_ComboBox, JobMaker_Form.Spec_LS1M_ComboBox,
        '                 JobMaker_Form.Spec_PRU_ComboBox, JobMaker_Form.Spec_LoadCell_ComboBox,
        '                 JobMaker_Form.Spec_FrontRearDr_ComboBox, JobMaker_Form.Spec_OpeSw_ComboBox}

        specTW_panel = {JobMaker_Form.Spec_DRAuto_Panel, JobMaker_Form.Spec_CancellCall_Panel,
                       JobMaker_Form.Spec_AutoFan_Panel, JobMaker_Form.Spec_AutoPass_Panel,
                       JobMaker_Form.Spec_Indep_Panel,
                       JobMaker_Form.Spec_HinCpi_Panel, JobMaker_Form.Spec_Fire_Panel,
                       JobMaker_Form.Spec_Fireman_Panel, JobMaker_Form.Spec_Parking_Panel,
                       JobMaker_Form.Spec_Seismic_Panel, JobMaker_Form.Spec_CPI_Panel,
                       JobMaker_Form.Spec_HallGong_Panel, JobMaker_Form.Spec_HPIMsg_Panel,
                       JobMaker_Form.Spec_CarGong_Panel, JobMaker_Form.Spec_CRD_Panel,
                       JobMaker_Form.Spec_VonicBz_Panel, JobMaker_Form.Spec_DrHold_Panel,
                       JobMaker_Form.Spec_Landic_Panel, JobMaker_Form.Spec_MFLReturn_Panel,
                       JobMaker_Form.Spec_Vonic_Panel, JobMaker_Form.Spec_Emer_Panel,
                       JobMaker_Form.Spec_Elvic_Panel, JobMaker_Form.Spec_WCOB_Panel,
                       JobMaker_Form.Spec_HLL_Panel, JobMaker_Form.Spec_ATT_Panel,
                       JobMaker_Form.Spec_Flood_Panel, JobMaker_Form.Spec_LS1M_Panel,
                       JobMaker_Form.Spec_PRU_Panel, JobMaker_Form.Spec_LoadCell_Panel,
                       JobMaker_Form.Spec_FrontRearDr_Panel, JobMaker_Form.Spec_OpeSw_Panel}

        'specTW_label = {JobMaker_Form.Spec_DRAuto_Label, JobMaker_Form.Spec_CancellCall_Label,
        '                JobMaker_Form.Spec_AutoFan_Label, JobMaker_Form.Spec_AutoPass_Label,
        '                JobMaker_Form.Spec_Indep_Label,
        '                JobMaker_Form.Spec_HinCpi_Label, JobMaker_Form.Spec_Fire_Label,
        '                JobMaker_Form.Spec_Fireman_Label, JobMaker_Form.Spec_Parking_Label,
        '                JobMaker_Form.Spec_Seismic_Label, JobMaker_Form.Spec_CPI_Label,
        '                JobMaker_Form.Spec_HallGong_Label, JobMaker_Form.Spec_HPIMsg_Label,
        '                JobMaker_Form.Spec_CarGong_Label, JobMaker_Form.Spec_CRD_Label,
        '                JobMaker_Form.Spec_VonicBz_Label, JobMaker_Form.Spec_DrHold_Label,
        '                JobMaker_Form.Spec_Landic_Label, JobMaker_Form.Spec_MFLReturn_Label,
        '                JobMaker_Form.Spec_Vonic_Label, JobMaker_Form.Spec_Emer_Label,
        '                JobMaker_Form.Spec_Elvic_Label, JobMaker_Form.Spec_WCOB_Label,
        '                JobMaker_Form.Spec_HLL_Label, JobMaker_Form.Spec_ATT_Label,
        '                JobMaker_Form.Spec_Flood_Label, JobMaker_Form.Spec_LS1M_Label,
        '                JobMaker_Form.Spec_PRU_Label, JobMaker_Form.Spec_LoadCell_Label,
        '                JobMaker_Form.Spec_FrontRearDr_Label, JobMaker_Form.Spec_OpeSw_Label}
    End Sub

    Public Function repalce_replaceName_to_myCtrlType_inMyCtrl(myCtrl As Control, myCtrlType As String, replaceName As String) As String
        Return Replace(myCtrl.Name, myCtrlType, replaceName)
    End Function


    ''' <summary>
    ''' 在myPanel中，取代之後的名字取得該控制項的Enabled    
    ''' </summary>
    ''' <param name="replace_name"></param>
    ''' <param name="myPanel"></param>
    ''' <returns></returns>
    Public Function getRelace_Enable_onPanel(replace_name As String, myPanel As Control) As Boolean
        Try
            For Each mCtrl As Control In myPanel.Controls
                If mCtrl.Name = replace_name Then
                    getRelace_Enable_onPanel = mCtrl.Enabled
                    Return mCtrl.Enabled
                    Exit For
                Else
                    getRelace_Enable_onPanel = False
                End If
            Next

            'If getRelace_Enable_onPanel = False Then
            '    Return False
            'End If
        Catch ex As Exception
            MsgBox($"getRelace_Enable_onPanel function error : {ex.ToString}")
        End Try
    End Function

    ''' <summary>
    ''' 在myPanel中，取代之後的名字(replace_name)取得該控制項的文字，例如:A_Panel > A_Label 取得A_Label.Text
    ''' </summary>
    ''' <param name="replace_name"></param>
    ''' <param name="myPanel"></param>
    ''' <returns></returns>
    Public Function getRelace_NameText_onPanel(replace_name As String, myPanel As Control) As String
        Try
            For Each mCtrl As Control In myPanel.Controls
                If mCtrl.Name = replace_name Then
                    Return mCtrl.Text
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox($"getRelace_NameText_onPanel function error : {ex.ToString}")
        End Try
    End Function

    ''' <summary>
    ''' 在myPanel中，取代之後的名字取得控制項的Checked狀態
    ''' </summary>
    ''' <param name="replace_name"></param>
    ''' <param name="myPanel"></param>
    ''' <returns></returns>
    Public Function getRelace_ChkBoxState_onPanel(replace_name As String, myPanel As Control) As Boolean
        Try
            For Each mCtrl As Control In myPanel.Controls
                If TypeOf (mCtrl) Is CheckBox And mCtrl.Name = replace_name Then
                    Dim mChkBox As CheckBox
                    mChkBox = mCtrl
                    Return mChkBox.Checked
                    Exit For
                Else
                    getRelace_ChkBoxState_onPanel = False
                End If
            Next
            'If getRelace_ChkBoxState_onPanel = False Then
            '    Return False
            'End If
        Catch ex As Exception
            MsgBox($"getRelace_ChkBoxState_onPanel function error : {ex.ToString}")
        End Try
    End Function
End Class
