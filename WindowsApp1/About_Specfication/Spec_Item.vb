Module Spec_Item
    Public specTW_panel() As Panel
    Public Const ctrlTypeName_Panel As String = "Panel"
    Public Const ctrlTypeName_ComboBox As String = "ComboBox"
    Public Const ctrlTypeName_Label As String = "Label"
    Public Const ctrlTypeName_TextBox As String = "TextBox"
    Public Const ctrlTypeName_CheckBox As String = "CheckBox"

    Public Sub ini_specTW_AllControler()
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
                       JobMaker_Form.Spec_Elvic_Panel,
                       JobMaker_Form.Spec_WTB_Panel, JobMaker_Form.Spec_WCOB_Panel,
                       JobMaker_Form.Spec_HLL_Panel, JobMaker_Form.Spec_ATT_Panel,
                       JobMaker_Form.Spec_Flood_Panel, JobMaker_Form.Spec_LS1M_Panel,
                       JobMaker_Form.Spec_PRU_Panel, JobMaker_Form.Spec_LoadCell_Panel,
                       JobMaker_Form.Spec_FrontRearDr_Panel, JobMaker_Form.Spec_OpeSw_Panel}
    End Sub

    ''' <summary>
    ''' 將myCtrl名稱的myCtrlType文字，取代為replaceName
    ''' </summary>
    ''' <param name="myCtrl">目標取代控制項</param>
    ''' <param name="myCtrlType">要取代的文字</param>
    ''' <param name="replaceName">取代後的文字</param>
    ''' <returns></returns>
    Public Function replace_replaceName_to_myCtrlType_inMyCtrl(myCtrl As Control, myCtrlType As String, replaceName As String) As String
        Return Replace(myCtrl.Name, myCtrlType, replaceName)
    End Function


    ''' <summary>
    ''' 在myPanel中，取代之後的名字取得該控制項的Enabled    
    ''' </summary>
    ''' <param name="replace_name"></param>
    ''' <param name="myPanel"></param>
    ''' <returns></returns>
    Public Function getRelace_Enable_onPanel(replace_name As String, myPanel As Control) As Boolean
        getRelace_Enable_onPanel = False
        Try
            For Each mCtrl As Control In myPanel.Controls
                If mCtrl.Name = replace_name Then
                    getRelace_Enable_onPanel = mCtrl.Enabled
                    Return getRelace_Enable_onPanel
                    Exit For
                Else
                    getRelace_Enable_onPanel = False
                End If
            Next
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
        getRelace_NameText_onPanel = ""
        Try
            For Each mCtrl As Control In myPanel.Controls
                If mCtrl.Name = replace_name Then
                    getRelace_NameText_onPanel = mCtrl.Text
                    Return getRelace_NameText_onPanel
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
        getRelace_ChkBoxState_onPanel = False
        Try
            For Each mCtrl As Control In myPanel.Controls
                If TypeOf (mCtrl) Is CheckBox And mCtrl.Name = replace_name Then
                    Dim mChkBox As CheckBox
                    mChkBox = mCtrl
                    getRelace_ChkBoxState_onPanel = mChkBox.Checked
                    Return getRelace_ChkBoxState_onPanel
                    Exit For
                Else
                    getRelace_ChkBoxState_onPanel = False
                End If
            Next

        Catch ex As Exception
            MsgBox($"getRelace_ChkBoxState_onPanel function error : {ex.ToString}")
        End Try
    End Function
End Module
