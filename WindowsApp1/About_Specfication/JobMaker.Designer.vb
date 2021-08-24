<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class JobMaker_Form
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(JobMaker_Form))
        Me.ResultCheck_Button = New System.Windows.Forms.Button()
        Me.ResultOutput_TextBox = New System.Windows.Forms.TextBox()
        Me.JobMaker_Timer = New System.Windows.Forms.Timer(Me.components)
        Me.G_TabPage = New System.Windows.Forms.TabPage()
        Me.GWeb_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label86 = New System.Windows.Forms.Label()
        Me.GWeb_Button = New System.Windows.Forms.Button()
        Me.Use_G_CheckBox = New System.Windows.Forms.CheckBox()
        Me.MMIC_TabPage = New System.Windows.Forms.TabPage()
        Me.MMIC_Panel = New System.Windows.Forms.Panel()
        Me.Panel17 = New System.Windows.Forms.Panel()
        Me.mmicType1_ObjNameBase_TextBox = New System.Windows.Forms.TextBox()
        Me.mmicType1_ObjName_TextBox = New System.Windows.Forms.TextBox()
        Me.mmicType1_CarNo_TextBox = New System.Windows.Forms.TextBox()
        Me.Panel15 = New System.Windows.Forms.Panel()
        Me.mmic_ObjName_TextBox = New System.Windows.Forms.TextBox()
        Me.mmic_CarNo_TextBox = New System.Windows.Forms.TextBox()
        Me.MMIC_VD10_GroupBox = New System.Windows.Forms.GroupBox()
        Me.MMIC_VD10_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.MMIC_VD10_Base_TextBox = New System.Windows.Forms.TextBox()
        Me.MMIC_VD10_Type_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label132 = New System.Windows.Forms.Label()
        Me.Label131 = New System.Windows.Forms.Label()
        Me.Label114 = New System.Windows.Forms.Label()
        Me.Label115 = New System.Windows.Forms.Label()
        Me.Label113 = New System.Windows.Forms.Label()
        Me.MMIC_VD10_Panel = New System.Windows.Forms.Panel()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.MMIC_VD10_ROM_ComboBox = New System.Windows.Forms.ComboBox()
        Me.MMIC_VD10_Quantity_ComboBox = New System.Windows.Forms.ComboBox()
        Me.MMIC_SV_E_GroupBox = New System.Windows.Forms.GroupBox()
        Me.MMIC_SV_E_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.MMIC_SV_ECarObj_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label106 = New System.Windows.Forms.Label()
        Me.MMIC_SV_E_Panel = New System.Windows.Forms.Panel()
        Me.Label107 = New System.Windows.Forms.Label()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.MMIC_SV_EBase_ComboBox = New System.Windows.Forms.ComboBox()
        Me.MMIC_SV_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label231 = New System.Windows.Forms.Label()
        Me.MMIC_SV_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Label130 = New System.Windows.Forms.Label()
        Me.MMIC_SV_Type_ComboBox = New System.Windows.Forms.ComboBox()
        Me.MMIC_SV_Base_TextBox = New System.Windows.Forms.TextBox()
        Me.Label129 = New System.Windows.Forms.Label()
        Me.Label103 = New System.Windows.Forms.Label()
        Me.MMIC_SV_Panel = New System.Windows.Forms.Panel()
        Me.Label104 = New System.Windows.Forms.Label()
        Me.MMIC_MR_E_GroupBox = New System.Windows.Forms.GroupBox()
        Me.MMIC_MR_E_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.MMIC_MR_ECarObj_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label100 = New System.Windows.Forms.Label()
        Me.MMIC_MR_E_Panel = New System.Windows.Forms.Panel()
        Me.Label101 = New System.Windows.Forms.Label()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.MMIC_MR_EBase_ComboBox = New System.Windows.Forms.ComboBox()
        Me.MMIC_MR_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label229 = New System.Windows.Forms.Label()
        Me.MMIC_MR_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.MMIC_MR_Base_TextBox = New System.Windows.Forms.TextBox()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.Label99 = New System.Windows.Forms.Label()
        Me.Label95 = New System.Windows.Forms.Label()
        Me.MMIC_MR_CP43x_ComboBox = New System.Windows.Forms.ComboBox()
        Me.MMIC_MR_Panel = New System.Windows.Forms.Panel()
        Me.Label128 = New System.Windows.Forms.Label()
        Me.MMIC_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label111 = New System.Windows.Forms.Label()
        Me.Label112 = New System.Windows.Forms.Label()
        Me.MMIC_FLEX_N_ComboBox = New System.Windows.Forms.ComboBox()
        Me.MMIC_MachineType_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Use_mmic_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Important_TabPage = New System.Windows.Forms.TabPage()
        Me.ImpSetting_GroupBox = New System.Windows.Forms.GroupBox()
        Me.HIN_TestButton = New System.Windows.Forms.Button()
        Me.HallIndicator_FlowLayoutPanel = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label93 = New System.Windows.Forms.Label()
        Me.Imp_MachineRoom_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Imp_DoorType_TextBox = New System.Windows.Forms.TextBox()
        Me.Label61 = New System.Windows.Forms.Label()
        Me.Label127 = New System.Windows.Forms.Label()
        Me.Label94 = New System.Windows.Forms.Label()
        Me.Label97 = New System.Windows.Forms.Label()
        Me.Label96 = New System.Windows.Forms.Label()
        Me.Imp_OverBalance_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Imp_WHB_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Imp_FAN_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Use_Imp_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec = New System.Windows.Forms.TabPage()
        Me.Spec_TabControl = New System.Windows.Forms.TabControl()
        Me.Spec_BasicAll_TabPage = New System.Windows.Forms.TabPage()
        Me.Spec_BasicAll_TabControl = New System.Windows.Forms.TabControl()
        Me.TabPage7 = New System.Windows.Forms.TabPage()
        Me.SpecBasic_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Spec_LiftNum_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.SpecBasic_LiftItem_Dynamic_Panel = New System.Windows.Forms.Panel()
        Me.SpecBasic_LiftItem_Panel = New System.Windows.Forms.Panel()
        Me.Spec_BtmFL_Real_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_TopFL_Real_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_Control_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_FLName_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_Speed_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_StopFL_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_LiftName_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_BtmFL_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_LiftMem_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_TopFL_TextBox = New System.Windows.Forms.TextBox()
        Me.TabPage8 = New System.Windows.Forms.TabPage()
        Me.SpecBasic_GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Spec_MachineType_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Spec_ControlWay_Panel = New System.Windows.Forms.Panel()
        Me.Spec_MachineType_Label = New System.Windows.Forms.Label()
        Me.SpecBasic_p2_base_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Base_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label189 = New System.Windows.Forms.Label()
        Me.Spec_Purpose_Panel = New System.Windows.Forms.Panel()
        Me.Spec_ControlWay_Label = New System.Windows.Forms.Label()
        Me.Spec_MachineType_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Purpose_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Use_SpecBasic_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_TW_TabPage = New System.Windows.Forms.TabPage()
        Me.Spec_TW_TabControl = New System.Windows.Forms.TabControl()
        Me.TabPage9 = New System.Windows.Forms.TabPage()
        Me.Spec_TW_FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Spec_DRAuto_Panel = New System.Windows.Forms.Panel()
        Me.Spec_DRAuto_Label = New System.Windows.Forms.Label()
        Me.Spec_MechSafety_Label = New System.Windows.Forms.Label()
        Me.Spec_MechSafety_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_PhotoEye_Label = New System.Windows.Forms.Label()
        Me.Spec_PhotoEye_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_DRAuto_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CancellCall_Panel = New System.Windows.Forms.Panel()
        Me.Spec_CancellCall_Label = New System.Windows.Forms.Label()
        Me.Spec_CancellCall_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_SCOB_Label = New System.Windows.Forms.Label()
        Me.Spec_SCOB_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_AutoFan_Panel = New System.Windows.Forms.Panel()
        Me.Spec_AutoFan_Label = New System.Windows.Forms.Label()
        Me.Spec_AutoFan_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_ION_Label = New System.Windows.Forms.Label()
        Me.Spec_ION_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_AutoPass_Panel = New System.Windows.Forms.Panel()
        Me.Spec_AutoPass_Label = New System.Windows.Forms.Label()
        Me.Spec_AutoPass_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Indep_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Indep_Label = New System.Windows.Forms.Label()
        Me.Spec_Indep_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_HinCpi_Panel = New System.Windows.Forms.Panel()
        Me.Spec_HinCpi_Label = New System.Windows.Forms.Label()
        Me.Spec_HinCpi_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Fire_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Fire_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label195 = New System.Windows.Forms.Label()
        Me.Spec_Fire_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_FireSignal_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_FireSignal_Label = New System.Windows.Forms.Label()
        Me.Spec_Fire_Label = New System.Windows.Forms.Label()
        Me.Spec_Fire_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_EscapeFL_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_EscapeFL_Label = New System.Windows.Forms.Label()
        Me.Spec_Fireman_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Fireman_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label55 = New System.Windows.Forms.Label()
        Me.Spec_Fireman_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_Fireman_Label = New System.Windows.Forms.Label()
        Me.Spec_Fireman_ComboBox = New System.Windows.Forms.ComboBox()
        Me.TabPage10 = New System.Windows.Forms.TabPage()
        Me.Spec_TW_FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Spec_Parking_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Parking_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label56 = New System.Windows.Forms.Label()
        Me.Spec_Parking_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_ParkingFL_DR_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_ParkingFL_DR_Label = New System.Windows.Forms.Label()
        Me.Spec_ParkingFL_HALL_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_ParkingFL_HALL_Label = New System.Windows.Forms.Label()
        Me.Spec_ParkingFL_COB_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_ParkingFL_COB_Label = New System.Windows.Forms.Label()
        Me.Spec_ParkingFL_WTB_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_ParkingFL_WTB_Label = New System.Windows.Forms.Label()
        Me.Spec_ParkingFL_ELVIC_Label = New System.Windows.Forms.Label()
        Me.Spec_ParkingFL_ELVIC_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Parking_FL_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_Parking_FL_Label = New System.Windows.Forms.Label()
        Me.Spec_Parking_Label = New System.Windows.Forms.Label()
        Me.Spec_Parking_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Seismic_Panel = New System.Windows.Forms.Panel()
        Me.Spec_SeismicSW_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label215 = New System.Windows.Forms.Label()
        Me.Spec_SeismicSW_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_SeismicSensor_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label214 = New System.Windows.Forms.Label()
        Me.Spec_SeismicSensor_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_Seismic_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label196 = New System.Windows.Forms.Label()
        Me.Spec_Seismic_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_SeismicSensor_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_SeismicSensor_Label = New System.Windows.Forms.Label()
        Me.Spec_SeismicSW_Label = New System.Windows.Forms.Label()
        Me.Spec_SeismicSW_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Seismic_Label = New System.Windows.Forms.Label()
        Me.Spec_Seismic_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CPI_Panel = New System.Windows.Forms.Panel()
        Me.Spec_CpiOLT_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label216 = New System.Windows.Forms.Label()
        Me.Spec_CpiOLT_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CpiOLT_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CpiOLT_Label = New System.Windows.Forms.Label()
        Me.Spec_CpiFM_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CpiFM_Label = New System.Windows.Forms.Label()
        Me.Spec_CpiEmer_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CpiEmer_Label = New System.Windows.Forms.Label()
        Me.Spec_CpiFire_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CpiFire_Label = New System.Windows.Forms.Label()
        Me.Spec_CpiSeismic_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CpiSeismic_Label = New System.Windows.Forms.Label()
        Me.Spec_CPI_Label = New System.Windows.Forms.Label()
        Me.Spec_CPI_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_HallGong_Panel = New System.Windows.Forms.Panel()
        Me.Spec_HallGong_Label = New System.Windows.Forms.Label()
        Me.Spec_HallGong_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_HPIMsg_Panel = New System.Windows.Forms.Panel()
        Me.Spec_HpiFM_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_HpiFM_Label = New System.Windows.Forms.Label()
        Me.Spec_HpiIndep_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_HpiIndep_Label = New System.Windows.Forms.Label()
        Me.Spec_HpiMain_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_HpiMain_Label = New System.Windows.Forms.Label()
        Me.Spec_HpiOLT_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_HpiOLT_Label = New System.Windows.Forms.Label()
        Me.Spec_HPIMsg_Label = New System.Windows.Forms.Label()
        Me.Spec_HPIMsg_ComboBox = New System.Windows.Forms.ComboBox()
        Me.TabPage12 = New System.Windows.Forms.TabPage()
        Me.Spec_TW_FlowLayoutPanel3 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Spec_CarGong_Panel = New System.Windows.Forms.Panel()
        Me.Spec_CarGong_VONIC_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CarGong_COB_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CarGong_TopBtm_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CarGong_Top_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CarGong_VONIC_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label225 = New System.Windows.Forms.Label()
        Me.Spec_CarGong_VONIC_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CarGong_COB_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label224 = New System.Windows.Forms.Label()
        Me.Spec_CarGong_COB_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CarGong_TopBtm_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label79 = New System.Windows.Forms.Label()
        Me.Spec_CarGong_TopBtm_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CarGong_VONIC_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_CarGong_COB_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_CarGong_TopBtm_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_CarGong_Top_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_CarGong_Top_Only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label223 = New System.Windows.Forms.Label()
        Me.Spec_CarGong_Top_Only_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_CarGong_Label = New System.Windows.Forms.Label()
        Me.Spec_CarGong_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CRD_Panel = New System.Windows.Forms.Panel()
        Me.Spec_CRDType_Label = New System.Windows.Forms.Label()
        Me.Spec_CRDType_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CRDID5_Label = New System.Windows.Forms.Label()
        Me.Spec_CRD_Label = New System.Windows.Forms.Label()
        Me.Spec_CRD_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CRDSpec_Label = New System.Windows.Forms.Label()
        Me.Spec_CRDSpec_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CRDCancell_Label = New System.Windows.Forms.Label()
        Me.Spec_CRDCancell_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CRDNuisance_Label = New System.Windows.Forms.Label()
        Me.Spec_CRDNuisance_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CRDReg_Label = New System.Windows.Forms.Label()
        Me.Spec_CRDReg_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CRDID4_Label = New System.Windows.Forms.Label()
        Me.Spec_CRDID4_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_CRDID5_ComboBox = New System.Windows.Forms.ComboBox()
        Me.TabPage13 = New System.Windows.Forms.TabPage()
        Me.Spec_TW_FlowLayoutPanel4 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Spec_VonicBz_Panel = New System.Windows.Forms.Panel()
        Me.Spec_VonicBz_Label = New System.Windows.Forms.Label()
        Me.Spec_VonicBz_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_DrHold_Panel = New System.Windows.Forms.Panel()
        Me.Spec_DrHold_Label = New System.Windows.Forms.Label()
        Me.Spec_DrHold_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Landic_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Landic_Label = New System.Windows.Forms.Label()
        Me.Spec_Landic_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_MFLReturn_Panel = New System.Windows.Forms.Panel()
        Me.Spec_MFLReturn_Label = New System.Windows.Forms.Label()
        Me.Spec_MFLReturn_FL_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_MFLReturn_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_MFLReturn_FL_Label = New System.Windows.Forms.Label()
        Me.Spec_Vonic_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Vonic_standard_Label = New System.Windows.Forms.Label()
        Me.Spec_Vonic_standard_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Vonic_Label = New System.Windows.Forms.Label()
        Me.Spec_Vonic_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Emer_Panel = New System.Windows.Forms.Panel()
        Me.Spec_EmerNum_NumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Spec_EmerCapacity_Label = New System.Windows.Forms.Label()
        Me.Spec_EmerSignal_Label = New System.Windows.Forms.Label()
        Me.Spec_EmerAddress_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_EmerInput_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_EmerAddress_Label = New System.Windows.Forms.Label()
        Me.Spec_emerGroup_TabControl = New System.Windows.Forms.TabControl()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Spec_EmerNum_Label = New System.Windows.Forms.Label()
        Me.Spec_Emer_Label = New System.Windows.Forms.Label()
        Me.Spec_Emer_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_EmerInput_Label = New System.Windows.Forms.Label()
        Me.Spec_EmerCapacity_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_EmerSignal_ComboBox = New System.Windows.Forms.ComboBox()
        Me.TabPage14 = New System.Windows.Forms.TabPage()
        Me.Spec_TW_FlowLayoutPanel5 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Spec_Elvic_Panel = New System.Windows.Forms.Panel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Spec_Elvic_Label = New System.Windows.Forms.Label()
        Me.Spec_Elvic_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Elvic_Parking_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_VIP_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label202 = New System.Windows.Forms.Label()
        Me.Spec_Elvic_Indep_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_FloorLockOut_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_Express_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_ReturnFL_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_Traffic_Peak_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_MainFL_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label203 = New System.Windows.Forms.Label()
        Me.Spec_Elvic_FloorLockOut_GR_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_Zoning_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_CarCall_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_Traffic_Peak_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Elvic_Fire_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_Elvic_Wavic_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label204 = New System.Windows.Forms.Label()
        Me.Spec_Elvic_CRD_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_WCOB_Panel = New System.Windows.Forms.Panel()
        Me.Spec_WSCOB_only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_WSCOB_only_TextBox = New System.Windows.Forms.TextBox()
        Me.Label227 = New System.Windows.Forms.Label()
        Me.Spec_WCOB_only_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_WCOB_only_TextBox = New System.Windows.Forms.TextBox()
        Me.Label123 = New System.Windows.Forms.Label()
        Me.Spec_WCOB_Label = New System.Windows.Forms.Label()
        Me.Spec_WCOB_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_WSCOB_Label = New System.Windows.Forms.Label()
        Me.Spec_WSCOB_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_WCOB_Ring_Label = New System.Windows.Forms.Label()
        Me.Spec_WCOB_Ring_ComboBox = New System.Windows.Forms.ComboBox()
        Me.TabPage15 = New System.Windows.Forms.TabPage()
        Me.Spec_TW_FlowLayoutPanel6 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Spec_HLL_Panel = New System.Windows.Forms.Panel()
        Me.Spec_HLL_Label = New System.Windows.Forms.Label()
        Me.Spec_HLL_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_ATT_Panel = New System.Windows.Forms.Panel()
        Me.Spec_ATT_Label = New System.Windows.Forms.Label()
        Me.Spec_ATT_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Flood_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Flood_Label = New System.Windows.Forms.Label()
        Me.Spec_Flood_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Flood_FL_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_Flood_FL_Label = New System.Windows.Forms.Label()
        Me.Spec_LS1M_Panel = New System.Windows.Forms.Panel()
        Me.Spec_LS1M_Label = New System.Windows.Forms.Label()
        Me.Spec_LS1M_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_PRU_Panel = New System.Windows.Forms.Panel()
        Me.Spec_PRU_Label = New System.Windows.Forms.Label()
        Me.Spec_PRU_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_LoadCell_Panel = New System.Windows.Forms.Panel()
        Me.Spec_LoadCellPos_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_LoadCellPos_Label = New System.Windows.Forms.Label()
        Me.Spec_LoadCell_Label = New System.Windows.Forms.Label()
        Me.Spec_LoadCell_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_FrontRearDr_Panel = New System.Windows.Forms.Panel()
        Me.Spec_FrontRearDr_Label = New System.Windows.Forms.Label()
        Me.Spec_FrontRearDr_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_OpeSw_Panel = New System.Windows.Forms.Panel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Spec_OpeSw_InputPos_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_OpeSw_InputAddress_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_OpeSw_InputPos_Label = New System.Windows.Forms.Label()
        Me.Spec_OpeSw_DevicePos_TextBox = New System.Windows.Forms.TextBox()
        Me.Spec_OpeSw_DevicePos_Label = New System.Windows.Forms.Label()
        Me.Spec_OpeSw_Label = New System.Windows.Forms.Label()
        Me.Spec_OpeSw_ComboBox = New System.Windows.Forms.ComboBox()
        Me.TabPage11 = New System.Windows.Forms.TabPage()
        Me.Spec_TW_unUse_FlowLayoutPanel = New System.Windows.Forms.FlowLayoutPanel()
        Me.Panel42 = New System.Windows.Forms.Panel()
        Me.Label155 = New System.Windows.Forms.Label()
        Me.Spec_CancellBehind_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Panel43 = New System.Windows.Forms.Panel()
        Me.Label156 = New System.Windows.Forms.Label()
        Me.Spec_LampChk_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Panel54 = New System.Windows.Forms.Panel()
        Me.Label163 = New System.Windows.Forms.Label()
        Me.Spec_CCCancell_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Panel66 = New System.Windows.Forms.Panel()
        Me.Spec_UCMP_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label169 = New System.Windows.Forms.Label()
        Me.Spec_WTB_Panel = New System.Windows.Forms.Panel()
        Me.Label144 = New System.Windows.Forms.Label()
        Me.Spec_WTB_EQMac_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label143 = New System.Windows.Forms.Label()
        Me.Spec_WTB_EQIND_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label142 = New System.Windows.Forms.Label()
        Me.Spec_WTB_Indep_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label141 = New System.Windows.Forms.Label()
        Me.Spec_WTB_EQ_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label140 = New System.Windows.Forms.Label()
        Me.Spec_WTB_Alart_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label137 = New System.Windows.Forms.Label()
        Me.Spec_WTB_BZSW_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label138 = New System.Windows.Forms.Label()
        Me.Spec_WTB_EQSW_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label139 = New System.Windows.Forms.Label()
        Me.Spec_WTB_PKSW_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label133 = New System.Windows.Forms.Label()
        Me.Spec_WTB_EmerPow_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label134 = New System.Windows.Forms.Label()
        Me.Spec_WTB_FO_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label135 = New System.Windows.Forms.Label()
        Me.Spec_WTB_Urgent_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label136 = New System.Windows.Forms.Label()
        Me.Spec_WTB_Normal_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label108 = New System.Windows.Forms.Label()
        Me.Spec_WTB_ChkSW_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label105 = New System.Windows.Forms.Label()
        Me.Spec_WTB_FM_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label102 = New System.Windows.Forms.Label()
        Me.Spec_WTB_Stop_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label98 = New System.Windows.Forms.Label()
        Me.Spec_WTB_Error_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label68 = New System.Windows.Forms.Label()
        Me.Spec_WTB_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_IF79x_Panel = New System.Windows.Forms.Panel()
        Me.Label120 = New System.Windows.Forms.Label()
        Me.Spec_IF79x_IDM0_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label121 = New System.Windows.Forms.Label()
        Me.Spec_IF79x_ID12_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label119 = New System.Windows.Forms.Label()
        Me.Spec_IF79x_ID5_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label118 = New System.Windows.Forms.Label()
        Me.Spec_IF79x_ID4_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label117 = New System.Windows.Forms.Label()
        Me.Spec_IF79x_ID0_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label69 = New System.Windows.Forms.Label()
        Me.Spec_IF79x_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_EachStop_Panel = New System.Windows.Forms.Panel()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.Spec_EachStop_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Panel115 = New System.Windows.Forms.Panel()
        Me.Label_SPEC_INSTALL_OPE = New System.Windows.Forms.Label()
        Me.Spec_install_ope_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Spec_Operation_Panel = New System.Windows.Forms.Panel()
        Me.Spec_Operation_Label = New System.Windows.Forms.Label()
        Me.Spec_Operation_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Use_SpecTWFP17_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Use_SpecTWIDU_CheckBox = New System.Windows.Forms.CheckBox()
        Me.DWG_TabPage = New System.Windows.Forms.TabPage()
        Me.DWG_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label194 = New System.Windows.Forms.Label()
        Me.DWG_VonicStd_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label193 = New System.Windows.Forms.Label()
        Me.Label192 = New System.Windows.Forms.Label()
        Me.DWG_Produce_CheckedListBox = New System.Windows.Forms.CheckedListBox()
        Me.DWG_Construction_CheckedListBox = New System.Windows.Forms.CheckedListBox()
        Me.DWG_StdPage_Button = New System.Windows.Forms.Button()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.DWG_Page_ChkAllButton = New System.Windows.Forms.Button()
        Me.DWG_PageNum_TextBox = New System.Windows.Forms.TextBox()
        Me.DWG_Page_CheckedListBox = New System.Windows.Forms.CheckedListBox()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.DWG_Page_unChkAllButton = New System.Windows.Forms.Button()
        Me.DWG_PrkName_ComboBox = New System.Windows.Forms.ComboBox()
        Me.DWG_Page_SubButton = New System.Windows.Forms.Button()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.DWG_Page_AddButton = New System.Windows.Forms.Button()
        Me.Use_prk_CheckBox = New System.Windows.Forms.CheckBox()
        Me.ProgramChange_TabPage = New System.Windows.Forms.TabPage()
        Me.TabControl3 = New System.Windows.Forms.TabControl()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.ProgramChange_FlowLayoutPanel = New System.Windows.Forms.FlowLayoutPanel()
        Me.use_ProgramChg_Panel1 = New System.Windows.Forms.Panel()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.PrmList_1_reason_TextBox = New System.Windows.Forms.TextBox()
        Me.use_ProgramChg_Panel2 = New System.Windows.Forms.Panel()
        Me.PrmList_2_Other_CheckBox = New System.Windows.Forms.CheckBox()
        Me.PrmList_2_Tower_CheckBox = New System.Windows.Forms.CheckBox()
        Me.PrmList_2_COP_CheckBox = New System.Windows.Forms.CheckBox()
        Me.PrmList_2_test_CheckBox = New System.Windows.Forms.CheckBox()
        Me.PrmList_2_test_TextBox = New System.Windows.Forms.TextBox()
        Me.PrmList_2_COP_TextBox = New System.Windows.Forms.TextBox()
        Me.PrmList_2_tower_TextBox = New System.Windows.Forms.TextBox()
        Me.PrmList_2_other_TextBox = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.use_ProgramChg_Panel3 = New System.Windows.Forms.Panel()
        Me.PrmList_3_debug_CheckBox = New System.Windows.Forms.CheckBox()
        Me.PrmList_3_excute_CheckBox = New System.Windows.Forms.CheckBox()
        Me.PrmList_3_confirm_CheckBox = New System.Windows.Forms.CheckBox()
        Me.PrmList_3_other_Checkbox = New System.Windows.Forms.CheckBox()
        Me.PrmList_3_test_CheckBox = New System.Windows.Forms.CheckBox()
        Me.PrmList_3_other_TextBox = New System.Windows.Forms.TextBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.use_ProgramChg_Panel5 = New System.Windows.Forms.Panel()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.PrmList_5_review_CheckBox = New System.Windows.Forms.CheckBox()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.use_ProgramChg_Panel4 = New System.Windows.Forms.Panel()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes12_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no12_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes8_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no8_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Panel12 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes11_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no11_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes4_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no4_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Panel13 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes10_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no10_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes7_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no7_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.Panel14 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes9_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no9_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes3_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no3_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes6_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no6_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes2_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no2_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes5_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no5_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.PrmList_4_yes1_RadioButton = New System.Windows.Forms.RadioButton()
        Me.PrmList_4_no1_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.PrmList_4_content12_TextBox = New System.Windows.Forms.TextBox()
        Me.Label50 = New System.Windows.Forms.Label()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.Use_Program_CheckBox = New System.Windows.Forms.CheckBox()
        Me.CheckList = New System.Windows.Forms.TabPage()
        Me.CheckList_GroupBox = New System.Windows.Forms.GroupBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.CheckList_FlowLayoutPanel = New System.Windows.Forms.FlowLayoutPanel()
        Me.ChkList_1_Panel = New System.Windows.Forms.Panel()
        Me.ChkList_1_no_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_1_yes_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_1_yes_Content_TextBox = New System.Windows.Forms.TextBox()
        Me.ChkList_1_yes_result_TextBox = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.ChkList_2_Panel = New System.Windows.Forms.Panel()
        Me.ChkList_2_yes_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_2_no_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ChkList_2_yes_Result_TextBox = New System.Windows.Forms.TextBox()
        Me.ChkList_2_yes_Content_TextBox = New System.Windows.Forms.TextBox()
        Me.ChkList_3_Panel = New System.Windows.Forms.Panel()
        Me.ChkList_3_yes_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_3_no_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.ChkList_3_yes_Man_TextBox = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.ChkList_3_yes_Content_TextBox = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.ChkList_3_yes_Result_TextBox = New System.Windows.Forms.TextBox()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.CheckList2_FlowLayoutPanel = New System.Windows.Forms.FlowLayoutPanel()
        Me.ChkList_6_Panel = New System.Windows.Forms.Panel()
        Me.Panel24 = New System.Windows.Forms.Panel()
        Me.ChkList_6_yesItem_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_6_yesChk_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_6_yes_Content_TextBox = New System.Windows.Forms.TextBox()
        Me.ChkList_6_no_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_6_yes_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.ChkList_4_Panel = New System.Windows.Forms.Panel()
        Me.ChkList_4_ObjName_TextBox = New System.Windows.Forms.TextBox()
        Me.ChkList_4_ObjBase_TextBox = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.ChkList_4_SV_TextBox = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.ChkList_4_SVBase_TextBox = New System.Windows.Forms.TextBox()
        Me.ChkList_5_Panel = New System.Windows.Forms.Panel()
        Me.ChkList_5_nstd_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_5_std_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_5_no_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.ChkList_5_std_Content_TextBox = New System.Windows.Forms.TextBox()
        Me.ChkList_5_nstd_Content_TextBox = New System.Windows.Forms.TextBox()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.CheckList3_FlowLayoutPanel = New System.Windows.Forms.FlowLayoutPanel()
        Me.ChkList_7_Panel = New System.Windows.Forms.Panel()
        Me.ChkList_7_yes_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_7_no_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.ChkList_7_yes1_content_TextBox = New System.Windows.Forms.TextBox()
        Me.ChkList_8_Panel = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ChkList_8_yes_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_8_no_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_8Item_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.ChkList_9_Panel = New System.Windows.Forms.Panel()
        Me.ChkList_9_no_RadioButton = New System.Windows.Forms.RadioButton()
        Me.ChkList_9_yes_RadioButton = New System.Windows.Forms.RadioButton()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.ChkList_PaSheet_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ChkList_OS_CheckBox = New System.Windows.Forms.CheckBox()
        Me.ChkList_Elec_DateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.ChkList_Confirm_DateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.ChkList_Confirm_CheckBox = New System.Windows.Forms.CheckBox()
        Me.ChkList_OS_DateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.ChkList_PaSheet_DateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.ChkList_Elec_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Use_ChkList_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Basic_TabPage = New System.Windows.Forms.TabPage()
        Me.Basic_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Basic_Local_Label = New System.Windows.Forms.Label()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Basic_JobNoOld_Label = New System.Windows.Forms.Label()
        Me.Label222 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label221 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label220 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label219 = New System.Windows.Forms.Label()
        Me.Basic_DesingerChinese_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label218 = New System.Windows.Forms.Label()
        Me.Basic_CheckerChinese_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label217 = New System.Windows.Forms.Label()
        Me.Basic_JobNoNew_Label = New System.Windows.Forms.Label()
        Me.Basic_ApproverEnglish_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Basic_Local_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Basic_CheckerEnglish_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Basic_ApproverChinese_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Basic_DesingerEnglish_ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.Basic_DrawDate_DateTimePicker = New System.Windows.Forms.DateTimePicker()
        Me.Basic_JobNoMOD_TextBox = New System.Windows.Forms.TextBox()
        Me.Basic_JobNoNew_TextBox = New System.Windows.Forms.TextBox()
        Me.Basic_JobNoOld_TextBox = New System.Windows.Forms.TextBox()
        Me.Basic_JobName_TextBox = New System.Windows.Forms.TextBox()
        Me.Basic_JobNoMOD_Label = New System.Windows.Forms.Label()
        Me.Use_Basic_CheckBox = New System.Windows.Forms.CheckBox()
        Me.ReminderMarquee_Label = New System.Windows.Forms.Label()
        Me.Load_TabPage = New System.Windows.Forms.TabPage()
        Me.Load_Other_btn_GroupBox = New System.Windows.Forms.GroupBox()
        Me.CheckList_OutputButton = New System.Windows.Forms.Button()
        Me.DWG_OutputButton = New System.Windows.Forms.Button()
        Me.Spec_OutputButton = New System.Windows.Forms.Button()
        Me.Load_SpecDWG_btn_GroupBox = New System.Windows.Forms.GroupBox()
        Me.All_OutputButton = New System.Windows.Forms.Button()
        Me.Load_TabControl = New System.Windows.Forms.TabControl()
        Me.AutoLoad_TabPage = New System.Windows.Forms.TabPage()
        Me.Load_AutoLoad_GroupBox = New System.Windows.Forms.GroupBox()
        Me.JMFileConfirm_AutoLoad_Button = New System.Windows.Forms.Button()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.JMFileCho_AutoLoad_Button = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.JMFileCho_AutoLoad_TextBox = New System.Windows.Forms.TextBox()
        Me.JobMaker_LOAD_AutoLoad_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Spec_TabPage = New System.Windows.Forms.TabPage()
        Me.Load_Spec_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.JM_JobSelect_Spec_ComboBox = New System.Windows.Forms.ComboBox()
        Me.JM_JobSelect_Spec_TextBox = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.JM_DefaultPath_Spec_Label = New System.Windows.Forms.Label()
        Me.Label149 = New System.Windows.Forms.Label()
        Me.JMFileCho_Spec_Button = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.JMFileCho_Spec_TextBox = New System.Windows.Forms.TextBox()
        Me.JobMaker_LOAD_Spec_CheckBox = New System.Windows.Forms.CheckBox()
        Me.CheckList_TabPage = New System.Windows.Forms.TabPage()
        Me.Load_ChkList_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.JM_JobSelect_CheckList_ComboBox = New System.Windows.Forms.ComboBox()
        Me.JM_JobSelect_CheckList_TextBox = New System.Windows.Forms.TextBox()
        Me.JM_DefaultPath_CheckList_Label = New System.Windows.Forms.Label()
        Me.Label173 = New System.Windows.Forms.Label()
        Me.JMFileCho_ChkList_Button = New System.Windows.Forms.Button()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.JMFileCho_ChkList_TextBox = New System.Windows.Forms.TextBox()
        Me.JobMaker_LOAD_ChkList_CheckBox = New System.Windows.Forms.CheckBox()
        Me.LoadSQL_TabPage = New System.Windows.Forms.TabPage()
        Me.Load_SQLite_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label109 = New System.Windows.Forms.Label()
        Me.JM_JobSelect_SQLite_ComboBox = New System.Windows.Forms.ComboBox()
        Me.JM_JobSelect_SQLite_TextBox = New System.Windows.Forms.TextBox()
        Me.JMFileConfirm_SQLite_Button = New System.Windows.Forms.Button()
        Me.JM_DefaultPath_SQLite_Label = New System.Windows.Forms.Label()
        Me.Label188 = New System.Windows.Forms.Label()
        Me.JMFileCho_SQLite_Button = New System.Windows.Forms.Button()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.JMFileCho_SQLite_TextBox = New System.Windows.Forms.TextBox()
        Me.JobMaker_LOAD_SQLite_CheckBox = New System.Windows.Forms.CheckBox()
        Me.JobMaker_TabControl = New System.Windows.Forms.TabControl()
        Me.EepData_TabPage = New System.Windows.Forms.TabPage()
        Me.Use_EepData_CheckBox = New System.Windows.Forms.CheckBox()
        Me.EepData_TabControl = New System.Windows.Forms.TabControl()
        Me.EepData_TabPage1 = New System.Windows.Forms.TabPage()
        Me.EepData_Page1_GroupBox = New System.Windows.Forms.GroupBox()
        Me.EepData_MachineRoom_Label = New System.Windows.Forms.Label()
        Me.EepData_MachineRoom_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Speed_Label = New System.Windows.Forms.Label()
        Me.EepData_Speed_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Capactity_Label = New System.Windows.Forms.Label()
        Me.EepData_Capactity_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_TopFL_Label = New System.Windows.Forms.Label()
        Me.EepData_TopFL_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_BtmFL_Label = New System.Windows.Forms.Label()
        Me.EepData_BtmFL_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_StopFL_Label = New System.Windows.Forms.Label()
        Me.EepData_StopFL_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_OpeType_Label = New System.Windows.Forms.Label()
        Me.EepData_OpeType_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_GspType_Label = New System.Windows.Forms.Label()
        Me.EepData_GspType_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Purpose_Label = New System.Windows.Forms.Label()
        Me.EepData_Purpose_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_GroupNo_Label = New System.Windows.Forms.Label()
        Me.EepData_GroupNo_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_CarNo_Label = New System.Windows.Forms.Label()
        Me.EepData_CarNo_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_DrCloser_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_DrCloser_Label = New System.Windows.Forms.Label()
        Me.EepData_DrType_Label = New System.Windows.Forms.Label()
        Me.EepData_DrType_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_TabPage2 = New System.Windows.Forms.TabPage()
        Me.EepData_Page2_GroupBox = New System.Windows.Forms.GroupBox()
        Me.Label80 = New System.Windows.Forms.Label()
        Me.Label78 = New System.Windows.Forms.Label()
        Me.Label77 = New System.Windows.Forms.Label()
        Me.EepData_DrFrontWidth_Label = New System.Windows.Forms.Label()
        Me.EepData_DrFrontWidth_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Landic_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Landic_Label = New System.Windows.Forms.Label()
        Me.EepData_EnergyRe_Label = New System.Windows.Forms.Label()
        Me.EepData_EnergyRe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Indep_Label = New System.Windows.Forms.Label()
        Me.EepData_Indep_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_SpecMainFL_Label = New System.Windows.Forms.Label()
        Me.EepData_SpecMainFL_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_MainFL_Label = New System.Windows.Forms.Label()
        Me.EepData_MainFL_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_MainFL_FR_Label = New System.Windows.Forms.Label()
        Me.EepData_MainFL_FR_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Seismic_Label = New System.Windows.Forms.Label()
        Me.EepData_Seismic_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Nudging_Label = New System.Windows.Forms.Label()
        Me.EepData_Nudging_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_AutoByPass_Label = New System.Windows.Forms.Label()
        Me.EepData_AutoByPass_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_AutoFan_Label = New System.Windows.Forms.Label()
        Me.EepData_AutoFan_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_DrHold_Label = New System.Windows.Forms.Label()
        Me.EepData_DrHold_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_TabPage3 = New System.Windows.Forms.TabPage()
        Me.EepData_Page3_GroupBox = New System.Windows.Forms.GroupBox()
        Me.EepData_DrCloseBtn_Label = New System.Windows.Forms.Label()
        Me.EepData_DrCloseBtn_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_CarChime_Label = New System.Windows.Forms.Label()
        Me.EepData_CarChime_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_HallChime_Label = New System.Windows.Forms.Label()
        Me.EepData_HallChime_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_ParkingOpe_Label = New System.Windows.Forms.Label()
        Me.EepData_ParkingOpe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_ParkingSW_Label = New System.Windows.Forms.Label()
        Me.EepData_ParkingSW_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_ParkingFL_Label = New System.Windows.Forms.Label()
        Me.EepData_ParkingFL_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_ParkingFL_ForR_Label = New System.Windows.Forms.Label()
        Me.EepData_ParkingFL_ForR_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_EscapeOpe_ForR_Label = New System.Windows.Forms.Label()
        Me.EepData_EscapeOpe_ForR_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_PhotoEye_Label = New System.Windows.Forms.Label()
        Me.EepData_PhotoEye_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_SafetyShoe_Label = New System.Windows.Forms.Label()
        Me.EepData_SafetyShoe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_EscapeOpe_Label = New System.Windows.Forms.Label()
        Me.EepData_EscapeOpe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_EscapeFL_Label = New System.Windows.Forms.Label()
        Me.EepData_EscapeFL_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Overbalance_Label = New System.Windows.Forms.Label()
        Me.EepData_Overbalance_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_TabPage4 = New System.Windows.Forms.TabPage()
        Me.EepData_Page4_GroupBox = New System.Windows.Forms.GroupBox()
        Me.EepData_SheaveDia_Label = New System.Windows.Forms.Label()
        Me.EepData_SheaveDia_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_MachineType_Label = New System.Windows.Forms.Label()
        Me.EepData_MachineType_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Gear_Label = New System.Windows.Forms.Label()
        Me.EepData_Gear_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Inverter_Label = New System.Windows.Forms.Label()
        Me.EepData_Inverter_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_MotorPole_Label = New System.Windows.Forms.Label()
        Me.EepData_MotorPole_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_MotorVoltage_Label = New System.Windows.Forms.Label()
        Me.EepData_MotorVoltage_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_MotorCapacity_Label = New System.Windows.Forms.Label()
        Me.EepData_MotorCapacity_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_MotorDirection_Label = New System.Windows.Forms.Label()
        Me.EepData_MotorDirection_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Encoder_Label = New System.Windows.Forms.Label()
        Me.EepData_Encoder_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_FireOpe_Label = New System.Windows.Forms.Label()
        Me.EepData_FireOpe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_FMSOpe_Label = New System.Windows.Forms.Label()
        Me.EepData_FMSOpe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_FMSSW_Label = New System.Windows.Forms.Label()
        Me.EepData_FMSSW_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_EmerOpe_Label = New System.Windows.Forms.Label()
        Me.EepData_EmerOpe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_TabPage5 = New System.Windows.Forms.TabPage()
        Me.EepData_Page5_GroupBox = New System.Windows.Forms.GroupBox()
        Me.EepData_FloodOpe_Label = New System.Windows.Forms.Label()
        Me.EepData_FloodOpe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Vonic_Label = New System.Windows.Forms.Label()
        Me.EepData_Vonic_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_HIN1_Label = New System.Windows.Forms.Label()
        Me.EepData_HIN1_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_HIN2_Label = New System.Windows.Forms.Label()
        Me.EepData_HIN2_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_HIN3_Label = New System.Windows.Forms.Label()
        Me.EepData_HIN3_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_HIN4_Label = New System.Windows.Forms.Label()
        Me.EepData_HIN4_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_SCOB_Label = New System.Windows.Forms.Label()
        Me.EepData_SCOB_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_WCOB_Label = New System.Windows.Forms.Label()
        Me.EepData_WCOB_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_WSCOB_Label = New System.Windows.Forms.Label()
        Me.EepData_WSCOB_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_WCOB_Spec_Label = New System.Windows.Forms.Label()
        Me.EepData_WCOB_Spec_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_FRDr_Label = New System.Windows.Forms.Label()
        Me.EepData_FRDr_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_AttOpe_Label = New System.Windows.Forms.Label()
        Me.EepData_AttOpe_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Rope_Label = New System.Windows.Forms.Label()
        Me.EepData_Rope_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_TabPage6 = New System.Windows.Forms.TabPage()
        Me.EepData_Page6_GroupBox = New System.Windows.Forms.GroupBox()
        Me.EepData_Travel_Label = New System.Windows.Forms.Label()
        Me.EepData_Travel_TextBox = New System.Windows.Forms.TextBox()
        Me.EepData_Hight_Label = New System.Windows.Forms.Label()
        Me.EepData_Hight_TextBox = New System.Windows.Forms.TextBox()
        Me.FinalCheck_TabPage = New System.Windows.Forms.TabPage()
        Me.FinalCheck_Button = New System.Windows.Forms.Button()
        Me.ResultFailOutput_TextBox = New System.Windows.Forms.TextBox()
        Me.JobMaker_Close_Button = New System.Windows.Forms.Button()
        Me.EntityCommand1 = New System.Data.Entity.Core.EntityClient.EntityCommand()
        Me.G_TabPage.SuspendLayout
        Me.GWeb_GroupBox.SuspendLayout
        Me.MMIC_TabPage.SuspendLayout
        Me.MMIC_Panel.SuspendLayout
        Me.Panel17.SuspendLayout
        Me.Panel15.SuspendLayout
        Me.MMIC_VD10_GroupBox.SuspendLayout
        CType(Me.MMIC_VD10_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.MMIC_SV_E_GroupBox.SuspendLayout
        CType(Me.MMIC_SV_E_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.MMIC_SV_GroupBox.SuspendLayout
        CType(Me.MMIC_SV_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.MMIC_MR_E_GroupBox.SuspendLayout
        CType(Me.MMIC_MR_E_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.MMIC_MR_GroupBox.SuspendLayout
        CType(Me.MMIC_MR_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.MMIC_GroupBox.SuspendLayout
        Me.Important_TabPage.SuspendLayout
        Me.ImpSetting_GroupBox.SuspendLayout
        Me.Spec.SuspendLayout
        Me.Spec_TabControl.SuspendLayout
        Me.Spec_BasicAll_TabPage.SuspendLayout
        Me.Spec_BasicAll_TabControl.SuspendLayout
        Me.TabPage7.SuspendLayout
        Me.SpecBasic_GroupBox.SuspendLayout
        CType(Me.Spec_LiftNum_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.SpecBasic_LiftItem_Panel.SuspendLayout
        Me.TabPage8.SuspendLayout
        Me.SpecBasic_GroupBox2.SuspendLayout
        CType(Me.Spec_MachineType_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.SpecBasic_p2_base_Panel.SuspendLayout
        CType(Me.Spec_Purpose_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.Spec_TW_TabPage.SuspendLayout
        Me.Spec_TW_TabControl.SuspendLayout
        Me.TabPage9.SuspendLayout
        Me.Spec_TW_FlowLayoutPanel1.SuspendLayout
        Me.Spec_DRAuto_Panel.SuspendLayout
        Me.Spec_CancellCall_Panel.SuspendLayout
        Me.Spec_AutoFan_Panel.SuspendLayout
        Me.Spec_AutoPass_Panel.SuspendLayout
        Me.Spec_Indep_Panel.SuspendLayout
        Me.Spec_HinCpi_Panel.SuspendLayout
        Me.Spec_Fire_Panel.SuspendLayout
        Me.Spec_Fireman_Panel.SuspendLayout
        Me.TabPage10.SuspendLayout
        Me.Spec_TW_FlowLayoutPanel2.SuspendLayout
        Me.Spec_Parking_Panel.SuspendLayout
        Me.Spec_Seismic_Panel.SuspendLayout
        Me.Spec_CPI_Panel.SuspendLayout
        Me.Spec_HallGong_Panel.SuspendLayout
        Me.Spec_HPIMsg_Panel.SuspendLayout
        Me.TabPage12.SuspendLayout
        Me.Spec_TW_FlowLayoutPanel3.SuspendLayout
        Me.Spec_CarGong_Panel.SuspendLayout
        Me.Spec_CRD_Panel.SuspendLayout
        Me.TabPage13.SuspendLayout
        Me.Spec_TW_FlowLayoutPanel4.SuspendLayout
        Me.Spec_VonicBz_Panel.SuspendLayout
        Me.Spec_DrHold_Panel.SuspendLayout
        Me.Spec_Landic_Panel.SuspendLayout
        Me.Spec_MFLReturn_Panel.SuspendLayout
        Me.Spec_Vonic_Panel.SuspendLayout
        Me.Spec_Emer_Panel.SuspendLayout
        CType(Me.Spec_EmerNum_NumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit
        Me.Spec_emerGroup_TabControl.SuspendLayout
        Me.TabPage14.SuspendLayout
        Me.Spec_TW_FlowLayoutPanel5.SuspendLayout
        Me.Spec_Elvic_Panel.SuspendLayout
        Me.Spec_WCOB_Panel.SuspendLayout
        Me.TabPage15.SuspendLayout
        Me.Spec_TW_FlowLayoutPanel6.SuspendLayout
        Me.Spec_HLL_Panel.SuspendLayout
        Me.Spec_ATT_Panel.SuspendLayout
        Me.Spec_Flood_Panel.SuspendLayout
        Me.Spec_LS1M_Panel.SuspendLayout
        Me.Spec_PRU_Panel.SuspendLayout
        Me.Spec_LoadCell_Panel.SuspendLayout
        Me.Spec_FrontRearDr_Panel.SuspendLayout
        Me.Spec_OpeSw_Panel.SuspendLayout
        Me.TabPage11.SuspendLayout
        Me.Spec_TW_unUse_FlowLayoutPanel.SuspendLayout
        Me.Panel42.SuspendLayout
        Me.Panel43.SuspendLayout
        Me.Panel54.SuspendLayout
        Me.Panel66.SuspendLayout
        Me.Spec_WTB_Panel.SuspendLayout
        Me.Spec_IF79x_Panel.SuspendLayout
        Me.Spec_EachStop_Panel.SuspendLayout
        Me.Panel115.SuspendLayout
        Me.Spec_Operation_Panel.SuspendLayout
        Me.DWG_TabPage.SuspendLayout
        Me.DWG_GroupBox.SuspendLayout
        Me.ProgramChange_TabPage.SuspendLayout
        Me.TabControl3.SuspendLayout
        Me.TabPage5.SuspendLayout
        Me.ProgramChange_FlowLayoutPanel.SuspendLayout
        Me.use_ProgramChg_Panel1.SuspendLayout
        Me.use_ProgramChg_Panel2.SuspendLayout
        Me.use_ProgramChg_Panel3.SuspendLayout
        Me.use_ProgramChg_Panel5.SuspendLayout
        Me.TabPage6.SuspendLayout
        Me.FlowLayoutPanel1.SuspendLayout
        Me.use_ProgramChg_Panel4.SuspendLayout
        Me.Panel11.SuspendLayout
        Me.Panel7.SuspendLayout
        Me.Panel12.SuspendLayout
        Me.Panel6.SuspendLayout
        Me.Panel13.SuspendLayout
        Me.Panel8.SuspendLayout
        Me.Panel14.SuspendLayout
        Me.Panel5.SuspendLayout
        Me.Panel9.SuspendLayout
        Me.Panel4.SuspendLayout
        Me.Panel10.SuspendLayout
        Me.Panel3.SuspendLayout
        Me.CheckList.SuspendLayout
        Me.CheckList_GroupBox.SuspendLayout
        Me.TabControl1.SuspendLayout
        Me.TabPage1.SuspendLayout
        Me.CheckList_FlowLayoutPanel.SuspendLayout
        Me.ChkList_1_Panel.SuspendLayout
        Me.ChkList_2_Panel.SuspendLayout
        Me.ChkList_3_Panel.SuspendLayout
        Me.TabPage3.SuspendLayout
        Me.CheckList2_FlowLayoutPanel.SuspendLayout
        Me.ChkList_6_Panel.SuspendLayout
        Me.Panel24.SuspendLayout
        Me.ChkList_4_Panel.SuspendLayout
        Me.ChkList_5_Panel.SuspendLayout
        Me.TabPage4.SuspendLayout
        Me.CheckList3_FlowLayoutPanel.SuspendLayout
        Me.ChkList_7_Panel.SuspendLayout
        Me.ChkList_8_Panel.SuspendLayout
        Me.Panel1.SuspendLayout
        Me.ChkList_9_Panel.SuspendLayout
        Me.Basic_TabPage.SuspendLayout
        Me.Basic_GroupBox.SuspendLayout
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit
        Me.Load_TabPage.SuspendLayout
        Me.Load_Other_btn_GroupBox.SuspendLayout
        Me.Load_SpecDWG_btn_GroupBox.SuspendLayout
        Me.Load_TabControl.SuspendLayout
        Me.AutoLoad_TabPage.SuspendLayout
        Me.Load_AutoLoad_GroupBox.SuspendLayout
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit
        Me.Spec_TabPage.SuspendLayout
        Me.Load_Spec_GroupBox.SuspendLayout
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit
        Me.CheckList_TabPage.SuspendLayout
        Me.Load_ChkList_GroupBox.SuspendLayout
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit
        Me.LoadSQL_TabPage.SuspendLayout
        Me.Load_SQLite_GroupBox.SuspendLayout
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit
        Me.JobMaker_TabControl.SuspendLayout
        Me.EepData_TabPage.SuspendLayout
        Me.EepData_TabControl.SuspendLayout
        Me.EepData_TabPage1.SuspendLayout
        Me.EepData_Page1_GroupBox.SuspendLayout
        Me.EepData_TabPage2.SuspendLayout
        Me.EepData_Page2_GroupBox.SuspendLayout
        Me.EepData_TabPage3.SuspendLayout
        Me.EepData_Page3_GroupBox.SuspendLayout
        Me.EepData_TabPage4.SuspendLayout
        Me.EepData_Page4_GroupBox.SuspendLayout
        Me.EepData_TabPage5.SuspendLayout
        Me.EepData_Page5_GroupBox.SuspendLayout
        Me.EepData_TabPage6.SuspendLayout
        Me.EepData_Page6_GroupBox.SuspendLayout
        Me.FinalCheck_TabPage.SuspendLayout
        Me.SuspendLayout
        '
        'ResultCheck_Button
        '
        Me.ResultCheck_Button.Location = New System.Drawing.Point(1045, 606)
        Me.ResultCheck_Button.Name = "ResultCheck_Button"
        Me.ResultCheck_Button.Size = New System.Drawing.Size(75, 23)
        Me.ResultCheck_Button.TabIndex = 7
        Me.ResultCheck_Button.Text = "關閉結果"
        Me.ResultCheck_Button.UseVisualStyleBackColor = True
        '
        'ResultOutput_TextBox
        '
        Me.ResultOutput_TextBox.Location = New System.Drawing.Point(708, 40)
        Me.ResultOutput_TextBox.Multiline = True
        Me.ResultOutput_TextBox.Name = "ResultOutput_TextBox"
        Me.ResultOutput_TextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.ResultOutput_TextBox.Size = New System.Drawing.Size(412, 227)
        Me.ResultOutput_TextBox.TabIndex = 6
        '
        'JobMaker_Timer
        '
        '
        'G_TabPage
        '
        Me.G_TabPage.Controls.Add(Me.GWeb_GroupBox)
        Me.G_TabPage.Controls.Add(Me.Use_G_CheckBox)
        Me.G_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.G_TabPage.Name = "G_TabPage"
        Me.G_TabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.G_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.G_TabPage.TabIndex = 8
        Me.G_TabPage.Text = "G值"
        Me.G_TabPage.UseVisualStyleBackColor = True
        '
        'GWeb_GroupBox
        '
        Me.GWeb_GroupBox.Controls.Add(Me.Label86)
        Me.GWeb_GroupBox.Controls.Add(Me.GWeb_Button)
        Me.GWeb_GroupBox.Enabled = False
        Me.GWeb_GroupBox.Location = New System.Drawing.Point(23, 17)
        Me.GWeb_GroupBox.Name = "GWeb_GroupBox"
        Me.GWeb_GroupBox.Size = New System.Drawing.Size(620, 547)
        Me.GWeb_GroupBox.TabIndex = 41
        Me.GWeb_GroupBox.TabStop = False
        '
        'Label86
        '
        Me.Label86.AutoSize = True
        Me.Label86.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label86.Location = New System.Drawing.Point(15, 19)
        Me.Label86.Name = "Label86"
        Me.Label86.Size = New System.Drawing.Size(52, 16)
        Me.Label86.TabIndex = 39
        Me.Label86.Text = "G_web :"
        '
        'GWeb_Button
        '
        Me.GWeb_Button.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.GWeb_Button.Location = New System.Drawing.Point(15, 38)
        Me.GWeb_Button.Name = "GWeb_Button"
        Me.GWeb_Button.Size = New System.Drawing.Size(148, 106)
        Me.GWeb_Button.TabIndex = 0
        Me.GWeb_Button.Text = "gogogo"
        Me.GWeb_Button.UseVisualStyleBackColor = True
        '
        'Use_G_CheckBox
        '
        Me.Use_G_CheckBox.AutoSize = True
        Me.Use_G_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_G_CheckBox.Name = "Use_G_CheckBox"
        Me.Use_G_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Use_G_CheckBox.TabIndex = 40
        Me.Use_G_CheckBox.UseVisualStyleBackColor = True
        '
        'MMIC_TabPage
        '
        Me.MMIC_TabPage.Controls.Add(Me.MMIC_Panel)
        Me.MMIC_TabPage.Controls.Add(Me.MMIC_GroupBox)
        Me.MMIC_TabPage.Controls.Add(Me.Use_mmic_CheckBox)
        Me.MMIC_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.MMIC_TabPage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_TabPage.Name = "MMIC_TabPage"
        Me.MMIC_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.MMIC_TabPage.TabIndex = 5
        Me.MMIC_TabPage.Text = "MMIC"
        Me.MMIC_TabPage.UseVisualStyleBackColor = True
        '
        'MMIC_Panel
        '
        Me.MMIC_Panel.AutoScroll = True
        Me.MMIC_Panel.Controls.Add(Me.Panel17)
        Me.MMIC_Panel.Controls.Add(Me.Panel15)
        Me.MMIC_Panel.Controls.Add(Me.MMIC_VD10_GroupBox)
        Me.MMIC_Panel.Controls.Add(Me.MMIC_SV_E_GroupBox)
        Me.MMIC_Panel.Controls.Add(Me.MMIC_SV_GroupBox)
        Me.MMIC_Panel.Controls.Add(Me.MMIC_MR_E_GroupBox)
        Me.MMIC_Panel.Controls.Add(Me.MMIC_MR_GroupBox)
        Me.MMIC_Panel.Enabled = False
        Me.MMIC_Panel.Location = New System.Drawing.Point(3, 63)
        Me.MMIC_Panel.Name = "MMIC_Panel"
        Me.MMIC_Panel.Size = New System.Drawing.Size(650, 516)
        Me.MMIC_Panel.TabIndex = 46
        '
        'Panel17
        '
        Me.Panel17.AutoScroll = True
        Me.Panel17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel17.Controls.Add(Me.mmicType1_ObjNameBase_TextBox)
        Me.Panel17.Controls.Add(Me.mmicType1_ObjName_TextBox)
        Me.Panel17.Controls.Add(Me.mmicType1_CarNo_TextBox)
        Me.Panel17.Location = New System.Drawing.Point(322, 769)
        Me.Panel17.Name = "Panel17"
        Me.Panel17.Size = New System.Drawing.Size(300, 58)
        Me.Panel17.TabIndex = 56
        Me.Panel17.Visible = False
        '
        'mmicType1_ObjNameBase_TextBox
        '
        Me.mmicType1_ObjNameBase_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.mmicType1_ObjNameBase_TextBox.Location = New System.Drawing.Point(184, 10)
        Me.mmicType1_ObjNameBase_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.mmicType1_ObjNameBase_TextBox.MaxLength = 50
        Me.mmicType1_ObjNameBase_TextBox.Name = "mmicType1_ObjNameBase_TextBox"
        Me.mmicType1_ObjNameBase_TextBox.Size = New System.Drawing.Size(100, 23)
        Me.mmicType1_ObjNameBase_TextBox.TabIndex = 6
        Me.mmicType1_ObjNameBase_TextBox.Text = "TJAMG11A(樣板)"
        Me.mmicType1_ObjNameBase_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.mmicType1_ObjNameBase_TextBox.Visible = False
        '
        'mmicType1_ObjName_TextBox
        '
        Me.mmicType1_ObjName_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.mmicType1_ObjName_TextBox.Location = New System.Drawing.Point(78, 10)
        Me.mmicType1_ObjName_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.mmicType1_ObjName_TextBox.MaxLength = 50
        Me.mmicType1_ObjName_TextBox.Name = "mmicType1_ObjName_TextBox"
        Me.mmicType1_ObjName_TextBox.Size = New System.Drawing.Size(100, 23)
        Me.mmicType1_ObjName_TextBox.TabIndex = 5
        Me.mmicType1_ObjName_TextBox.Text = "TJAMG11A(樣板)"
        Me.mmicType1_ObjName_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.mmicType1_ObjName_TextBox.Visible = False
        '
        'mmicType1_CarNo_TextBox
        '
        Me.mmicType1_CarNo_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.mmicType1_CarNo_TextBox.Location = New System.Drawing.Point(10, 10)
        Me.mmicType1_CarNo_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.mmicType1_CarNo_TextBox.MaxLength = 50
        Me.mmicType1_CarNo_TextBox.Name = "mmicType1_CarNo_TextBox"
        Me.mmicType1_CarNo_TextBox.Size = New System.Drawing.Size(62, 23)
        Me.mmicType1_CarNo_TextBox.TabIndex = 3
        Me.mmicType1_CarNo_TextBox.Text = "L#1(樣板)"
        Me.mmicType1_CarNo_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.mmicType1_CarNo_TextBox.Visible = False
        '
        'Panel15
        '
        Me.Panel15.AutoScroll = True
        Me.Panel15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel15.Controls.Add(Me.mmic_ObjName_TextBox)
        Me.Panel15.Controls.Add(Me.mmic_CarNo_TextBox)
        Me.Panel15.Location = New System.Drawing.Point(322, 833)
        Me.Panel15.Name = "Panel15"
        Me.Panel15.Size = New System.Drawing.Size(300, 54)
        Me.Panel15.TabIndex = 55
        Me.Panel15.Visible = False
        '
        'mmic_ObjName_TextBox
        '
        Me.mmic_ObjName_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.mmic_ObjName_TextBox.Location = New System.Drawing.Point(95, 10)
        Me.mmic_ObjName_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.mmic_ObjName_TextBox.MaxLength = 50
        Me.mmic_ObjName_TextBox.Name = "mmic_ObjName_TextBox"
        Me.mmic_ObjName_TextBox.Size = New System.Drawing.Size(130, 23)
        Me.mmic_ObjName_TextBox.TabIndex = 5
        Me.mmic_ObjName_TextBox.Text = "TJAMG11A(樣板)"
        Me.mmic_ObjName_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.mmic_ObjName_TextBox.Visible = False
        '
        'mmic_CarNo_TextBox
        '
        Me.mmic_CarNo_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.mmic_CarNo_TextBox.Location = New System.Drawing.Point(10, 10)
        Me.mmic_CarNo_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.mmic_CarNo_TextBox.MaxLength = 50
        Me.mmic_CarNo_TextBox.Name = "mmic_CarNo_TextBox"
        Me.mmic_CarNo_TextBox.Size = New System.Drawing.Size(62, 23)
        Me.mmic_CarNo_TextBox.TabIndex = 3
        Me.mmic_CarNo_TextBox.Text = "L#1(樣板)"
        Me.mmic_CarNo_TextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.mmic_CarNo_TextBox.Visible = False
        '
        'MMIC_VD10_GroupBox
        '
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.MMIC_VD10_NumericUpDown)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.MMIC_VD10_Base_TextBox)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.MMIC_VD10_Type_ComboBox)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.Label132)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.Label131)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.Label114)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.Label115)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.Label113)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.MMIC_VD10_Panel)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.Label65)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.MMIC_VD10_ROM_ComboBox)
        Me.MMIC_VD10_GroupBox.Controls.Add(Me.MMIC_VD10_Quantity_ComboBox)
        Me.MMIC_VD10_GroupBox.Location = New System.Drawing.Point(6, 557)
        Me.MMIC_VD10_GroupBox.Name = "MMIC_VD10_GroupBox"
        Me.MMIC_VD10_GroupBox.Size = New System.Drawing.Size(310, 350)
        Me.MMIC_VD10_GroupBox.TabIndex = 54
        Me.MMIC_VD10_GroupBox.TabStop = False
        Me.MMIC_VD10_GroupBox.Text = "VONIC ROM(VD10x)."
        '
        'MMIC_VD10_NumericUpDown
        '
        Me.MMIC_VD10_NumericUpDown.Location = New System.Drawing.Point(245, 110)
        Me.MMIC_VD10_NumericUpDown.Name = "MMIC_VD10_NumericUpDown"
        Me.MMIC_VD10_NumericUpDown.ReadOnly = True
        Me.MMIC_VD10_NumericUpDown.Size = New System.Drawing.Size(52, 23)
        Me.MMIC_VD10_NumericUpDown.TabIndex = 62
        '
        'MMIC_VD10_Base_TextBox
        '
        Me.MMIC_VD10_Base_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.MMIC_VD10_Base_TextBox.Location = New System.Drawing.Point(45, 110)
        Me.MMIC_VD10_Base_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_VD10_Base_TextBox.MaxLength = 50
        Me.MMIC_VD10_Base_TextBox.Name = "MMIC_VD10_Base_TextBox"
        Me.MMIC_VD10_Base_TextBox.Size = New System.Drawing.Size(140, 23)
        Me.MMIC_VD10_Base_TextBox.TabIndex = 61
        '
        'MMIC_VD10_Type_ComboBox
        '
        Me.MMIC_VD10_Type_ComboBox.FormattingEnabled = True
        Me.MMIC_VD10_Type_ComboBox.Location = New System.Drawing.Point(45, 78)
        Me.MMIC_VD10_Type_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_VD10_Type_ComboBox.Name = "MMIC_VD10_Type_ComboBox"
        Me.MMIC_VD10_Type_ComboBox.Size = New System.Drawing.Size(241, 24)
        Me.MMIC_VD10_Type_ComboBox.TabIndex = 60
        '
        'Label132
        '
        Me.Label132.AutoSize = True
        Me.Label132.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label132.Location = New System.Drawing.Point(6, 82)
        Me.Label132.Name = "Label132"
        Me.Label132.Size = New System.Drawing.Size(36, 16)
        Me.Label132.TabIndex = 59
        Me.Label132.Text = "TYPE"
        '
        'Label131
        '
        Me.Label131.AutoSize = True
        Me.Label131.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label131.Location = New System.Drawing.Point(6, 113)
        Me.Label131.Name = "Label131"
        Me.Label131.Size = New System.Drawing.Size(37, 16)
        Me.Label131.TabIndex = 57
        Me.Label131.Text = "BASE"
        '
        'Label114
        '
        Me.Label114.AutoSize = True
        Me.Label114.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label114.Location = New System.Drawing.Point(115, 145)
        Me.Label114.Name = "Label114"
        Me.Label114.Size = New System.Drawing.Size(87, 16)
        Me.Label114.TabIndex = 4
        Me.Label114.Text = "Object Name."
        '
        'Label115
        '
        Me.Label115.AutoSize = True
        Me.Label115.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label115.Location = New System.Drawing.Point(23, 145)
        Me.Label115.Name = "Label115"
        Me.Label115.Size = New System.Drawing.Size(51, 16)
        Me.Label115.TabIndex = 2
        Me.Label115.Text = "Car No."
        '
        'Label113
        '
        Me.Label113.AutoSize = True
        Me.Label113.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label113.Location = New System.Drawing.Point(6, 20)
        Me.Label113.Name = "Label113"
        Me.Label113.Size = New System.Drawing.Size(83, 16)
        Me.Label113.TabIndex = 16
        Me.Label113.Text = "ROM DEVICE"
        '
        'MMIC_VD10_Panel
        '
        Me.MMIC_VD10_Panel.AutoScroll = True
        Me.MMIC_VD10_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MMIC_VD10_Panel.Location = New System.Drawing.Point(5, 170)
        Me.MMIC_VD10_Panel.Name = "MMIC_VD10_Panel"
        Me.MMIC_VD10_Panel.Size = New System.Drawing.Size(300, 160)
        Me.MMIC_VD10_Panel.TabIndex = 17
        '
        'Label65
        '
        Me.Label65.AutoSize = True
        Me.Label65.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label65.Location = New System.Drawing.Point(6, 50)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(88, 16)
        Me.Label65.TabIndex = 26
        Me.Label65.Text = "Quantity(幾片)"
        '
        'MMIC_VD10_ROM_ComboBox
        '
        Me.MMIC_VD10_ROM_ComboBox.FormattingEnabled = True
        Me.MMIC_VD10_ROM_ComboBox.Items.AddRange(New Object() {"4Mb", "8Mb"})
        Me.MMIC_VD10_ROM_ComboBox.Location = New System.Drawing.Point(104, 16)
        Me.MMIC_VD10_ROM_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_VD10_ROM_ComboBox.Name = "MMIC_VD10_ROM_ComboBox"
        Me.MMIC_VD10_ROM_ComboBox.Size = New System.Drawing.Size(55, 24)
        Me.MMIC_VD10_ROM_ComboBox.TabIndex = 46
        '
        'MMIC_VD10_Quantity_ComboBox
        '
        Me.MMIC_VD10_Quantity_ComboBox.FormattingEnabled = True
        Me.MMIC_VD10_Quantity_ComboBox.Items.AddRange(New Object() {"1", "2", "-"})
        Me.MMIC_VD10_Quantity_ComboBox.Location = New System.Drawing.Point(104, 46)
        Me.MMIC_VD10_Quantity_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_VD10_Quantity_ComboBox.Name = "MMIC_VD10_Quantity_ComboBox"
        Me.MMIC_VD10_Quantity_ComboBox.Size = New System.Drawing.Size(55, 24)
        Me.MMIC_VD10_Quantity_ComboBox.TabIndex = 47
        '
        'MMIC_SV_E_GroupBox
        '
        Me.MMIC_SV_E_GroupBox.Controls.Add(Me.MMIC_SV_E_NumericUpDown)
        Me.MMIC_SV_E_GroupBox.Controls.Add(Me.MMIC_SV_ECarObj_ComboBox)
        Me.MMIC_SV_E_GroupBox.Controls.Add(Me.Label106)
        Me.MMIC_SV_E_GroupBox.Controls.Add(Me.MMIC_SV_E_Panel)
        Me.MMIC_SV_E_GroupBox.Controls.Add(Me.Label107)
        Me.MMIC_SV_E_GroupBox.Controls.Add(Me.Label63)
        Me.MMIC_SV_E_GroupBox.Controls.Add(Me.MMIC_SV_EBase_ComboBox)
        Me.MMIC_SV_E_GroupBox.Location = New System.Drawing.Point(320, 281)
        Me.MMIC_SV_E_GroupBox.Name = "MMIC_SV_E_GroupBox"
        Me.MMIC_SV_E_GroupBox.Size = New System.Drawing.Size(310, 265)
        Me.MMIC_SV_E_GroupBox.TabIndex = 53
        Me.MMIC_SV_E_GroupBox.TabStop = False
        Me.MMIC_SV_E_GroupBox.Text = "MAIN COMPUTER EEPROM DATA."
        '
        'MMIC_SV_E_NumericUpDown
        '
        Me.MMIC_SV_E_NumericUpDown.Location = New System.Drawing.Point(253, 16)
        Me.MMIC_SV_E_NumericUpDown.Name = "MMIC_SV_E_NumericUpDown"
        Me.MMIC_SV_E_NumericUpDown.ReadOnly = True
        Me.MMIC_SV_E_NumericUpDown.Size = New System.Drawing.Size(52, 23)
        Me.MMIC_SV_E_NumericUpDown.TabIndex = 59
        '
        'MMIC_SV_ECarObj_ComboBox
        '
        Me.MMIC_SV_ECarObj_ComboBox.FormattingEnabled = True
        Me.MMIC_SV_ECarObj_ComboBox.Location = New System.Drawing.Point(52, 49)
        Me.MMIC_SV_ECarObj_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_SV_ECarObj_ComboBox.Name = "MMIC_SV_ECarObj_ComboBox"
        Me.MMIC_SV_ECarObj_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.MMIC_SV_ECarObj_ComboBox.TabIndex = 51
        Me.MMIC_SV_ECarObj_ComboBox.Text = "預設值可選擇"
        '
        'Label106
        '
        Me.Label106.AutoSize = True
        Me.Label106.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label106.Location = New System.Drawing.Point(130, 77)
        Me.Label106.Name = "Label106"
        Me.Label106.Size = New System.Drawing.Size(98, 16)
        Me.Label106.TabIndex = 4
        Me.Label106.Text = "Data File Name."
        '
        'MMIC_SV_E_Panel
        '
        Me.MMIC_SV_E_Panel.AutoScroll = True
        Me.MMIC_SV_E_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MMIC_SV_E_Panel.Location = New System.Drawing.Point(5, 96)
        Me.MMIC_SV_E_Panel.Name = "MMIC_SV_E_Panel"
        Me.MMIC_SV_E_Panel.Size = New System.Drawing.Size(300, 160)
        Me.MMIC_SV_E_Panel.TabIndex = 15
        '
        'Label107
        '
        Me.Label107.AutoSize = True
        Me.Label107.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label107.Location = New System.Drawing.Point(26, 77)
        Me.Label107.Name = "Label107"
        Me.Label107.Size = New System.Drawing.Size(51, 16)
        Me.Label107.TabIndex = 2
        Me.Label107.Text = "Car No."
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label63.Location = New System.Drawing.Point(9, 23)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(37, 16)
        Me.Label63.TabIndex = 23
        Me.Label63.Text = "BASE"
        '
        'MMIC_SV_EBase_ComboBox
        '
        Me.MMIC_SV_EBase_ComboBox.FormattingEnabled = True
        Me.MMIC_SV_EBase_ComboBox.Location = New System.Drawing.Point(52, 19)
        Me.MMIC_SV_EBase_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_SV_EBase_ComboBox.Name = "MMIC_SV_EBase_ComboBox"
        Me.MMIC_SV_EBase_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.MMIC_SV_EBase_ComboBox.TabIndex = 45
        '
        'MMIC_SV_GroupBox
        '
        Me.MMIC_SV_GroupBox.Controls.Add(Me.Label231)
        Me.MMIC_SV_GroupBox.Controls.Add(Me.MMIC_SV_NumericUpDown)
        Me.MMIC_SV_GroupBox.Controls.Add(Me.Label130)
        Me.MMIC_SV_GroupBox.Controls.Add(Me.MMIC_SV_Type_ComboBox)
        Me.MMIC_SV_GroupBox.Controls.Add(Me.MMIC_SV_Base_TextBox)
        Me.MMIC_SV_GroupBox.Controls.Add(Me.Label129)
        Me.MMIC_SV_GroupBox.Controls.Add(Me.Label103)
        Me.MMIC_SV_GroupBox.Controls.Add(Me.MMIC_SV_Panel)
        Me.MMIC_SV_GroupBox.Controls.Add(Me.Label104)
        Me.MMIC_SV_GroupBox.Location = New System.Drawing.Point(6, 281)
        Me.MMIC_SV_GroupBox.Name = "MMIC_SV_GroupBox"
        Me.MMIC_SV_GroupBox.Size = New System.Drawing.Size(310, 265)
        Me.MMIC_SV_GroupBox.TabIndex = 52
        Me.MMIC_SV_GroupBox.TabStop = False
        Me.MMIC_SV_GroupBox.Text = "MAIN COMPUTER(CP40x)."
        '
        'Label231
        '
        Me.Label231.AutoSize = True
        Me.Label231.BackColor = System.Drawing.Color.Transparent
        Me.Label231.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label231.Location = New System.Drawing.Point(185, 77)
        Me.Label231.Name = "Label231"
        Me.Label231.Size = New System.Drawing.Size(37, 16)
        Me.Label231.TabIndex = 59
        Me.Label231.Text = "Base."
        Me.Label231.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MMIC_SV_NumericUpDown
        '
        Me.MMIC_SV_NumericUpDown.Location = New System.Drawing.Point(253, 16)
        Me.MMIC_SV_NumericUpDown.Name = "MMIC_SV_NumericUpDown"
        Me.MMIC_SV_NumericUpDown.ReadOnly = True
        Me.MMIC_SV_NumericUpDown.Size = New System.Drawing.Size(52, 23)
        Me.MMIC_SV_NumericUpDown.TabIndex = 58
        '
        'Label130
        '
        Me.Label130.AutoSize = True
        Me.Label130.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label130.Location = New System.Drawing.Point(6, 52)
        Me.Label130.Name = "Label130"
        Me.Label130.Size = New System.Drawing.Size(37, 16)
        Me.Label130.TabIndex = 57
        Me.Label130.Text = "BASE"
        '
        'MMIC_SV_Type_ComboBox
        '
        Me.MMIC_SV_Type_ComboBox.FormattingEnabled = True
        Me.MMIC_SV_Type_ComboBox.Location = New System.Drawing.Point(45, 20)
        Me.MMIC_SV_Type_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_SV_Type_ComboBox.Name = "MMIC_SV_Type_ComboBox"
        Me.MMIC_SV_Type_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.MMIC_SV_Type_ComboBox.TabIndex = 56
        '
        'MMIC_SV_Base_TextBox
        '
        Me.MMIC_SV_Base_TextBox.Location = New System.Drawing.Point(45, 49)
        Me.MMIC_SV_Base_TextBox.Name = "MMIC_SV_Base_TextBox"
        Me.MMIC_SV_Base_TextBox.Size = New System.Drawing.Size(140, 23)
        Me.MMIC_SV_Base_TextBox.TabIndex = 55
        '
        'Label129
        '
        Me.Label129.AutoSize = True
        Me.Label129.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label129.Location = New System.Drawing.Point(6, 24)
        Me.Label129.Name = "Label129"
        Me.Label129.Size = New System.Drawing.Size(36, 16)
        Me.Label129.TabIndex = 54
        Me.Label129.Text = "Type"
        '
        'Label103
        '
        Me.Label103.AutoSize = True
        Me.Label103.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label103.Location = New System.Drawing.Point(80, 77)
        Me.Label103.Name = "Label103"
        Me.Label103.Size = New System.Drawing.Size(87, 16)
        Me.Label103.TabIndex = 4
        Me.Label103.Text = "Object Name."
        '
        'MMIC_SV_Panel
        '
        Me.MMIC_SV_Panel.AutoScroll = True
        Me.MMIC_SV_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MMIC_SV_Panel.Location = New System.Drawing.Point(5, 96)
        Me.MMIC_SV_Panel.Name = "MMIC_SV_Panel"
        Me.MMIC_SV_Panel.Size = New System.Drawing.Size(300, 160)
        Me.MMIC_SV_Panel.TabIndex = 9
        '
        'Label104
        '
        Me.Label104.AutoSize = True
        Me.Label104.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label104.Location = New System.Drawing.Point(23, 77)
        Me.Label104.Name = "Label104"
        Me.Label104.Size = New System.Drawing.Size(51, 16)
        Me.Label104.TabIndex = 2
        Me.Label104.Text = "Car No."
        '
        'MMIC_MR_E_GroupBox
        '
        Me.MMIC_MR_E_GroupBox.Controls.Add(Me.MMIC_MR_E_NumericUpDown)
        Me.MMIC_MR_E_GroupBox.Controls.Add(Me.MMIC_MR_ECarObj_ComboBox)
        Me.MMIC_MR_E_GroupBox.Controls.Add(Me.Label100)
        Me.MMIC_MR_E_GroupBox.Controls.Add(Me.MMIC_MR_E_Panel)
        Me.MMIC_MR_E_GroupBox.Controls.Add(Me.Label101)
        Me.MMIC_MR_E_GroupBox.Controls.Add(Me.Label62)
        Me.MMIC_MR_E_GroupBox.Controls.Add(Me.MMIC_MR_EBase_ComboBox)
        Me.MMIC_MR_E_GroupBox.Location = New System.Drawing.Point(320, 10)
        Me.MMIC_MR_E_GroupBox.Name = "MMIC_MR_E_GroupBox"
        Me.MMIC_MR_E_GroupBox.Size = New System.Drawing.Size(310, 265)
        Me.MMIC_MR_E_GroupBox.TabIndex = 51
        Me.MMIC_MR_E_GroupBox.TabStop = False
        Me.MMIC_MR_E_GroupBox.Text = "MR-MIC EEPROM DATA."
        '
        'MMIC_MR_E_NumericUpDown
        '
        Me.MMIC_MR_E_NumericUpDown.Location = New System.Drawing.Point(252, 16)
        Me.MMIC_MR_E_NumericUpDown.Name = "MMIC_MR_E_NumericUpDown"
        Me.MMIC_MR_E_NumericUpDown.ReadOnly = True
        Me.MMIC_MR_E_NumericUpDown.Size = New System.Drawing.Size(52, 23)
        Me.MMIC_MR_E_NumericUpDown.TabIndex = 55
        '
        'MMIC_MR_ECarObj_ComboBox
        '
        Me.MMIC_MR_ECarObj_ComboBox.FormattingEnabled = True
        Me.MMIC_MR_ECarObj_ComboBox.Location = New System.Drawing.Point(52, 48)
        Me.MMIC_MR_ECarObj_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_MR_ECarObj_ComboBox.Name = "MMIC_MR_ECarObj_ComboBox"
        Me.MMIC_MR_ECarObj_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.MMIC_MR_ECarObj_ComboBox.TabIndex = 43
        Me.MMIC_MR_ECarObj_ComboBox.Text = "預設值可選擇"
        '
        'Label100
        '
        Me.Label100.AutoSize = True
        Me.Label100.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label100.Location = New System.Drawing.Point(130, 77)
        Me.Label100.Name = "Label100"
        Me.Label100.Size = New System.Drawing.Size(98, 16)
        Me.Label100.TabIndex = 4
        Me.Label100.Text = "Data File Name."
        '
        'MMIC_MR_E_Panel
        '
        Me.MMIC_MR_E_Panel.AutoScroll = True
        Me.MMIC_MR_E_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MMIC_MR_E_Panel.Location = New System.Drawing.Point(5, 96)
        Me.MMIC_MR_E_Panel.Name = "MMIC_MR_E_Panel"
        Me.MMIC_MR_E_Panel.Size = New System.Drawing.Size(300, 160)
        Me.MMIC_MR_E_Panel.TabIndex = 7
        '
        'Label101
        '
        Me.Label101.AutoSize = True
        Me.Label101.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label101.Location = New System.Drawing.Point(26, 77)
        Me.Label101.Name = "Label101"
        Me.Label101.Size = New System.Drawing.Size(51, 16)
        Me.Label101.TabIndex = 2
        Me.Label101.Text = "Car No."
        '
        'Label62
        '
        Me.Label62.AutoSize = True
        Me.Label62.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label62.Location = New System.Drawing.Point(9, 23)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(37, 16)
        Me.Label62.TabIndex = 22
        Me.Label62.Text = "BASE"
        '
        'MMIC_MR_EBase_ComboBox
        '
        Me.MMIC_MR_EBase_ComboBox.FormattingEnabled = True
        Me.MMIC_MR_EBase_ComboBox.Location = New System.Drawing.Point(52, 19)
        Me.MMIC_MR_EBase_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_MR_EBase_ComboBox.Name = "MMIC_MR_EBase_ComboBox"
        Me.MMIC_MR_EBase_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.MMIC_MR_EBase_ComboBox.TabIndex = 44
        '
        'MMIC_MR_GroupBox
        '
        Me.MMIC_MR_GroupBox.Controls.Add(Me.Label229)
        Me.MMIC_MR_GroupBox.Controls.Add(Me.MMIC_MR_NumericUpDown)
        Me.MMIC_MR_GroupBox.Controls.Add(Me.MMIC_MR_Base_TextBox)
        Me.MMIC_MR_GroupBox.Controls.Add(Me.Label64)
        Me.MMIC_MR_GroupBox.Controls.Add(Me.Label99)
        Me.MMIC_MR_GroupBox.Controls.Add(Me.Label95)
        Me.MMIC_MR_GroupBox.Controls.Add(Me.MMIC_MR_CP43x_ComboBox)
        Me.MMIC_MR_GroupBox.Controls.Add(Me.MMIC_MR_Panel)
        Me.MMIC_MR_GroupBox.Controls.Add(Me.Label128)
        Me.MMIC_MR_GroupBox.Location = New System.Drawing.Point(6, 10)
        Me.MMIC_MR_GroupBox.Name = "MMIC_MR_GroupBox"
        Me.MMIC_MR_GroupBox.Size = New System.Drawing.Size(310, 265)
        Me.MMIC_MR_GroupBox.TabIndex = 50
        Me.MMIC_MR_GroupBox.TabStop = False
        Me.MMIC_MR_GroupBox.Text = "MR-MIC(CP41x)"
        '
        'Label229
        '
        Me.Label229.AutoSize = True
        Me.Label229.BackColor = System.Drawing.Color.Transparent
        Me.Label229.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label229.Location = New System.Drawing.Point(185, 76)
        Me.Label229.Name = "Label229"
        Me.Label229.Size = New System.Drawing.Size(37, 16)
        Me.Label229.TabIndex = 55
        Me.Label229.Text = "Base."
        Me.Label229.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'MMIC_MR_NumericUpDown
        '
        Me.MMIC_MR_NumericUpDown.Location = New System.Drawing.Point(252, 16)
        Me.MMIC_MR_NumericUpDown.Name = "MMIC_MR_NumericUpDown"
        Me.MMIC_MR_NumericUpDown.ReadOnly = True
        Me.MMIC_MR_NumericUpDown.Size = New System.Drawing.Size(52, 23)
        Me.MMIC_MR_NumericUpDown.TabIndex = 54
        '
        'MMIC_MR_Base_TextBox
        '
        Me.MMIC_MR_Base_TextBox.Location = New System.Drawing.Point(45, 20)
        Me.MMIC_MR_Base_TextBox.Name = "MMIC_MR_Base_TextBox"
        Me.MMIC_MR_Base_TextBox.Size = New System.Drawing.Size(140, 23)
        Me.MMIC_MR_Base_TextBox.TabIndex = 53
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label64.Location = New System.Drawing.Point(6, 23)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(37, 16)
        Me.Label64.TabIndex = 52
        Me.Label64.Text = "BASE"
        '
        'Label99
        '
        Me.Label99.AutoSize = True
        Me.Label99.BackColor = System.Drawing.Color.Transparent
        Me.Label99.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label99.Location = New System.Drawing.Point(80, 76)
        Me.Label99.Name = "Label99"
        Me.Label99.Size = New System.Drawing.Size(87, 16)
        Me.Label99.TabIndex = 4
        Me.Label99.Text = "Object Name."
        Me.Label99.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label95
        '
        Me.Label95.AutoSize = True
        Me.Label95.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label95.Location = New System.Drawing.Point(23, 76)
        Me.Label95.Name = "Label95"
        Me.Label95.Size = New System.Drawing.Size(51, 16)
        Me.Label95.TabIndex = 2
        Me.Label95.Text = "Car No."
        '
        'MMIC_MR_CP43x_ComboBox
        '
        Me.MMIC_MR_CP43x_ComboBox.FormattingEnabled = True
        Me.MMIC_MR_CP43x_ComboBox.Location = New System.Drawing.Point(45, 48)
        Me.MMIC_MR_CP43x_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_MR_CP43x_ComboBox.Name = "MMIC_MR_CP43x_ComboBox"
        Me.MMIC_MR_CP43x_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.MMIC_MR_CP43x_ComboBox.TabIndex = 49
        '
        'MMIC_MR_Panel
        '
        Me.MMIC_MR_Panel.AutoScroll = True
        Me.MMIC_MR_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MMIC_MR_Panel.Location = New System.Drawing.Point(5, 96)
        Me.MMIC_MR_Panel.Name = "MMIC_MR_Panel"
        Me.MMIC_MR_Panel.Size = New System.Drawing.Size(300, 160)
        Me.MMIC_MR_Panel.TabIndex = 5
        '
        'Label128
        '
        Me.Label128.AutoSize = True
        Me.Label128.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label128.Location = New System.Drawing.Point(191, 52)
        Me.Label128.Name = "Label128"
        Me.Label128.Size = New System.Drawing.Size(43, 16)
        Me.Label128.TabIndex = 48
        Me.Label128.Text = "CP43x"
        '
        'MMIC_GroupBox
        '
        Me.MMIC_GroupBox.Controls.Add(Me.Label111)
        Me.MMIC_GroupBox.Controls.Add(Me.Label112)
        Me.MMIC_GroupBox.Controls.Add(Me.MMIC_FLEX_N_ComboBox)
        Me.MMIC_GroupBox.Controls.Add(Me.MMIC_MachineType_ComboBox)
        Me.MMIC_GroupBox.Enabled = False
        Me.MMIC_GroupBox.Location = New System.Drawing.Point(11, 12)
        Me.MMIC_GroupBox.Name = "MMIC_GroupBox"
        Me.MMIC_GroupBox.Size = New System.Drawing.Size(629, 45)
        Me.MMIC_GroupBox.TabIndex = 45
        Me.MMIC_GroupBox.TabStop = False
        '
        'Label111
        '
        Me.Label111.AutoSize = True
        Me.Label111.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label111.Location = New System.Drawing.Point(6, 19)
        Me.Label111.Name = "Label111"
        Me.Label111.Size = New System.Drawing.Size(35, 16)
        Me.Label111.TabIndex = 41
        Me.Label111.Text = "機種."
        '
        'Label112
        '
        Me.Label112.AutoSize = True
        Me.Label112.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label112.Location = New System.Drawing.Point(304, 19)
        Me.Label112.Name = "Label112"
        Me.Label112.Size = New System.Drawing.Size(74, 16)
        Me.Label112.TabIndex = 39
        Me.Label112.Text = "FLEX-N幾百"
        '
        'MMIC_FLEX_N_ComboBox
        '
        Me.MMIC_FLEX_N_ComboBox.FormattingEnabled = True
        Me.MMIC_FLEX_N_ComboBox.Location = New System.Drawing.Point(385, 15)
        Me.MMIC_FLEX_N_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_FLEX_N_ComboBox.Name = "MMIC_FLEX_N_ComboBox"
        Me.MMIC_FLEX_N_ComboBox.Size = New System.Drawing.Size(170, 24)
        Me.MMIC_FLEX_N_ComboBox.TabIndex = 40
        '
        'MMIC_MachineType_ComboBox
        '
        Me.MMIC_MachineType_ComboBox.FormattingEnabled = True
        Me.MMIC_MachineType_ComboBox.Location = New System.Drawing.Point(47, 15)
        Me.MMIC_MachineType_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MMIC_MachineType_ComboBox.Name = "MMIC_MachineType_ComboBox"
        Me.MMIC_MachineType_ComboBox.Size = New System.Drawing.Size(250, 24)
        Me.MMIC_MachineType_ComboBox.TabIndex = 42
        '
        'Use_mmic_CheckBox
        '
        Me.Use_mmic_CheckBox.AutoSize = True
        Me.Use_mmic_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_mmic_CheckBox.Name = "Use_mmic_CheckBox"
        Me.Use_mmic_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Use_mmic_CheckBox.TabIndex = 44
        Me.Use_mmic_CheckBox.UseVisualStyleBackColor = True
        '
        'Important_TabPage
        '
        Me.Important_TabPage.Controls.Add(Me.ImpSetting_GroupBox)
        Me.Important_TabPage.Controls.Add(Me.Use_Imp_CheckBox)
        Me.Important_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.Important_TabPage.Name = "Important_TabPage"
        Me.Important_TabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.Important_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.Important_TabPage.TabIndex = 7
        Me.Important_TabPage.Text = "重要設定"
        Me.Important_TabPage.UseVisualStyleBackColor = True
        '
        'ImpSetting_GroupBox
        '
        Me.ImpSetting_GroupBox.Controls.Add(Me.HIN_TestButton)
        Me.ImpSetting_GroupBox.Controls.Add(Me.HallIndicator_FlowLayoutPanel)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Label93)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Imp_MachineRoom_ComboBox)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Imp_DoorType_TextBox)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Label61)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Label127)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Label94)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Label97)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Label96)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Imp_OverBalance_ComboBox)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Imp_WHB_ComboBox)
        Me.ImpSetting_GroupBox.Controls.Add(Me.Imp_FAN_ComboBox)
        Me.ImpSetting_GroupBox.Enabled = False
        Me.ImpSetting_GroupBox.Location = New System.Drawing.Point(6, 20)
        Me.ImpSetting_GroupBox.Name = "ImpSetting_GroupBox"
        Me.ImpSetting_GroupBox.Size = New System.Drawing.Size(644, 558)
        Me.ImpSetting_GroupBox.TabIndex = 18
        Me.ImpSetting_GroupBox.TabStop = False
        '
        'HIN_TestButton
        '
        Me.HIN_TestButton.Location = New System.Drawing.Point(143, 132)
        Me.HIN_TestButton.Name = "HIN_TestButton"
        Me.HIN_TestButton.Size = New System.Drawing.Size(75, 23)
        Me.HIN_TestButton.TabIndex = 20
        Me.HIN_TestButton.Text = "產出測試"
        Me.HIN_TestButton.UseVisualStyleBackColor = True
        '
        'HallIndicator_FlowLayoutPanel
        '
        Me.HallIndicator_FlowLayoutPanel.AutoScroll = True
        Me.HallIndicator_FlowLayoutPanel.BackColor = System.Drawing.Color.Cornsilk
        Me.HallIndicator_FlowLayoutPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.HallIndicator_FlowLayoutPanel.Location = New System.Drawing.Point(13, 161)
        Me.HallIndicator_FlowLayoutPanel.Margin = New System.Windows.Forms.Padding(10)
        Me.HallIndicator_FlowLayoutPanel.Name = "HallIndicator_FlowLayoutPanel"
        Me.HallIndicator_FlowLayoutPanel.Size = New System.Drawing.Size(618, 384)
        Me.HallIndicator_FlowLayoutPanel.TabIndex = 18
        '
        'Label93
        '
        Me.Label93.AutoSize = True
        Me.Label93.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label93.Location = New System.Drawing.Point(10, 92)
        Me.Label93.Name = "Label93"
        Me.Label93.Size = New System.Drawing.Size(47, 16)
        Me.Label93.TabIndex = 4
        Me.Label93.Text = "機械室."
        Me.Label93.Visible = False
        '
        'Imp_MachineRoom_ComboBox
        '
        Me.Imp_MachineRoom_ComboBox.FormattingEnabled = True
        Me.Imp_MachineRoom_ComboBox.Items.AddRange(New Object() {"WITH", "WITHOUT"})
        Me.Imp_MachineRoom_ComboBox.Location = New System.Drawing.Point(120, 88)
        Me.Imp_MachineRoom_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Imp_MachineRoom_ComboBox.Name = "Imp_MachineRoom_ComboBox"
        Me.Imp_MachineRoom_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.Imp_MachineRoom_ComboBox.TabIndex = 7
        Me.Imp_MachineRoom_ComboBox.Visible = False
        '
        'Imp_DoorType_TextBox
        '
        Me.Imp_DoorType_TextBox.Location = New System.Drawing.Point(445, 54)
        Me.Imp_DoorType_TextBox.Name = "Imp_DoorType_TextBox"
        Me.Imp_DoorType_TextBox.Size = New System.Drawing.Size(140, 23)
        Me.Imp_DoorType_TextBox.TabIndex = 17
        '
        'Label61
        '
        Me.Label61.AutoSize = True
        Me.Label61.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label61.Location = New System.Drawing.Point(275, 19)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(152, 16)
        Me.Label61.TabIndex = 13
        Me.Label61.Text = "WHEEL CHAIR MAIN COB"
        '
        'Label127
        '
        Me.Label127.AutoSize = True
        Me.Label127.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label127.Location = New System.Drawing.Point(275, 57)
        Me.Label127.Name = "Label127"
        Me.Label127.Size = New System.Drawing.Size(157, 16)
        Me.Label127.TabIndex = 16
        Me.Label127.Text = "Door Type(DRD-17需設定)"
        '
        'Label94
        '
        Me.Label94.AutoSize = True
        Me.Label94.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label94.Location = New System.Drawing.Point(275, 91)
        Me.Label94.Name = "Label94"
        Me.Label94.Size = New System.Drawing.Size(59, 16)
        Me.Label94.TabIndex = 8
        Me.Label94.Text = "風扇連動."
        Me.Label94.Visible = False
        '
        'Label97
        '
        Me.Label97.AutoSize = True
        Me.Label97.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label97.Location = New System.Drawing.Point(10, 19)
        Me.Label97.Name = "Label97"
        Me.Label97.Size = New System.Drawing.Size(104, 16)
        Me.Label97.TabIndex = 10
        Me.Label97.Text = "Over Balance(%)."
        '
        'Label96
        '
        Me.Label96.AutoSize = True
        Me.Label96.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label96.Location = New System.Drawing.Point(10, 135)
        Me.Label96.Name = "Label96"
        Me.Label96.Size = New System.Drawing.Size(127, 16)
        Me.Label96.TabIndex = 11
        Me.Label96.Text = "Hall Indicator(制御階)"
        '
        'Imp_OverBalance_ComboBox
        '
        Me.Imp_OverBalance_ComboBox.FormattingEnabled = True
        Me.Imp_OverBalance_ComboBox.Items.AddRange(New Object() {"45%", "50%"})
        Me.Imp_OverBalance_ComboBox.Location = New System.Drawing.Point(120, 15)
        Me.Imp_OverBalance_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Imp_OverBalance_ComboBox.Name = "Imp_OverBalance_ComboBox"
        Me.Imp_OverBalance_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.Imp_OverBalance_ComboBox.TabIndex = 15
        '
        'Imp_WHB_ComboBox
        '
        Me.Imp_WHB_ComboBox.FormattingEnabled = True
        Me.Imp_WHB_ComboBox.Items.AddRange(New Object() {"WITH", "WITH(Y17ABWC)", "WITHOUT"})
        Me.Imp_WHB_ComboBox.Location = New System.Drawing.Point(445, 15)
        Me.Imp_WHB_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Imp_WHB_ComboBox.Name = "Imp_WHB_ComboBox"
        Me.Imp_WHB_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.Imp_WHB_ComboBox.TabIndex = 14
        '
        'Imp_FAN_ComboBox
        '
        Me.Imp_FAN_ComboBox.FormattingEnabled = True
        Me.Imp_FAN_ComboBox.Items.AddRange(New Object() {"WITH", "WITH(ION)", "WITHOUT"})
        Me.Imp_FAN_ComboBox.Location = New System.Drawing.Point(445, 87)
        Me.Imp_FAN_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Imp_FAN_ComboBox.Name = "Imp_FAN_ComboBox"
        Me.Imp_FAN_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.Imp_FAN_ComboBox.TabIndex = 9
        Me.Imp_FAN_ComboBox.Visible = False
        '
        'Use_Imp_CheckBox
        '
        Me.Use_Imp_CheckBox.AutoSize = True
        Me.Use_Imp_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_Imp_CheckBox.Name = "Use_Imp_CheckBox"
        Me.Use_Imp_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Use_Imp_CheckBox.TabIndex = 16
        Me.Use_Imp_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec
        '
        Me.Spec.Controls.Add(Me.Spec_TabControl)
        Me.Spec.Location = New System.Drawing.Point(4, 25)
        Me.Spec.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec.Name = "Spec"
        Me.Spec.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec.Size = New System.Drawing.Size(664, 584)
        Me.Spec.TabIndex = 1
        Me.Spec.Text = "仕樣"
        Me.Spec.UseVisualStyleBackColor = True
        '
        'Spec_TabControl
        '
        Me.Spec_TabControl.Controls.Add(Me.Spec_BasicAll_TabPage)
        Me.Spec_TabControl.Controls.Add(Me.Spec_TW_TabPage)
        Me.Spec_TabControl.Location = New System.Drawing.Point(7, 5)
        Me.Spec_TabControl.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_TabControl.Name = "Spec_TabControl"
        Me.Spec_TabControl.SelectedIndex = 0
        Me.Spec_TabControl.Size = New System.Drawing.Size(652, 575)
        Me.Spec_TabControl.TabIndex = 12
        '
        'Spec_BasicAll_TabPage
        '
        Me.Spec_BasicAll_TabPage.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.Spec_BasicAll_TabPage.Controls.Add(Me.Spec_BasicAll_TabControl)
        Me.Spec_BasicAll_TabPage.Controls.Add(Me.Use_SpecBasic_CheckBox)
        Me.Spec_BasicAll_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.Spec_BasicAll_TabPage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_BasicAll_TabPage.Name = "Spec_BasicAll_TabPage"
        Me.Spec_BasicAll_TabPage.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_BasicAll_TabPage.Size = New System.Drawing.Size(644, 546)
        Me.Spec_BasicAll_TabPage.TabIndex = 0
        Me.Spec_BasicAll_TabPage.Text = "Basic_all"
        Me.Spec_BasicAll_TabPage.UseVisualStyleBackColor = True
        '
        'Spec_BasicAll_TabControl
        '
        Me.Spec_BasicAll_TabControl.Controls.Add(Me.TabPage7)
        Me.Spec_BasicAll_TabControl.Controls.Add(Me.TabPage8)
        Me.Spec_BasicAll_TabControl.Location = New System.Drawing.Point(7, 19)
        Me.Spec_BasicAll_TabControl.Name = "Spec_BasicAll_TabControl"
        Me.Spec_BasicAll_TabControl.SelectedIndex = 0
        Me.Spec_BasicAll_TabControl.Size = New System.Drawing.Size(628, 520)
        Me.Spec_BasicAll_TabControl.TabIndex = 68
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.SpecBasic_GroupBox)
        Me.TabPage7.Location = New System.Drawing.Point(4, 25)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage7.Size = New System.Drawing.Size(620, 491)
        Me.TabPage7.TabIndex = 0
        Me.TabPage7.Text = "Page1"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'SpecBasic_GroupBox
        '
        Me.SpecBasic_GroupBox.Controls.Add(Me.Spec_LiftNum_NumericUpDown)
        Me.SpecBasic_GroupBox.Controls.Add(Me.Label8)
        Me.SpecBasic_GroupBox.Controls.Add(Me.SpecBasic_LiftItem_Dynamic_Panel)
        Me.SpecBasic_GroupBox.Controls.Add(Me.SpecBasic_LiftItem_Panel)
        Me.SpecBasic_GroupBox.Enabled = False
        Me.SpecBasic_GroupBox.Location = New System.Drawing.Point(3, 3)
        Me.SpecBasic_GroupBox.Name = "SpecBasic_GroupBox"
        Me.SpecBasic_GroupBox.Size = New System.Drawing.Size(614, 482)
        Me.SpecBasic_GroupBox.TabIndex = 24
        Me.SpecBasic_GroupBox.TabStop = False
        '
        'Spec_LiftNum_NumericUpDown
        '
        Me.Spec_LiftNum_NumericUpDown.Location = New System.Drawing.Point(90, 13)
        Me.Spec_LiftNum_NumericUpDown.Name = "Spec_LiftNum_NumericUpDown"
        Me.Spec_LiftNum_NumericUpDown.ReadOnly = True
        Me.Spec_LiftNum_NumericUpDown.Size = New System.Drawing.Size(54, 23)
        Me.Spec_LiftNum_NumericUpDown.TabIndex = 23
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label8.Location = New System.Drawing.Point(9, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(62, 16)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "電梯總數 :"
        '
        'SpecBasic_LiftItem_Dynamic_Panel
        '
        Me.SpecBasic_LiftItem_Dynamic_Panel.AutoScroll = True
        Me.SpecBasic_LiftItem_Dynamic_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SpecBasic_LiftItem_Dynamic_Panel.Location = New System.Drawing.Point(5, 117)
        Me.SpecBasic_LiftItem_Dynamic_Panel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.SpecBasic_LiftItem_Dynamic_Panel.Name = "SpecBasic_LiftItem_Dynamic_Panel"
        Me.SpecBasic_LiftItem_Dynamic_Panel.Size = New System.Drawing.Size(606, 358)
        Me.SpecBasic_LiftItem_Dynamic_Panel.TabIndex = 0
        '
        'SpecBasic_LiftItem_Panel
        '
        Me.SpecBasic_LiftItem_Panel.AutoScroll = True
        Me.SpecBasic_LiftItem_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_BtmFL_Real_TextBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_TopFL_Real_TextBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_Control_ComboBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_FLName_TextBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_Speed_TextBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_StopFL_TextBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_LiftName_TextBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_BtmFL_TextBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_LiftMem_TextBox)
        Me.SpecBasic_LiftItem_Panel.Controls.Add(Me.Spec_TopFL_TextBox)
        Me.SpecBasic_LiftItem_Panel.Location = New System.Drawing.Point(5, 43)
        Me.SpecBasic_LiftItem_Panel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.SpecBasic_LiftItem_Panel.Name = "SpecBasic_LiftItem_Panel"
        Me.SpecBasic_LiftItem_Panel.Size = New System.Drawing.Size(606, 66)
        Me.SpecBasic_LiftItem_Panel.TabIndex = 13
        '
        'Spec_BtmFL_Real_TextBox
        '
        Me.Spec_BtmFL_Real_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_BtmFL_Real_TextBox.Location = New System.Drawing.Point(447, 11)
        Me.Spec_BtmFL_Real_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_BtmFL_Real_TextBox.MaxLength = 2
        Me.Spec_BtmFL_Real_TextBox.Name = "Spec_BtmFL_Real_TextBox"
        Me.Spec_BtmFL_Real_TextBox.ReadOnly = True
        Me.Spec_BtmFL_Real_TextBox.Size = New System.Drawing.Size(35, 23)
        Me.Spec_BtmFL_Real_TextBox.TabIndex = 23
        Me.Spec_BtmFL_Real_TextBox.Text = "FL"
        '
        'Spec_TopFL_Real_TextBox
        '
        Me.Spec_TopFL_Real_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_TopFL_Real_TextBox.Location = New System.Drawing.Point(348, 11)
        Me.Spec_TopFL_Real_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_TopFL_Real_TextBox.MaxLength = 2
        Me.Spec_TopFL_Real_TextBox.Name = "Spec_TopFL_Real_TextBox"
        Me.Spec_TopFL_Real_TextBox.ReadOnly = True
        Me.Spec_TopFL_Real_TextBox.Size = New System.Drawing.Size(35, 23)
        Me.Spec_TopFL_Real_TextBox.TabIndex = 22
        Me.Spec_TopFL_Real_TextBox.Text = "FL"
        '
        'Spec_Control_ComboBox
        '
        Me.Spec_Control_ComboBox.Enabled = False
        Me.Spec_Control_ComboBox.FormattingEnabled = True
        Me.Spec_Control_ComboBox.Location = New System.Drawing.Point(208, 10)
        Me.Spec_Control_ComboBox.Name = "Spec_Control_ComboBox"
        Me.Spec_Control_ComboBox.Size = New System.Drawing.Size(76, 24)
        Me.Spec_Control_ComboBox.TabIndex = 0
        Me.Spec_Control_ComboBox.Text = "操作方式"
        '
        'Spec_FLName_TextBox
        '
        Me.Spec_FLName_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_FLName_TextBox.Location = New System.Drawing.Point(698, 11)
        Me.Spec_FLName_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_FLName_TextBox.MaxLength = 2
        Me.Spec_FLName_TextBox.Name = "Spec_FLName_TextBox"
        Me.Spec_FLName_TextBox.ReadOnly = True
        Me.Spec_FLName_TextBox.Size = New System.Drawing.Size(116, 23)
        Me.Spec_FLName_TextBox.TabIndex = 21
        Me.Spec_FLName_TextBox.Text = "表示階名"
        '
        'Spec_Speed_TextBox
        '
        Me.Spec_Speed_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Speed_TextBox.Location = New System.Drawing.Point(604, 11)
        Me.Spec_Speed_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_Speed_TextBox.MaxLength = 2
        Me.Spec_Speed_TextBox.Name = "Spec_Speed_TextBox"
        Me.Spec_Speed_TextBox.ReadOnly = True
        Me.Spec_Speed_TextBox.Size = New System.Drawing.Size(69, 23)
        Me.Spec_Speed_TextBox.TabIndex = 17
        Me.Spec_Speed_TextBox.Text = "速度"
        '
        'Spec_StopFL_TextBox
        '
        Me.Spec_StopFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_StopFL_TextBox.Location = New System.Drawing.Point(505, 11)
        Me.Spec_StopFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_StopFL_TextBox.MaxLength = 2
        Me.Spec_StopFL_TextBox.Name = "Spec_StopFL_TextBox"
        Me.Spec_StopFL_TextBox.ReadOnly = True
        Me.Spec_StopFL_TextBox.Size = New System.Drawing.Size(69, 23)
        Me.Spec_StopFL_TextBox.TabIndex = 16
        Me.Spec_StopFL_TextBox.Text = "停止數"
        '
        'Spec_LiftName_TextBox
        '
        Me.Spec_LiftName_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_LiftName_TextBox.Location = New System.Drawing.Point(9, 10)
        Me.Spec_LiftName_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_LiftName_TextBox.MaxLength = 2
        Me.Spec_LiftName_TextBox.Name = "Spec_LiftName_TextBox"
        Me.Spec_LiftName_TextBox.ReadOnly = True
        Me.Spec_LiftName_TextBox.Size = New System.Drawing.Size(69, 23)
        Me.Spec_LiftName_TextBox.TabIndex = 11
        Me.Spec_LiftName_TextBox.Text = "號機名"
        '
        'Spec_BtmFL_TextBox
        '
        Me.Spec_BtmFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_BtmFL_TextBox.Location = New System.Drawing.Point(406, 11)
        Me.Spec_BtmFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_BtmFL_TextBox.MaxLength = 2
        Me.Spec_BtmFL_TextBox.Name = "Spec_BtmFL_TextBox"
        Me.Spec_BtmFL_TextBox.ReadOnly = True
        Me.Spec_BtmFL_TextBox.Size = New System.Drawing.Size(35, 23)
        Me.Spec_BtmFL_TextBox.TabIndex = 15
        Me.Spec_BtmFL_TextBox.Text = "Btm"
        '
        'Spec_LiftMem_TextBox
        '
        Me.Spec_LiftMem_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_LiftMem_TextBox.Location = New System.Drawing.Point(108, 10)
        Me.Spec_LiftMem_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_LiftMem_TextBox.MaxLength = 2
        Me.Spec_LiftMem_TextBox.Name = "Spec_LiftMem_TextBox"
        Me.Spec_LiftMem_TextBox.ReadOnly = True
        Me.Spec_LiftMem_TextBox.Size = New System.Drawing.Size(69, 23)
        Me.Spec_LiftMem_TextBox.TabIndex = 12
        Me.Spec_LiftMem_TextBox.Text = "號機"
        '
        'Spec_TopFL_TextBox
        '
        Me.Spec_TopFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_TopFL_TextBox.Location = New System.Drawing.Point(307, 11)
        Me.Spec_TopFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_TopFL_TextBox.MaxLength = 2
        Me.Spec_TopFL_TextBox.Name = "Spec_TopFL_TextBox"
        Me.Spec_TopFL_TextBox.ReadOnly = True
        Me.Spec_TopFL_TextBox.Size = New System.Drawing.Size(35, 23)
        Me.Spec_TopFL_TextBox.TabIndex = 14
        Me.Spec_TopFL_TextBox.Text = "Top"
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.SpecBasic_GroupBox2)
        Me.TabPage8.Location = New System.Drawing.Point(4, 25)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage8.Size = New System.Drawing.Size(620, 491)
        Me.TabPage8.TabIndex = 1
        Me.TabPage8.Text = "Page2"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'SpecBasic_GroupBox2
        '
        Me.SpecBasic_GroupBox2.Controls.Add(Me.Spec_MachineType_NumericUpDown)
        Me.SpecBasic_GroupBox2.Controls.Add(Me.Spec_ControlWay_Panel)
        Me.SpecBasic_GroupBox2.Controls.Add(Me.Spec_MachineType_Label)
        Me.SpecBasic_GroupBox2.Controls.Add(Me.SpecBasic_p2_base_Panel)
        Me.SpecBasic_GroupBox2.Controls.Add(Me.Label189)
        Me.SpecBasic_GroupBox2.Controls.Add(Me.Spec_Purpose_Panel)
        Me.SpecBasic_GroupBox2.Controls.Add(Me.Spec_ControlWay_Label)
        Me.SpecBasic_GroupBox2.Controls.Add(Me.Spec_MachineType_Panel)
        Me.SpecBasic_GroupBox2.Controls.Add(Me.Spec_Purpose_NumericUpDown)
        Me.SpecBasic_GroupBox2.Enabled = False
        Me.SpecBasic_GroupBox2.Location = New System.Drawing.Point(6, 6)
        Me.SpecBasic_GroupBox2.Name = "SpecBasic_GroupBox2"
        Me.SpecBasic_GroupBox2.Size = New System.Drawing.Size(605, 479)
        Me.SpecBasic_GroupBox2.TabIndex = 76
        Me.SpecBasic_GroupBox2.TabStop = False
        '
        'Spec_MachineType_NumericUpDown
        '
        Me.Spec_MachineType_NumericUpDown.Location = New System.Drawing.Point(17, 22)
        Me.Spec_MachineType_NumericUpDown.Name = "Spec_MachineType_NumericUpDown"
        Me.Spec_MachineType_NumericUpDown.ReadOnly = True
        Me.Spec_MachineType_NumericUpDown.Size = New System.Drawing.Size(47, 23)
        Me.Spec_MachineType_NumericUpDown.TabIndex = 69
        '
        'Spec_ControlWay_Panel
        '
        Me.Spec_ControlWay_Panel.AutoScroll = True
        Me.Spec_ControlWay_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_ControlWay_Panel.Location = New System.Drawing.Point(305, 72)
        Me.Spec_ControlWay_Panel.Name = "Spec_ControlWay_Panel"
        Me.Spec_ControlWay_Panel.Size = New System.Drawing.Size(295, 127)
        Me.Spec_ControlWay_Panel.TabIndex = 75
        '
        'Spec_MachineType_Label
        '
        Me.Spec_MachineType_Label.AutoSize = True
        Me.Spec_MachineType_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_MachineType_Label.Location = New System.Drawing.Point(8, 53)
        Me.Spec_MachineType_Label.Name = "Spec_MachineType_Label"
        Me.Spec_MachineType_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_MachineType_Label.TabIndex = 30
        Me.Spec_MachineType_Label.Text = "機種 :"
        '
        'SpecBasic_p2_base_Panel
        '
        Me.SpecBasic_p2_base_Panel.AutoScroll = True
        Me.SpecBasic_p2_base_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SpecBasic_p2_base_Panel.Controls.Add(Me.Spec_Base_ComboBox)
        Me.SpecBasic_p2_base_Panel.Location = New System.Drawing.Point(305, 254)
        Me.SpecBasic_p2_base_Panel.Name = "SpecBasic_p2_base_Panel"
        Me.SpecBasic_p2_base_Panel.Size = New System.Drawing.Size(295, 127)
        Me.SpecBasic_p2_base_Panel.TabIndex = 73
        Me.SpecBasic_p2_base_Panel.Visible = False
        '
        'Spec_Base_ComboBox
        '
        Me.Spec_Base_ComboBox.FormattingEnabled = True
        Me.Spec_Base_ComboBox.Location = New System.Drawing.Point(12, 12)
        Me.Spec_Base_ComboBox.Name = "Spec_Base_ComboBox"
        Me.Spec_Base_ComboBox.Size = New System.Drawing.Size(268, 24)
        Me.Spec_Base_ComboBox.TabIndex = 33
        '
        'Label189
        '
        Me.Label189.AutoSize = True
        Me.Label189.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label189.Location = New System.Drawing.Point(8, 235)
        Me.Label189.Name = "Label189"
        Me.Label189.Size = New System.Drawing.Size(38, 16)
        Me.Label189.TabIndex = 34
        Me.Label189.Text = "用途 :"
        '
        'Spec_Purpose_Panel
        '
        Me.Spec_Purpose_Panel.AutoScroll = True
        Me.Spec_Purpose_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Purpose_Panel.Location = New System.Drawing.Point(8, 254)
        Me.Spec_Purpose_Panel.Name = "Spec_Purpose_Panel"
        Me.Spec_Purpose_Panel.Size = New System.Drawing.Size(295, 127)
        Me.Spec_Purpose_Panel.TabIndex = 72
        '
        'Spec_ControlWay_Label
        '
        Me.Spec_ControlWay_Label.AutoSize = True
        Me.Spec_ControlWay_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_ControlWay_Label.Location = New System.Drawing.Point(305, 53)
        Me.Spec_ControlWay_Label.Name = "Spec_ControlWay_Label"
        Me.Spec_ControlWay_Label.Size = New System.Drawing.Size(62, 16)
        Me.Spec_ControlWay_Label.TabIndex = 33
        Me.Spec_ControlWay_Label.Text = "控制方式 :"
        '
        'Spec_MachineType_Panel
        '
        Me.Spec_MachineType_Panel.AutoScroll = True
        Me.Spec_MachineType_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_MachineType_Panel.Location = New System.Drawing.Point(8, 72)
        Me.Spec_MachineType_Panel.Name = "Spec_MachineType_Panel"
        Me.Spec_MachineType_Panel.Size = New System.Drawing.Size(295, 127)
        Me.Spec_MachineType_Panel.TabIndex = 72
        '
        'Spec_Purpose_NumericUpDown
        '
        Me.Spec_Purpose_NumericUpDown.Location = New System.Drawing.Point(17, 205)
        Me.Spec_Purpose_NumericUpDown.Name = "Spec_Purpose_NumericUpDown"
        Me.Spec_Purpose_NumericUpDown.ReadOnly = True
        Me.Spec_Purpose_NumericUpDown.Size = New System.Drawing.Size(47, 23)
        Me.Spec_Purpose_NumericUpDown.TabIndex = 70
        '
        'Use_SpecBasic_CheckBox
        '
        Me.Use_SpecBasic_CheckBox.AutoSize = True
        Me.Use_SpecBasic_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_SpecBasic_CheckBox.Name = "Use_SpecBasic_CheckBox"
        Me.Use_SpecBasic_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Use_SpecBasic_CheckBox.TabIndex = 0
        Me.Use_SpecBasic_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_TW_TabPage
        '
        Me.Spec_TW_TabPage.Controls.Add(Me.Spec_TW_TabControl)
        Me.Spec_TW_TabPage.Controls.Add(Me.Use_SpecTWFP17_CheckBox)
        Me.Spec_TW_TabPage.Controls.Add(Me.Use_SpecTWIDU_CheckBox)
        Me.Spec_TW_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.Spec_TW_TabPage.Name = "Spec_TW_TabPage"
        Me.Spec_TW_TabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.Spec_TW_TabPage.Size = New System.Drawing.Size(644, 546)
        Me.Spec_TW_TabPage.TabIndex = 2
        Me.Spec_TW_TabPage.Text = "TW台灣"
        Me.Spec_TW_TabPage.UseVisualStyleBackColor = True
        '
        'Spec_TW_TabControl
        '
        Me.Spec_TW_TabControl.Controls.Add(Me.TabPage9)
        Me.Spec_TW_TabControl.Controls.Add(Me.TabPage10)
        Me.Spec_TW_TabControl.Controls.Add(Me.TabPage12)
        Me.Spec_TW_TabControl.Controls.Add(Me.TabPage13)
        Me.Spec_TW_TabControl.Controls.Add(Me.TabPage14)
        Me.Spec_TW_TabControl.Controls.Add(Me.TabPage15)
        Me.Spec_TW_TabControl.Controls.Add(Me.TabPage11)
        Me.Spec_TW_TabControl.Location = New System.Drawing.Point(3, 42)
        Me.Spec_TW_TabControl.Name = "Spec_TW_TabControl"
        Me.Spec_TW_TabControl.SelectedIndex = 0
        Me.Spec_TW_TabControl.Size = New System.Drawing.Size(635, 498)
        Me.Spec_TW_TabControl.TabIndex = 18
        '
        'TabPage9
        '
        Me.TabPage9.Controls.Add(Me.Spec_TW_FlowLayoutPanel1)
        Me.TabPage9.Location = New System.Drawing.Point(4, 25)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage9.Size = New System.Drawing.Size(627, 469)
        Me.TabPage9.TabIndex = 0
        Me.TabPage9.Text = "Page1"
        Me.TabPage9.UseVisualStyleBackColor = True
        '
        'Spec_TW_FlowLayoutPanel1
        '
        Me.Spec_TW_FlowLayoutPanel1.AutoScroll = True
        Me.Spec_TW_FlowLayoutPanel1.Controls.Add(Me.Spec_DRAuto_Panel)
        Me.Spec_TW_FlowLayoutPanel1.Controls.Add(Me.Spec_CancellCall_Panel)
        Me.Spec_TW_FlowLayoutPanel1.Controls.Add(Me.Spec_AutoFan_Panel)
        Me.Spec_TW_FlowLayoutPanel1.Controls.Add(Me.Spec_AutoPass_Panel)
        Me.Spec_TW_FlowLayoutPanel1.Controls.Add(Me.Spec_Indep_Panel)
        Me.Spec_TW_FlowLayoutPanel1.Controls.Add(Me.Spec_HinCpi_Panel)
        Me.Spec_TW_FlowLayoutPanel1.Controls.Add(Me.Spec_Fire_Panel)
        Me.Spec_TW_FlowLayoutPanel1.Controls.Add(Me.Spec_Fireman_Panel)
        Me.Spec_TW_FlowLayoutPanel1.Enabled = False
        Me.Spec_TW_FlowLayoutPanel1.Location = New System.Drawing.Point(6, 6)
        Me.Spec_TW_FlowLayoutPanel1.Name = "Spec_TW_FlowLayoutPanel1"
        Me.Spec_TW_FlowLayoutPanel1.Size = New System.Drawing.Size(615, 457)
        Me.Spec_TW_FlowLayoutPanel1.TabIndex = 0
        '
        'Spec_DRAuto_Panel
        '
        Me.Spec_DRAuto_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_DRAuto_Panel.Controls.Add(Me.Spec_DRAuto_Label)
        Me.Spec_DRAuto_Panel.Controls.Add(Me.Spec_MechSafety_Label)
        Me.Spec_DRAuto_Panel.Controls.Add(Me.Spec_MechSafety_ComboBox)
        Me.Spec_DRAuto_Panel.Controls.Add(Me.Spec_PhotoEye_Label)
        Me.Spec_DRAuto_Panel.Controls.Add(Me.Spec_PhotoEye_ComboBox)
        Me.Spec_DRAuto_Panel.Controls.Add(Me.Spec_DRAuto_ComboBox)
        Me.Spec_DRAuto_Panel.Location = New System.Drawing.Point(3, 3)
        Me.Spec_DRAuto_Panel.Name = "Spec_DRAuto_Panel"
        Me.Spec_DRAuto_Panel.Size = New System.Drawing.Size(580, 73)
        Me.Spec_DRAuto_Panel.TabIndex = 161
        '
        'Spec_DRAuto_Label
        '
        Me.Spec_DRAuto_Label.AutoSize = True
        Me.Spec_DRAuto_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_DRAuto_Label.Location = New System.Drawing.Point(27, 10)
        Me.Spec_DRAuto_Label.Name = "Spec_DRAuto_Label"
        Me.Spec_DRAuto_Label.Size = New System.Drawing.Size(104, 16)
        Me.Spec_DRAuto_Label.TabIndex = 13
        Me.Spec_DRAuto_Label.Text = "開門時限自動調節"
        '
        'Spec_MechSafety_Label
        '
        Me.Spec_MechSafety_Label.AutoSize = True
        Me.Spec_MechSafety_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_MechSafety_Label.Location = New System.Drawing.Point(208, 47)
        Me.Spec_MechSafety_Label.Name = "Spec_MechSafety_Label"
        Me.Spec_MechSafety_Label.Size = New System.Drawing.Size(68, 16)
        Me.Spec_MechSafety_Label.TabIndex = 27
        Me.Spec_MechSafety_Label.Text = "機械式裝置"
        '
        'Spec_MechSafety_ComboBox
        '
        Me.Spec_MechSafety_ComboBox.FormattingEnabled = True
        Me.Spec_MechSafety_ComboBox.Items.AddRange(New Object() {"WITH", "WITHOUT"})
        Me.Spec_MechSafety_ComboBox.Location = New System.Drawing.Point(291, 43)
        Me.Spec_MechSafety_ComboBox.Name = "Spec_MechSafety_ComboBox"
        Me.Spec_MechSafety_ComboBox.Size = New System.Drawing.Size(76, 24)
        Me.Spec_MechSafety_ComboBox.TabIndex = 28
        '
        'Spec_PhotoEye_Label
        '
        Me.Spec_PhotoEye_Label.AutoSize = True
        Me.Spec_PhotoEye_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_PhotoEye_Label.Location = New System.Drawing.Point(208, 10)
        Me.Spec_PhotoEye_Label.Name = "Spec_PhotoEye_Label"
        Me.Spec_PhotoEye_Label.Size = New System.Drawing.Size(56, 16)
        Me.Spec_PhotoEye_Label.TabIndex = 25
        Me.Spec_PhotoEye_Label.Text = "光電裝置"
        '
        'Spec_PhotoEye_ComboBox
        '
        Me.Spec_PhotoEye_ComboBox.FormattingEnabled = True
        Me.Spec_PhotoEye_ComboBox.Items.AddRange(New Object() {"WITH", "WITHOUT"})
        Me.Spec_PhotoEye_ComboBox.Location = New System.Drawing.Point(291, 7)
        Me.Spec_PhotoEye_ComboBox.Name = "Spec_PhotoEye_ComboBox"
        Me.Spec_PhotoEye_ComboBox.Size = New System.Drawing.Size(76, 24)
        Me.Spec_PhotoEye_ComboBox.TabIndex = 24
        '
        'Spec_DRAuto_ComboBox
        '
        Me.Spec_DRAuto_ComboBox.FormattingEnabled = True
        Me.Spec_DRAuto_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_DRAuto_ComboBox.Location = New System.Drawing.Point(147, 6)
        Me.Spec_DRAuto_ComboBox.Name = "Spec_DRAuto_ComboBox"
        Me.Spec_DRAuto_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_DRAuto_ComboBox.TabIndex = 14
        '
        'Spec_CancellCall_Panel
        '
        Me.Spec_CancellCall_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_CancellCall_Panel.Controls.Add(Me.Spec_CancellCall_Label)
        Me.Spec_CancellCall_Panel.Controls.Add(Me.Spec_CancellCall_ComboBox)
        Me.Spec_CancellCall_Panel.Controls.Add(Me.Spec_SCOB_Label)
        Me.Spec_CancellCall_Panel.Controls.Add(Me.Spec_SCOB_ComboBox)
        Me.Spec_CancellCall_Panel.Location = New System.Drawing.Point(3, 82)
        Me.Spec_CancellCall_Panel.Name = "Spec_CancellCall_Panel"
        Me.Spec_CancellCall_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_CancellCall_Panel.TabIndex = 163
        '
        'Spec_CancellCall_Label
        '
        Me.Spec_CancellCall_Label.AutoSize = True
        Me.Spec_CancellCall_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CancellCall_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_CancellCall_Label.Name = "Spec_CancellCall_Label"
        Me.Spec_CancellCall_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_CancellCall_Label.TabIndex = 15
        Me.Spec_CancellCall_Label.Text = "取消嬉戲呼叫"
        '
        'Spec_CancellCall_ComboBox
        '
        Me.Spec_CancellCall_ComboBox.FormattingEnabled = True
        Me.Spec_CancellCall_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CancellCall_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_CancellCall_ComboBox.Name = "Spec_CancellCall_ComboBox"
        Me.Spec_CancellCall_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CancellCall_ComboBox.TabIndex = 29
        '
        'Spec_SCOB_Label
        '
        Me.Spec_SCOB_Label.AutoSize = True
        Me.Spec_SCOB_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_SCOB_Label.Location = New System.Drawing.Point(208, 9)
        Me.Spec_SCOB_Label.Name = "Spec_SCOB_Label"
        Me.Spec_SCOB_Label.Size = New System.Drawing.Size(45, 16)
        Me.Spec_SCOB_Label.TabIndex = 30
        Me.Spec_SCOB_Label.Text = "副COB"
        '
        'Spec_SCOB_ComboBox
        '
        Me.Spec_SCOB_ComboBox.FormattingEnabled = True
        Me.Spec_SCOB_ComboBox.Items.AddRange(New Object() {"WITH", "WITHOUT"})
        Me.Spec_SCOB_ComboBox.Location = New System.Drawing.Point(291, 5)
        Me.Spec_SCOB_ComboBox.Name = "Spec_SCOB_ComboBox"
        Me.Spec_SCOB_ComboBox.Size = New System.Drawing.Size(76, 24)
        Me.Spec_SCOB_ComboBox.TabIndex = 31
        '
        'Spec_AutoFan_Panel
        '
        Me.Spec_AutoFan_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_AutoFan_Panel.Controls.Add(Me.Spec_AutoFan_Label)
        Me.Spec_AutoFan_Panel.Controls.Add(Me.Spec_AutoFan_ComboBox)
        Me.Spec_AutoFan_Panel.Controls.Add(Me.Spec_ION_Label)
        Me.Spec_AutoFan_Panel.Controls.Add(Me.Spec_ION_ComboBox)
        Me.Spec_AutoFan_Panel.Location = New System.Drawing.Point(3, 124)
        Me.Spec_AutoFan_Panel.Name = "Spec_AutoFan_Panel"
        Me.Spec_AutoFan_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_AutoFan_Panel.TabIndex = 168
        '
        'Spec_AutoFan_Label
        '
        Me.Spec_AutoFan_Label.AutoSize = True
        Me.Spec_AutoFan_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_AutoFan_Label.Location = New System.Drawing.Point(27, 11)
        Me.Spec_AutoFan_Label.Name = "Spec_AutoFan_Label"
        Me.Spec_AutoFan_Label.Size = New System.Drawing.Size(56, 16)
        Me.Spec_AutoFan_Label.TabIndex = 20
        Me.Spec_AutoFan_Label.Text = "風扇連動"
        '
        'Spec_AutoFan_ComboBox
        '
        Me.Spec_AutoFan_ComboBox.FormattingEnabled = True
        Me.Spec_AutoFan_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_AutoFan_ComboBox.Location = New System.Drawing.Point(147, 7)
        Me.Spec_AutoFan_ComboBox.Name = "Spec_AutoFan_ComboBox"
        Me.Spec_AutoFan_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_AutoFan_ComboBox.TabIndex = 36
        '
        'Spec_ION_Label
        '
        Me.Spec_ION_Label.AutoSize = True
        Me.Spec_ION_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_ION_Label.Location = New System.Drawing.Point(208, 11)
        Me.Spec_ION_Label.Name = "Spec_ION_Label"
        Me.Spec_ION_Label.Size = New System.Drawing.Size(195, 16)
        Me.Spec_ION_Label.TabIndex = 37
        Me.Spec_ION_Label.Text = "離子除菌(TW為標準/TMB參考KEY)"
        '
        'Spec_ION_ComboBox
        '
        Me.Spec_ION_ComboBox.FormattingEnabled = True
        Me.Spec_ION_ComboBox.Items.AddRange(New Object() {"WITH", "WITHOUT"})
        Me.Spec_ION_ComboBox.Location = New System.Drawing.Point(409, 7)
        Me.Spec_ION_ComboBox.Name = "Spec_ION_ComboBox"
        Me.Spec_ION_ComboBox.Size = New System.Drawing.Size(76, 24)
        Me.Spec_ION_ComboBox.TabIndex = 38
        '
        'Spec_AutoPass_Panel
        '
        Me.Spec_AutoPass_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_AutoPass_Panel.Controls.Add(Me.Spec_AutoPass_Label)
        Me.Spec_AutoPass_Panel.Controls.Add(Me.Spec_AutoPass_ComboBox)
        Me.Spec_AutoPass_Panel.Location = New System.Drawing.Point(3, 166)
        Me.Spec_AutoPass_Panel.Name = "Spec_AutoPass_Panel"
        Me.Spec_AutoPass_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_AutoPass_Panel.TabIndex = 172
        '
        'Spec_AutoPass_Label
        '
        Me.Spec_AutoPass_Label.AutoSize = True
        Me.Spec_AutoPass_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_AutoPass_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_AutoPass_Label.Name = "Spec_AutoPass_Label"
        Me.Spec_AutoPass_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_AutoPass_Label.TabIndex = 43
        Me.Spec_AutoPass_Label.Text = "自動滿員通過"
        '
        'Spec_AutoPass_ComboBox
        '
        Me.Spec_AutoPass_ComboBox.FormattingEnabled = True
        Me.Spec_AutoPass_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_AutoPass_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_AutoPass_ComboBox.Name = "Spec_AutoPass_ComboBox"
        Me.Spec_AutoPass_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_AutoPass_ComboBox.TabIndex = 44
        '
        'Spec_Indep_Panel
        '
        Me.Spec_Indep_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Indep_Panel.Controls.Add(Me.Spec_Indep_Label)
        Me.Spec_Indep_Panel.Controls.Add(Me.Spec_Indep_ComboBox)
        Me.Spec_Indep_Panel.Location = New System.Drawing.Point(3, 208)
        Me.Spec_Indep_Panel.Name = "Spec_Indep_Panel"
        Me.Spec_Indep_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_Indep_Panel.TabIndex = 176
        '
        'Spec_Indep_Label
        '
        Me.Spec_Indep_Label.AutoSize = True
        Me.Spec_Indep_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Indep_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Indep_Label.Name = "Spec_Indep_Label"
        Me.Spec_Indep_Label.Size = New System.Drawing.Size(56, 16)
        Me.Spec_Indep_Label.TabIndex = 52
        Me.Spec_Indep_Label.Text = "專用運轉"
        '
        'Spec_Indep_ComboBox
        '
        Me.Spec_Indep_ComboBox.FormattingEnabled = True
        Me.Spec_Indep_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Indep_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_Indep_ComboBox.Name = "Spec_Indep_ComboBox"
        Me.Spec_Indep_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Indep_ComboBox.TabIndex = 53
        '
        'Spec_HinCpi_Panel
        '
        Me.Spec_HinCpi_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_HinCpi_Panel.Controls.Add(Me.Spec_HinCpi_Label)
        Me.Spec_HinCpi_Panel.Controls.Add(Me.Spec_HinCpi_ComboBox)
        Me.Spec_HinCpi_Panel.Location = New System.Drawing.Point(3, 250)
        Me.Spec_HinCpi_Panel.Name = "Spec_HinCpi_Panel"
        Me.Spec_HinCpi_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_HinCpi_Panel.TabIndex = 178
        '
        'Spec_HinCpi_Label
        '
        Me.Spec_HinCpi_Label.AutoSize = True
        Me.Spec_HinCpi_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_HinCpi_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_HinCpi_Label.Name = "Spec_HinCpi_Label"
        Me.Spec_HinCpi_Label.Size = New System.Drawing.Size(53, 16)
        Me.Spec_HinCpi_Label.TabIndex = 56
        Me.Spec_HinCpi_Label.Text = "HIN/CPI"
        '
        'Spec_HinCpi_ComboBox
        '
        Me.Spec_HinCpi_ComboBox.FormattingEnabled = True
        Me.Spec_HinCpi_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_HinCpi_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_HinCpi_ComboBox.Name = "Spec_HinCpi_ComboBox"
        Me.Spec_HinCpi_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_HinCpi_ComboBox.TabIndex = 57
        '
        'Spec_Fire_Panel
        '
        Me.Spec_Fire_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Fire_Panel.Controls.Add(Me.Spec_Fire_Only_CheckBox)
        Me.Spec_Fire_Panel.Controls.Add(Me.Label195)
        Me.Spec_Fire_Panel.Controls.Add(Me.Spec_Fire_Only_TextBox)
        Me.Spec_Fire_Panel.Controls.Add(Me.Spec_FireSignal_ComboBox)
        Me.Spec_Fire_Panel.Controls.Add(Me.Spec_FireSignal_Label)
        Me.Spec_Fire_Panel.Controls.Add(Me.Spec_Fire_Label)
        Me.Spec_Fire_Panel.Controls.Add(Me.Spec_Fire_ComboBox)
        Me.Spec_Fire_Panel.Controls.Add(Me.Spec_EscapeFL_TextBox)
        Me.Spec_Fire_Panel.Controls.Add(Me.Spec_EscapeFL_Label)
        Me.Spec_Fire_Panel.Location = New System.Drawing.Point(3, 292)
        Me.Spec_Fire_Panel.Name = "Spec_Fire_Panel"
        Me.Spec_Fire_Panel.Size = New System.Drawing.Size(580, 73)
        Me.Spec_Fire_Panel.TabIndex = 179
        '
        'Spec_Fire_Only_CheckBox
        '
        Me.Spec_Fire_Only_CheckBox.AutoSize = True
        Me.Spec_Fire_Only_CheckBox.Location = New System.Drawing.Point(216, 9)
        Me.Spec_Fire_Only_CheckBox.Name = "Spec_Fire_Only_CheckBox"
        Me.Spec_Fire_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_Fire_Only_CheckBox.TabIndex = 123
        Me.Spec_Fire_Only_CheckBox.Text = "Only"
        Me.Spec_Fire_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label195
        '
        Me.Label195.AutoSize = True
        Me.Label195.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label195.Location = New System.Drawing.Point(386, 11)
        Me.Label195.Name = "Label195"
        Me.Label195.Size = New System.Drawing.Size(32, 16)
        Me.Label195.TabIndex = 126
        Me.Label195.Text = "號機"
        '
        'Spec_Fire_Only_TextBox
        '
        Me.Spec_Fire_Only_TextBox.Enabled = False
        Me.Spec_Fire_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Fire_Only_TextBox.Location = New System.Drawing.Point(274, 8)
        Me.Spec_Fire_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_Fire_Only_TextBox.MaxLength = 50
        Me.Spec_Fire_Only_TextBox.Name = "Spec_Fire_Only_TextBox"
        Me.Spec_Fire_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_Fire_Only_TextBox.TabIndex = 124
        '
        'Spec_FireSignal_ComboBox
        '
        Me.Spec_FireSignal_ComboBox.FormattingEnabled = True
        Me.Spec_FireSignal_ComboBox.Items.AddRange(New Object() {"N/O", "N/C"})
        Me.Spec_FireSignal_ComboBox.Location = New System.Drawing.Point(274, 38)
        Me.Spec_FireSignal_ComboBox.Name = "Spec_FireSignal_ComboBox"
        Me.Spec_FireSignal_ComboBox.Size = New System.Drawing.Size(67, 24)
        Me.Spec_FireSignal_ComboBox.TabIndex = 122
        '
        'Spec_FireSignal_Label
        '
        Me.Spec_FireSignal_Label.AutoSize = True
        Me.Spec_FireSignal_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_FireSignal_Label.Location = New System.Drawing.Point(213, 42)
        Me.Spec_FireSignal_Label.Name = "Spec_FireSignal_Label"
        Me.Spec_FireSignal_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_FireSignal_Label.TabIndex = 121
        Me.Spec_FireSignal_Label.Text = "訊號 :"
        '
        'Spec_Fire_Label
        '
        Me.Spec_Fire_Label.AutoSize = True
        Me.Spec_Fire_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Fire_Label.Location = New System.Drawing.Point(27, 10)
        Me.Spec_Fire_Label.Name = "Spec_Fire_Label"
        Me.Spec_Fire_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_Fire_Label.TabIndex = 58
        Me.Spec_Fire_Label.Text = "火災管制運轉"
        '
        'Spec_Fire_ComboBox
        '
        Me.Spec_Fire_ComboBox.FormattingEnabled = True
        Me.Spec_Fire_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Fire_ComboBox.Location = New System.Drawing.Point(147, 6)
        Me.Spec_Fire_ComboBox.Name = "Spec_Fire_ComboBox"
        Me.Spec_Fire_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Fire_ComboBox.TabIndex = 59
        '
        'Spec_EscapeFL_TextBox
        '
        Me.Spec_EscapeFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_EscapeFL_TextBox.Location = New System.Drawing.Point(474, 39)
        Me.Spec_EscapeFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_EscapeFL_TextBox.MaxLength = 50
        Me.Spec_EscapeFL_TextBox.Name = "Spec_EscapeFL_TextBox"
        Me.Spec_EscapeFL_TextBox.Size = New System.Drawing.Size(61, 23)
        Me.Spec_EscapeFL_TextBox.TabIndex = 115
        '
        'Spec_EscapeFL_Label
        '
        Me.Spec_EscapeFL_Label.AutoSize = True
        Me.Spec_EscapeFL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_EscapeFL_Label.Location = New System.Drawing.Point(416, 42)
        Me.Spec_EscapeFL_Label.Name = "Spec_EscapeFL_Label"
        Me.Spec_EscapeFL_Label.Size = New System.Drawing.Size(50, 16)
        Me.Spec_EscapeFL_Label.TabIndex = 114
        Me.Spec_EscapeFL_Label.Text = "避難階 :"
        '
        'Spec_Fireman_Panel
        '
        Me.Spec_Fireman_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Fireman_Panel.Controls.Add(Me.Spec_Fireman_Only_CheckBox)
        Me.Spec_Fireman_Panel.Controls.Add(Me.Label55)
        Me.Spec_Fireman_Panel.Controls.Add(Me.Spec_Fireman_Only_TextBox)
        Me.Spec_Fireman_Panel.Controls.Add(Me.Spec_Fireman_Label)
        Me.Spec_Fireman_Panel.Controls.Add(Me.Spec_Fireman_ComboBox)
        Me.Spec_Fireman_Panel.Location = New System.Drawing.Point(3, 371)
        Me.Spec_Fireman_Panel.Name = "Spec_Fireman_Panel"
        Me.Spec_Fireman_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_Fireman_Panel.TabIndex = 180
        '
        'Spec_Fireman_Only_CheckBox
        '
        Me.Spec_Fireman_Only_CheckBox.AutoSize = True
        Me.Spec_Fireman_Only_CheckBox.Location = New System.Drawing.Point(216, 7)
        Me.Spec_Fireman_Only_CheckBox.Name = "Spec_Fireman_Only_CheckBox"
        Me.Spec_Fireman_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_Fireman_Only_CheckBox.TabIndex = 18
        Me.Spec_Fireman_Only_CheckBox.Text = "Only"
        Me.Spec_Fireman_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label55
        '
        Me.Label55.AutoSize = True
        Me.Label55.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label55.Location = New System.Drawing.Point(387, 9)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(32, 16)
        Me.Label55.TabIndex = 118
        Me.Label55.Text = "號機"
        '
        'Spec_Fireman_Only_TextBox
        '
        Me.Spec_Fireman_Only_TextBox.Enabled = False
        Me.Spec_Fireman_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Fireman_Only_TextBox.Location = New System.Drawing.Point(275, 6)
        Me.Spec_Fireman_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_Fireman_Only_TextBox.MaxLength = 50
        Me.Spec_Fireman_Only_TextBox.Name = "Spec_Fireman_Only_TextBox"
        Me.Spec_Fireman_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_Fireman_Only_TextBox.TabIndex = 116
        '
        'Spec_Fireman_Label
        '
        Me.Spec_Fireman_Label.AutoSize = True
        Me.Spec_Fireman_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Fireman_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Fireman_Label.Name = "Spec_Fireman_Label"
        Me.Spec_Fireman_Label.Size = New System.Drawing.Size(44, 16)
        Me.Spec_Fireman_Label.TabIndex = 60
        Me.Spec_Fireman_Label.Text = "消防梯"
        '
        'Spec_Fireman_ComboBox
        '
        Me.Spec_Fireman_ComboBox.FormattingEnabled = True
        Me.Spec_Fireman_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Fireman_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_Fireman_ComboBox.Name = "Spec_Fireman_ComboBox"
        Me.Spec_Fireman_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Fireman_ComboBox.TabIndex = 61
        '
        'TabPage10
        '
        Me.TabPage10.Controls.Add(Me.Spec_TW_FlowLayoutPanel2)
        Me.TabPage10.Location = New System.Drawing.Point(4, 25)
        Me.TabPage10.Name = "TabPage10"
        Me.TabPage10.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage10.Size = New System.Drawing.Size(627, 469)
        Me.TabPage10.TabIndex = 1
        Me.TabPage10.Text = "Page2"
        Me.TabPage10.UseVisualStyleBackColor = True
        '
        'Spec_TW_FlowLayoutPanel2
        '
        Me.Spec_TW_FlowLayoutPanel2.AutoScroll = True
        Me.Spec_TW_FlowLayoutPanel2.Controls.Add(Me.Spec_Parking_Panel)
        Me.Spec_TW_FlowLayoutPanel2.Controls.Add(Me.Spec_Seismic_Panel)
        Me.Spec_TW_FlowLayoutPanel2.Controls.Add(Me.Spec_CPI_Panel)
        Me.Spec_TW_FlowLayoutPanel2.Controls.Add(Me.Spec_HallGong_Panel)
        Me.Spec_TW_FlowLayoutPanel2.Controls.Add(Me.Spec_HPIMsg_Panel)
        Me.Spec_TW_FlowLayoutPanel2.Enabled = False
        Me.Spec_TW_FlowLayoutPanel2.Location = New System.Drawing.Point(6, 6)
        Me.Spec_TW_FlowLayoutPanel2.Name = "Spec_TW_FlowLayoutPanel2"
        Me.Spec_TW_FlowLayoutPanel2.Size = New System.Drawing.Size(615, 457)
        Me.Spec_TW_FlowLayoutPanel2.TabIndex = 0
        '
        'Spec_Parking_Panel
        '
        Me.Spec_Parking_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_Parking_Only_CheckBox)
        Me.Spec_Parking_Panel.Controls.Add(Me.Label56)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_Parking_Only_TextBox)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_DR_ComboBox)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_DR_Label)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_HALL_ComboBox)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_HALL_Label)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_COB_ComboBox)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_COB_Label)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_WTB_ComboBox)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_WTB_Label)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_ELVIC_Label)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_ParkingFL_ELVIC_ComboBox)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_Parking_FL_TextBox)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_Parking_FL_Label)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_Parking_Label)
        Me.Spec_Parking_Panel.Controls.Add(Me.Spec_Parking_ComboBox)
        Me.Spec_Parking_Panel.Location = New System.Drawing.Point(3, 3)
        Me.Spec_Parking_Panel.Name = "Spec_Parking_Panel"
        Me.Spec_Parking_Panel.Size = New System.Drawing.Size(580, 112)
        Me.Spec_Parking_Panel.TabIndex = 183
        '
        'Spec_Parking_Only_CheckBox
        '
        Me.Spec_Parking_Only_CheckBox.AutoSize = True
        Me.Spec_Parking_Only_CheckBox.Location = New System.Drawing.Point(214, 8)
        Me.Spec_Parking_Only_CheckBox.Name = "Spec_Parking_Only_CheckBox"
        Me.Spec_Parking_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_Parking_Only_CheckBox.TabIndex = 129
        Me.Spec_Parking_Only_CheckBox.Text = "Only"
        Me.Spec_Parking_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label56
        '
        Me.Label56.AutoSize = True
        Me.Label56.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label56.Location = New System.Drawing.Point(383, 10)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(32, 16)
        Me.Label56.TabIndex = 131
        Me.Label56.Text = "號機"
        '
        'Spec_Parking_Only_TextBox
        '
        Me.Spec_Parking_Only_TextBox.Enabled = False
        Me.Spec_Parking_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Parking_Only_TextBox.Location = New System.Drawing.Point(271, 7)
        Me.Spec_Parking_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_Parking_Only_TextBox.MaxLength = 50
        Me.Spec_Parking_Only_TextBox.Name = "Spec_Parking_Only_TextBox"
        Me.Spec_Parking_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_Parking_Only_TextBox.TabIndex = 130
        '
        'Spec_ParkingFL_DR_ComboBox
        '
        Me.Spec_ParkingFL_DR_ComboBox.FormattingEnabled = True
        Me.Spec_ParkingFL_DR_ComboBox.Location = New System.Drawing.Point(252, 76)
        Me.Spec_ParkingFL_DR_ComboBox.Name = "Spec_ParkingFL_DR_ComboBox"
        Me.Spec_ParkingFL_DR_ComboBox.Size = New System.Drawing.Size(72, 24)
        Me.Spec_ParkingFL_DR_ComboBox.TabIndex = 128
        '
        'Spec_ParkingFL_DR_Label
        '
        Me.Spec_ParkingFL_DR_Label.AutoSize = True
        Me.Spec_ParkingFL_DR_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_ParkingFL_DR_Label.Location = New System.Drawing.Point(211, 80)
        Me.Spec_ParkingFL_DR_Label.Name = "Spec_ParkingFL_DR_Label"
        Me.Spec_ParkingFL_DR_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_ParkingFL_DR_Label.TabIndex = 127
        Me.Spec_ParkingFL_DR_Label.Text = "休止 :"
        '
        'Spec_ParkingFL_HALL_ComboBox
        '
        Me.Spec_ParkingFL_HALL_ComboBox.FormattingEnabled = True
        Me.Spec_ParkingFL_HALL_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_ParkingFL_HALL_ComboBox.Location = New System.Drawing.Point(493, 76)
        Me.Spec_ParkingFL_HALL_ComboBox.Name = "Spec_ParkingFL_HALL_ComboBox"
        Me.Spec_ParkingFL_HALL_ComboBox.Size = New System.Drawing.Size(35, 24)
        Me.Spec_ParkingFL_HALL_ComboBox.TabIndex = 126
        '
        'Spec_ParkingFL_HALL_Label
        '
        Me.Spec_ParkingFL_HALL_Label.AutoSize = True
        Me.Spec_ParkingFL_HALL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_ParkingFL_HALL_Label.Location = New System.Drawing.Point(438, 80)
        Me.Spec_ParkingFL_HALL_Label.Name = "Spec_ParkingFL_HALL_Label"
        Me.Spec_ParkingFL_HALL_Label.Size = New System.Drawing.Size(43, 16)
        Me.Spec_ParkingFL_HALL_Label.TabIndex = 125
        Me.Spec_ParkingFL_HALL_Label.Text = "HALL :"
        '
        'Spec_ParkingFL_COB_ComboBox
        '
        Me.Spec_ParkingFL_COB_ComboBox.FormattingEnabled = True
        Me.Spec_ParkingFL_COB_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_ParkingFL_COB_ComboBox.Location = New System.Drawing.Point(391, 76)
        Me.Spec_ParkingFL_COB_ComboBox.Name = "Spec_ParkingFL_COB_ComboBox"
        Me.Spec_ParkingFL_COB_ComboBox.Size = New System.Drawing.Size(35, 24)
        Me.Spec_ParkingFL_COB_ComboBox.TabIndex = 124
        '
        'Spec_ParkingFL_COB_Label
        '
        Me.Spec_ParkingFL_COB_Label.AutoSize = True
        Me.Spec_ParkingFL_COB_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_ParkingFL_COB_Label.Location = New System.Drawing.Point(340, 80)
        Me.Spec_ParkingFL_COB_Label.Name = "Spec_ParkingFL_COB_Label"
        Me.Spec_ParkingFL_COB_Label.Size = New System.Drawing.Size(39, 16)
        Me.Spec_ParkingFL_COB_Label.TabIndex = 123
        Me.Spec_ParkingFL_COB_Label.Text = "COB :"
        '
        'Spec_ParkingFL_WTB_ComboBox
        '
        Me.Spec_ParkingFL_WTB_ComboBox.FormattingEnabled = True
        Me.Spec_ParkingFL_WTB_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_ParkingFL_WTB_ComboBox.Location = New System.Drawing.Point(493, 40)
        Me.Spec_ParkingFL_WTB_ComboBox.Name = "Spec_ParkingFL_WTB_ComboBox"
        Me.Spec_ParkingFL_WTB_ComboBox.Size = New System.Drawing.Size(35, 24)
        Me.Spec_ParkingFL_WTB_ComboBox.TabIndex = 122
        '
        'Spec_ParkingFL_WTB_Label
        '
        Me.Spec_ParkingFL_WTB_Label.AutoSize = True
        Me.Spec_ParkingFL_WTB_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_ParkingFL_WTB_Label.Location = New System.Drawing.Point(440, 44)
        Me.Spec_ParkingFL_WTB_Label.Name = "Spec_ParkingFL_WTB_Label"
        Me.Spec_ParkingFL_WTB_Label.Size = New System.Drawing.Size(40, 16)
        Me.Spec_ParkingFL_WTB_Label.TabIndex = 121
        Me.Spec_ParkingFL_WTB_Label.Text = "WTB :"
        '
        'Spec_ParkingFL_ELVIC_Label
        '
        Me.Spec_ParkingFL_ELVIC_Label.AutoSize = True
        Me.Spec_ParkingFL_ELVIC_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_ParkingFL_ELVIC_Label.Location = New System.Drawing.Point(333, 44)
        Me.Spec_ParkingFL_ELVIC_Label.Name = "Spec_ParkingFL_ELVIC_Label"
        Me.Spec_ParkingFL_ELVIC_Label.Size = New System.Drawing.Size(46, 16)
        Me.Spec_ParkingFL_ELVIC_Label.TabIndex = 119
        Me.Spec_ParkingFL_ELVIC_Label.Text = "ELVIC :"
        '
        'Spec_ParkingFL_ELVIC_ComboBox
        '
        Me.Spec_ParkingFL_ELVIC_ComboBox.FormattingEnabled = True
        Me.Spec_ParkingFL_ELVIC_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_ParkingFL_ELVIC_ComboBox.Location = New System.Drawing.Point(392, 40)
        Me.Spec_ParkingFL_ELVIC_ComboBox.Name = "Spec_ParkingFL_ELVIC_ComboBox"
        Me.Spec_ParkingFL_ELVIC_ComboBox.Size = New System.Drawing.Size(35, 24)
        Me.Spec_ParkingFL_ELVIC_ComboBox.TabIndex = 118
        '
        'Spec_Parking_FL_TextBox
        '
        Me.Spec_Parking_FL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Parking_FL_TextBox.Location = New System.Drawing.Point(263, 41)
        Me.Spec_Parking_FL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_Parking_FL_TextBox.MaxLength = 50
        Me.Spec_Parking_FL_TextBox.Name = "Spec_Parking_FL_TextBox"
        Me.Spec_Parking_FL_TextBox.Size = New System.Drawing.Size(61, 23)
        Me.Spec_Parking_FL_TextBox.TabIndex = 117
        '
        'Spec_Parking_FL_Label
        '
        Me.Spec_Parking_FL_Label.AutoSize = True
        Me.Spec_Parking_FL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Parking_FL_Label.Location = New System.Drawing.Point(211, 44)
        Me.Spec_Parking_FL_Label.Name = "Spec_Parking_FL_Label"
        Me.Spec_Parking_FL_Label.Size = New System.Drawing.Size(50, 16)
        Me.Spec_Parking_FL_Label.TabIndex = 116
        Me.Spec_Parking_FL_Label.Text = "停車階 :"
        '
        'Spec_Parking_Label
        '
        Me.Spec_Parking_Label.AutoSize = True
        Me.Spec_Parking_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Parking_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Parking_Label.Name = "Spec_Parking_Label"
        Me.Spec_Parking_Label.Size = New System.Drawing.Size(68, 16)
        Me.Spec_Parking_Label.TabIndex = 64
        Me.Spec_Parking_Label.Text = "停車階運轉"
        '
        'Spec_Parking_ComboBox
        '
        Me.Spec_Parking_ComboBox.FormattingEnabled = True
        Me.Spec_Parking_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Parking_ComboBox.Location = New System.Drawing.Point(147, 6)
        Me.Spec_Parking_ComboBox.Name = "Spec_Parking_ComboBox"
        Me.Spec_Parking_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Parking_ComboBox.TabIndex = 65
        '
        'Spec_Seismic_Panel
        '
        Me.Spec_Seismic_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_SeismicSW_Only_CheckBox)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Label215)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_SeismicSW_Only_TextBox)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_SeismicSensor_Only_CheckBox)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Label214)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_SeismicSensor_Only_TextBox)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_Seismic_Only_CheckBox)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Label196)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_Seismic_Only_TextBox)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_SeismicSensor_ComboBox)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_SeismicSensor_Label)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_SeismicSW_Label)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_SeismicSW_ComboBox)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_Seismic_Label)
        Me.Spec_Seismic_Panel.Controls.Add(Me.Spec_Seismic_ComboBox)
        Me.Spec_Seismic_Panel.Location = New System.Drawing.Point(3, 121)
        Me.Spec_Seismic_Panel.Name = "Spec_Seismic_Panel"
        Me.Spec_Seismic_Panel.Size = New System.Drawing.Size(580, 116)
        Me.Spec_Seismic_Panel.TabIndex = 184
        '
        'Spec_SeismicSW_Only_CheckBox
        '
        Me.Spec_SeismicSW_Only_CheckBox.AutoSize = True
        Me.Spec_SeismicSW_Only_CheckBox.Location = New System.Drawing.Point(350, 82)
        Me.Spec_SeismicSW_Only_CheckBox.Name = "Spec_SeismicSW_Only_CheckBox"
        Me.Spec_SeismicSW_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_SeismicSW_Only_CheckBox.TabIndex = 138
        Me.Spec_SeismicSW_Only_CheckBox.Text = "Only"
        Me.Spec_SeismicSW_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label215
        '
        Me.Label215.AutoSize = True
        Me.Label215.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label215.Location = New System.Drawing.Point(521, 84)
        Me.Label215.Name = "Label215"
        Me.Label215.Size = New System.Drawing.Size(32, 16)
        Me.Label215.TabIndex = 140
        Me.Label215.Text = "號機"
        '
        'Spec_SeismicSW_Only_TextBox
        '
        Me.Spec_SeismicSW_Only_TextBox.Enabled = False
        Me.Spec_SeismicSW_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_SeismicSW_Only_TextBox.Location = New System.Drawing.Point(409, 81)
        Me.Spec_SeismicSW_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_SeismicSW_Only_TextBox.MaxLength = 50
        Me.Spec_SeismicSW_Only_TextBox.Name = "Spec_SeismicSW_Only_TextBox"
        Me.Spec_SeismicSW_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_SeismicSW_Only_TextBox.TabIndex = 139
        '
        'Spec_SeismicSensor_Only_CheckBox
        '
        Me.Spec_SeismicSensor_Only_CheckBox.AutoSize = True
        Me.Spec_SeismicSensor_Only_CheckBox.Location = New System.Drawing.Point(350, 42)
        Me.Spec_SeismicSensor_Only_CheckBox.Name = "Spec_SeismicSensor_Only_CheckBox"
        Me.Spec_SeismicSensor_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_SeismicSensor_Only_CheckBox.TabIndex = 135
        Me.Spec_SeismicSensor_Only_CheckBox.Text = "Only"
        Me.Spec_SeismicSensor_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label214
        '
        Me.Label214.AutoSize = True
        Me.Label214.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label214.Location = New System.Drawing.Point(521, 44)
        Me.Label214.Name = "Label214"
        Me.Label214.Size = New System.Drawing.Size(32, 16)
        Me.Label214.TabIndex = 137
        Me.Label214.Text = "號機"
        '
        'Spec_SeismicSensor_Only_TextBox
        '
        Me.Spec_SeismicSensor_Only_TextBox.Enabled = False
        Me.Spec_SeismicSensor_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_SeismicSensor_Only_TextBox.Location = New System.Drawing.Point(409, 41)
        Me.Spec_SeismicSensor_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_SeismicSensor_Only_TextBox.MaxLength = 50
        Me.Spec_SeismicSensor_Only_TextBox.Name = "Spec_SeismicSensor_Only_TextBox"
        Me.Spec_SeismicSensor_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_SeismicSensor_Only_TextBox.TabIndex = 136
        '
        'Spec_Seismic_Only_CheckBox
        '
        Me.Spec_Seismic_Only_CheckBox.AutoSize = True
        Me.Spec_Seismic_Only_CheckBox.Location = New System.Drawing.Point(206, 7)
        Me.Spec_Seismic_Only_CheckBox.Name = "Spec_Seismic_Only_CheckBox"
        Me.Spec_Seismic_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_Seismic_Only_CheckBox.TabIndex = 132
        Me.Spec_Seismic_Only_CheckBox.Text = "Only"
        Me.Spec_Seismic_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label196
        '
        Me.Label196.AutoSize = True
        Me.Label196.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label196.Location = New System.Drawing.Point(375, 9)
        Me.Label196.Name = "Label196"
        Me.Label196.Size = New System.Drawing.Size(32, 16)
        Me.Label196.TabIndex = 134
        Me.Label196.Text = "號機"
        '
        'Spec_Seismic_Only_TextBox
        '
        Me.Spec_Seismic_Only_TextBox.Enabled = False
        Me.Spec_Seismic_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Seismic_Only_TextBox.Location = New System.Drawing.Point(263, 6)
        Me.Spec_Seismic_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_Seismic_Only_TextBox.MaxLength = 50
        Me.Spec_Seismic_Only_TextBox.Name = "Spec_Seismic_Only_TextBox"
        Me.Spec_Seismic_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_Seismic_Only_TextBox.TabIndex = 133
        '
        'Spec_SeismicSensor_ComboBox
        '
        Me.Spec_SeismicSensor_ComboBox.FormattingEnabled = True
        Me.Spec_SeismicSensor_ComboBox.Items.AddRange(New Object() {"1", "2", "3"})
        Me.Spec_SeismicSensor_ComboBox.Location = New System.Drawing.Point(273, 40)
        Me.Spec_SeismicSensor_ComboBox.Name = "Spec_SeismicSensor_ComboBox"
        Me.Spec_SeismicSensor_ComboBox.Size = New System.Drawing.Size(64, 24)
        Me.Spec_SeismicSensor_ComboBox.TabIndex = 70
        '
        'Spec_SeismicSensor_Label
        '
        Me.Spec_SeismicSensor_Label.AutoSize = True
        Me.Spec_SeismicSensor_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_SeismicSensor_Label.Location = New System.Drawing.Point(197, 44)
        Me.Spec_SeismicSensor_Label.Name = "Spec_SeismicSensor_Label"
        Me.Spec_SeismicSensor_Label.Size = New System.Drawing.Size(75, 16)
        Me.Spec_SeismicSensor_Label.TabIndex = 68
        Me.Spec_SeismicSensor_Label.Text = "感知器N段 : "
        '
        'Spec_SeismicSW_Label
        '
        Me.Spec_SeismicSW_Label.AutoSize = True
        Me.Spec_SeismicSW_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_SeismicSW_Label.Location = New System.Drawing.Point(183, 84)
        Me.Spec_SeismicSW_Label.Name = "Spec_SeismicSW_Label"
        Me.Spec_SeismicSW_Label.Size = New System.Drawing.Size(89, 16)
        Me.Spec_SeismicSW_Label.TabIndex = 68
        Me.Spec_SeismicSW_Label.Text = "自動解除開關 : "
        '
        'Spec_SeismicSW_ComboBox
        '
        Me.Spec_SeismicSW_ComboBox.FormattingEnabled = True
        Me.Spec_SeismicSW_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_SeismicSW_ComboBox.Location = New System.Drawing.Point(274, 80)
        Me.Spec_SeismicSW_ComboBox.Name = "Spec_SeismicSW_ComboBox"
        Me.Spec_SeismicSW_ComboBox.Size = New System.Drawing.Size(64, 24)
        Me.Spec_SeismicSW_ComboBox.TabIndex = 69
        '
        'Spec_Seismic_Label
        '
        Me.Spec_Seismic_Label.AutoSize = True
        Me.Spec_Seismic_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Seismic_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Seismic_Label.Name = "Spec_Seismic_Label"
        Me.Spec_Seismic_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_Seismic_Label.TabIndex = 66
        Me.Spec_Seismic_Label.Text = "地震管制運轉"
        '
        'Spec_Seismic_ComboBox
        '
        Me.Spec_Seismic_ComboBox.FormattingEnabled = True
        Me.Spec_Seismic_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Seismic_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_Seismic_ComboBox.Name = "Spec_Seismic_ComboBox"
        Me.Spec_Seismic_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Seismic_ComboBox.TabIndex = 67
        '
        'Spec_CPI_Panel
        '
        Me.Spec_CPI_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiOLT_Only_CheckBox)
        Me.Spec_CPI_Panel.Controls.Add(Me.Label216)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiOLT_Only_TextBox)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiOLT_ComboBox)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiOLT_Label)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiFM_ComboBox)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiFM_Label)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiEmer_ComboBox)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiEmer_Label)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiFire_ComboBox)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiFire_Label)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiSeismic_ComboBox)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CpiSeismic_Label)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CPI_Label)
        Me.Spec_CPI_Panel.Controls.Add(Me.Spec_CPI_ComboBox)
        Me.Spec_CPI_Panel.Location = New System.Drawing.Point(3, 243)
        Me.Spec_CPI_Panel.Name = "Spec_CPI_Panel"
        Me.Spec_CPI_Panel.Size = New System.Drawing.Size(580, 107)
        Me.Spec_CPI_Panel.TabIndex = 185
        '
        'Spec_CpiOLT_Only_CheckBox
        '
        Me.Spec_CpiOLT_Only_CheckBox.AutoSize = True
        Me.Spec_CpiOLT_Only_CheckBox.Location = New System.Drawing.Point(309, 70)
        Me.Spec_CpiOLT_Only_CheckBox.Name = "Spec_CpiOLT_Only_CheckBox"
        Me.Spec_CpiOLT_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_CpiOLT_Only_CheckBox.TabIndex = 141
        Me.Spec_CpiOLT_Only_CheckBox.Text = "Only"
        Me.Spec_CpiOLT_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label216
        '
        Me.Label216.AutoSize = True
        Me.Label216.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label216.Location = New System.Drawing.Point(480, 72)
        Me.Label216.Name = "Label216"
        Me.Label216.Size = New System.Drawing.Size(32, 16)
        Me.Label216.TabIndex = 143
        Me.Label216.Text = "號機"
        '
        'Spec_CpiOLT_Only_TextBox
        '
        Me.Spec_CpiOLT_Only_TextBox.Enabled = False
        Me.Spec_CpiOLT_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CpiOLT_Only_TextBox.Location = New System.Drawing.Point(368, 69)
        Me.Spec_CpiOLT_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CpiOLT_Only_TextBox.MaxLength = 50
        Me.Spec_CpiOLT_Only_TextBox.Name = "Spec_CpiOLT_Only_TextBox"
        Me.Spec_CpiOLT_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_CpiOLT_Only_TextBox.TabIndex = 142
        '
        'Spec_CpiOLT_ComboBox
        '
        Me.Spec_CpiOLT_ComboBox.FormattingEnabled = True
        Me.Spec_CpiOLT_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CpiOLT_ComboBox.Location = New System.Drawing.Point(249, 68)
        Me.Spec_CpiOLT_ComboBox.Name = "Spec_CpiOLT_ComboBox"
        Me.Spec_CpiOLT_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CpiOLT_ComboBox.TabIndex = 79
        '
        'Spec_CpiOLT_Label
        '
        Me.Spec_CpiOLT_Label.AutoSize = True
        Me.Spec_CpiOLT_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CpiOLT_Label.Location = New System.Drawing.Point(208, 72)
        Me.Spec_CpiOLT_Label.Name = "Spec_CpiOLT_Label"
        Me.Spec_CpiOLT_Label.Size = New System.Drawing.Size(41, 16)
        Me.Spec_CpiOLT_Label.TabIndex = 78
        Me.Spec_CpiOLT_Label.Text = "滿載 : "
        '
        'Spec_CpiFM_ComboBox
        '
        Me.Spec_CpiFM_ComboBox.FormattingEnabled = True
        Me.Spec_CpiFM_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CpiFM_ComboBox.Location = New System.Drawing.Point(249, 38)
        Me.Spec_CpiFM_ComboBox.Name = "Spec_CpiFM_ComboBox"
        Me.Spec_CpiFM_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CpiFM_ComboBox.TabIndex = 77
        '
        'Spec_CpiFM_Label
        '
        Me.Spec_CpiFM_Label.AutoSize = True
        Me.Spec_CpiFM_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CpiFM_Label.Location = New System.Drawing.Point(208, 42)
        Me.Spec_CpiFM_Label.Name = "Spec_CpiFM_Label"
        Me.Spec_CpiFM_Label.Size = New System.Drawing.Size(41, 16)
        Me.Spec_CpiFM_Label.TabIndex = 76
        Me.Spec_CpiFM_Label.Text = "緊急 : "
        '
        'Spec_CpiEmer_ComboBox
        '
        Me.Spec_CpiEmer_ComboBox.FormattingEnabled = True
        Me.Spec_CpiEmer_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CpiEmer_ComboBox.Location = New System.Drawing.Point(368, 38)
        Me.Spec_CpiEmer_ComboBox.Name = "Spec_CpiEmer_ComboBox"
        Me.Spec_CpiEmer_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CpiEmer_ComboBox.TabIndex = 75
        '
        'Spec_CpiEmer_Label
        '
        Me.Spec_CpiEmer_Label.AutoSize = True
        Me.Spec_CpiEmer_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CpiEmer_Label.Location = New System.Drawing.Point(327, 42)
        Me.Spec_CpiEmer_Label.Name = "Spec_CpiEmer_Label"
        Me.Spec_CpiEmer_Label.Size = New System.Drawing.Size(41, 16)
        Me.Spec_CpiEmer_Label.TabIndex = 74
        Me.Spec_CpiEmer_Label.Text = "自發 : "
        '
        'Spec_CpiFire_ComboBox
        '
        Me.Spec_CpiFire_ComboBox.FormattingEnabled = True
        Me.Spec_CpiFire_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CpiFire_ComboBox.Location = New System.Drawing.Point(368, 5)
        Me.Spec_CpiFire_ComboBox.Name = "Spec_CpiFire_ComboBox"
        Me.Spec_CpiFire_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CpiFire_ComboBox.TabIndex = 73
        '
        'Spec_CpiFire_Label
        '
        Me.Spec_CpiFire_Label.AutoSize = True
        Me.Spec_CpiFire_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CpiFire_Label.Location = New System.Drawing.Point(327, 9)
        Me.Spec_CpiFire_Label.Name = "Spec_CpiFire_Label"
        Me.Spec_CpiFire_Label.Size = New System.Drawing.Size(41, 16)
        Me.Spec_CpiFire_Label.TabIndex = 72
        Me.Spec_CpiFire_Label.Text = "火災 : "
        '
        'Spec_CpiSeismic_ComboBox
        '
        Me.Spec_CpiSeismic_ComboBox.FormattingEnabled = True
        Me.Spec_CpiSeismic_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CpiSeismic_ComboBox.Location = New System.Drawing.Point(249, 5)
        Me.Spec_CpiSeismic_ComboBox.Name = "Spec_CpiSeismic_ComboBox"
        Me.Spec_CpiSeismic_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CpiSeismic_ComboBox.TabIndex = 71
        '
        'Spec_CpiSeismic_Label
        '
        Me.Spec_CpiSeismic_Label.AutoSize = True
        Me.Spec_CpiSeismic_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CpiSeismic_Label.Location = New System.Drawing.Point(208, 9)
        Me.Spec_CpiSeismic_Label.Name = "Spec_CpiSeismic_Label"
        Me.Spec_CpiSeismic_Label.Size = New System.Drawing.Size(41, 16)
        Me.Spec_CpiSeismic_Label.TabIndex = 70
        Me.Spec_CpiSeismic_Label.Text = "地震 : "
        '
        'Spec_CPI_Label
        '
        Me.Spec_CPI_Label.AutoSize = True
        Me.Spec_CPI_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CPI_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_CPI_Label.Name = "Spec_CPI_Label"
        Me.Spec_CPI_Label.Size = New System.Drawing.Size(92, 16)
        Me.Spec_CPI_Label.TabIndex = 68
        Me.Spec_CPI_Label.Text = "車廂管制運轉燈"
        '
        'Spec_CPI_ComboBox
        '
        Me.Spec_CPI_ComboBox.FormattingEnabled = True
        Me.Spec_CPI_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CPI_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_CPI_ComboBox.Name = "Spec_CPI_ComboBox"
        Me.Spec_CPI_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CPI_ComboBox.TabIndex = 69
        '
        'Spec_HallGong_Panel
        '
        Me.Spec_HallGong_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_HallGong_Panel.Controls.Add(Me.Spec_HallGong_Label)
        Me.Spec_HallGong_Panel.Controls.Add(Me.Spec_HallGong_ComboBox)
        Me.Spec_HallGong_Panel.Location = New System.Drawing.Point(3, 356)
        Me.Spec_HallGong_Panel.Name = "Spec_HallGong_Panel"
        Me.Spec_HallGong_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_HallGong_Panel.TabIndex = 187
        '
        'Spec_HallGong_Label
        '
        Me.Spec_HallGong_Label.AutoSize = True
        Me.Spec_HallGong_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_HallGong_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_HallGong_Label.Name = "Spec_HallGong_Label"
        Me.Spec_HallGong_Label.Size = New System.Drawing.Size(68, 16)
        Me.Spec_HallGong_Label.TabIndex = 72
        Me.Spec_HallGong_Label.Text = "乘場到著鈴"
        '
        'Spec_HallGong_ComboBox
        '
        Me.Spec_HallGong_ComboBox.FormattingEnabled = True
        Me.Spec_HallGong_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_HallGong_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_HallGong_ComboBox.Name = "Spec_HallGong_ComboBox"
        Me.Spec_HallGong_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_HallGong_ComboBox.TabIndex = 73
        '
        'Spec_HPIMsg_Panel
        '
        Me.Spec_HPIMsg_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HpiFM_ComboBox)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HpiFM_Label)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HpiIndep_ComboBox)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HpiIndep_Label)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HpiMain_ComboBox)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HpiMain_Label)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HpiOLT_ComboBox)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HpiOLT_Label)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HPIMsg_Label)
        Me.Spec_HPIMsg_Panel.Controls.Add(Me.Spec_HPIMsg_ComboBox)
        Me.Spec_HPIMsg_Panel.Location = New System.Drawing.Point(3, 398)
        Me.Spec_HPIMsg_Panel.Name = "Spec_HPIMsg_Panel"
        Me.Spec_HPIMsg_Panel.Size = New System.Drawing.Size(580, 75)
        Me.Spec_HPIMsg_Panel.TabIndex = 188
        '
        'Spec_HpiFM_ComboBox
        '
        Me.Spec_HpiFM_ComboBox.FormattingEnabled = True
        Me.Spec_HpiFM_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_HpiFM_ComboBox.Location = New System.Drawing.Point(252, 39)
        Me.Spec_HpiFM_ComboBox.Name = "Spec_HpiFM_ComboBox"
        Me.Spec_HpiFM_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_HpiFM_ComboBox.TabIndex = 83
        '
        'Spec_HpiFM_Label
        '
        Me.Spec_HpiFM_Label.AutoSize = True
        Me.Spec_HpiFM_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_HpiFM_Label.Location = New System.Drawing.Point(205, 43)
        Me.Spec_HpiFM_Label.Name = "Spec_HpiFM_Label"
        Me.Spec_HpiFM_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_HpiFM_Label.TabIndex = 82
        Me.Spec_HpiFM_Label.Text = "緊急 :"
        '
        'Spec_HpiIndep_ComboBox
        '
        Me.Spec_HpiIndep_ComboBox.FormattingEnabled = True
        Me.Spec_HpiIndep_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_HpiIndep_ComboBox.Location = New System.Drawing.Point(448, 5)
        Me.Spec_HpiIndep_ComboBox.Name = "Spec_HpiIndep_ComboBox"
        Me.Spec_HpiIndep_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_HpiIndep_ComboBox.TabIndex = 81
        '
        'Spec_HpiIndep_Label
        '
        Me.Spec_HpiIndep_Label.AutoSize = True
        Me.Spec_HpiIndep_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_HpiIndep_Label.Location = New System.Drawing.Point(401, 9)
        Me.Spec_HpiIndep_Label.Name = "Spec_HpiIndep_Label"
        Me.Spec_HpiIndep_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_HpiIndep_Label.TabIndex = 80
        Me.Spec_HpiIndep_Label.Text = "專用 :"
        '
        'Spec_HpiMain_ComboBox
        '
        Me.Spec_HpiMain_ComboBox.FormattingEnabled = True
        Me.Spec_HpiMain_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_HpiMain_ComboBox.Location = New System.Drawing.Point(350, 5)
        Me.Spec_HpiMain_ComboBox.Name = "Spec_HpiMain_ComboBox"
        Me.Spec_HpiMain_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_HpiMain_ComboBox.TabIndex = 79
        '
        'Spec_HpiMain_Label
        '
        Me.Spec_HpiMain_Label.AutoSize = True
        Me.Spec_HpiMain_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_HpiMain_Label.Location = New System.Drawing.Point(303, 9)
        Me.Spec_HpiMain_Label.Name = "Spec_HpiMain_Label"
        Me.Spec_HpiMain_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_HpiMain_Label.TabIndex = 78
        Me.Spec_HpiMain_Label.Text = "保養 :"
        '
        'Spec_HpiOLT_ComboBox
        '
        Me.Spec_HpiOLT_ComboBox.FormattingEnabled = True
        Me.Spec_HpiOLT_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_HpiOLT_ComboBox.Location = New System.Drawing.Point(252, 5)
        Me.Spec_HpiOLT_ComboBox.Name = "Spec_HpiOLT_ComboBox"
        Me.Spec_HpiOLT_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_HpiOLT_ComboBox.TabIndex = 77
        '
        'Spec_HpiOLT_Label
        '
        Me.Spec_HpiOLT_Label.AutoSize = True
        Me.Spec_HpiOLT_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_HpiOLT_Label.Location = New System.Drawing.Point(205, 9)
        Me.Spec_HpiOLT_Label.Name = "Spec_HpiOLT_Label"
        Me.Spec_HpiOLT_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_HpiOLT_Label.TabIndex = 76
        Me.Spec_HpiOLT_Label.Text = "滿載 :"
        '
        'Spec_HPIMsg_Label
        '
        Me.Spec_HPIMsg_Label.AutoSize = True
        Me.Spec_HPIMsg_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_HPIMsg_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_HPIMsg_Label.Name = "Spec_HPIMsg_Label"
        Me.Spec_HPIMsg_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_HPIMsg_Label.TabIndex = 74
        Me.Spec_HPIMsg_Label.Text = "乘場信號文字"
        '
        'Spec_HPIMsg_ComboBox
        '
        Me.Spec_HPIMsg_ComboBox.FormattingEnabled = True
        Me.Spec_HPIMsg_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_HPIMsg_ComboBox.Location = New System.Drawing.Point(147, 8)
        Me.Spec_HPIMsg_ComboBox.Name = "Spec_HPIMsg_ComboBox"
        Me.Spec_HPIMsg_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_HPIMsg_ComboBox.TabIndex = 75
        '
        'TabPage12
        '
        Me.TabPage12.Controls.Add(Me.Spec_TW_FlowLayoutPanel3)
        Me.TabPage12.Location = New System.Drawing.Point(4, 25)
        Me.TabPage12.Name = "TabPage12"
        Me.TabPage12.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage12.Size = New System.Drawing.Size(627, 469)
        Me.TabPage12.TabIndex = 3
        Me.TabPage12.Text = "Page3"
        Me.TabPage12.UseVisualStyleBackColor = True
        '
        'Spec_TW_FlowLayoutPanel3
        '
        Me.Spec_TW_FlowLayoutPanel3.AutoScroll = True
        Me.Spec_TW_FlowLayoutPanel3.Controls.Add(Me.Spec_CarGong_Panel)
        Me.Spec_TW_FlowLayoutPanel3.Controls.Add(Me.Spec_CRD_Panel)
        Me.Spec_TW_FlowLayoutPanel3.Enabled = False
        Me.Spec_TW_FlowLayoutPanel3.Location = New System.Drawing.Point(6, 6)
        Me.Spec_TW_FlowLayoutPanel3.Name = "Spec_TW_FlowLayoutPanel3"
        Me.Spec_TW_FlowLayoutPanel3.Size = New System.Drawing.Size(615, 457)
        Me.Spec_TW_FlowLayoutPanel3.TabIndex = 1
        '
        'Spec_CarGong_Panel
        '
        Me.Spec_CarGong_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_VONIC_TextBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_COB_TextBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_TopBtm_TextBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_Top_TextBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_VONIC_Only_CheckBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Label225)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_VONIC_Only_TextBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_COB_Only_CheckBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Label224)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_COB_Only_TextBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_TopBtm_Only_CheckBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Label79)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_TopBtm_Only_TextBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_VONIC_CheckBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_COB_CheckBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_TopBtm_CheckBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_Top_CheckBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_Top_Only_CheckBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Label223)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_Top_Only_TextBox)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_Label)
        Me.Spec_CarGong_Panel.Controls.Add(Me.Spec_CarGong_ComboBox)
        Me.Spec_CarGong_Panel.Location = New System.Drawing.Point(3, 3)
        Me.Spec_CarGong_Panel.Name = "Spec_CarGong_Panel"
        Me.Spec_CarGong_Panel.Size = New System.Drawing.Size(580, 155)
        Me.Spec_CarGong_Panel.TabIndex = 186
        '
        'Spec_CarGong_VONIC_TextBox
        '
        Me.Spec_CarGong_VONIC_TextBox.Enabled = False
        Me.Spec_CarGong_VONIC_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_VONIC_TextBox.Location = New System.Drawing.Point(227, 119)
        Me.Spec_CarGong_VONIC_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CarGong_VONIC_TextBox.MaxLength = 50
        Me.Spec_CarGong_VONIC_TextBox.Name = "Spec_CarGong_VONIC_TextBox"
        Me.Spec_CarGong_VONIC_TextBox.Size = New System.Drawing.Size(143, 23)
        Me.Spec_CarGong_VONIC_TextBox.TabIndex = 163
        '
        'Spec_CarGong_COB_TextBox
        '
        Me.Spec_CarGong_COB_TextBox.Enabled = False
        Me.Spec_CarGong_COB_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_COB_TextBox.Location = New System.Drawing.Point(227, 85)
        Me.Spec_CarGong_COB_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CarGong_COB_TextBox.MaxLength = 50
        Me.Spec_CarGong_COB_TextBox.Name = "Spec_CarGong_COB_TextBox"
        Me.Spec_CarGong_COB_TextBox.Size = New System.Drawing.Size(143, 23)
        Me.Spec_CarGong_COB_TextBox.TabIndex = 162
        '
        'Spec_CarGong_TopBtm_TextBox
        '
        Me.Spec_CarGong_TopBtm_TextBox.Enabled = False
        Me.Spec_CarGong_TopBtm_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_TopBtm_TextBox.Location = New System.Drawing.Point(227, 47)
        Me.Spec_CarGong_TopBtm_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CarGong_TopBtm_TextBox.MaxLength = 50
        Me.Spec_CarGong_TopBtm_TextBox.Name = "Spec_CarGong_TopBtm_TextBox"
        Me.Spec_CarGong_TopBtm_TextBox.Size = New System.Drawing.Size(143, 23)
        Me.Spec_CarGong_TopBtm_TextBox.TabIndex = 161
        '
        'Spec_CarGong_Top_TextBox
        '
        Me.Spec_CarGong_Top_TextBox.Enabled = False
        Me.Spec_CarGong_Top_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_Top_TextBox.Location = New System.Drawing.Point(227, 9)
        Me.Spec_CarGong_Top_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CarGong_Top_TextBox.MaxLength = 50
        Me.Spec_CarGong_Top_TextBox.Name = "Spec_CarGong_Top_TextBox"
        Me.Spec_CarGong_Top_TextBox.Size = New System.Drawing.Size(143, 23)
        Me.Spec_CarGong_Top_TextBox.TabIndex = 160
        '
        'Spec_CarGong_VONIC_Only_CheckBox
        '
        Me.Spec_CarGong_VONIC_Only_CheckBox.AutoSize = True
        Me.Spec_CarGong_VONIC_Only_CheckBox.Enabled = False
        Me.Spec_CarGong_VONIC_Only_CheckBox.Location = New System.Drawing.Point(377, 120)
        Me.Spec_CarGong_VONIC_Only_CheckBox.Name = "Spec_CarGong_VONIC_Only_CheckBox"
        Me.Spec_CarGong_VONIC_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_CarGong_VONIC_Only_CheckBox.TabIndex = 157
        Me.Spec_CarGong_VONIC_Only_CheckBox.Text = "Only"
        Me.Spec_CarGong_VONIC_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label225
        '
        Me.Label225.AutoSize = True
        Me.Label225.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label225.Location = New System.Drawing.Point(544, 122)
        Me.Label225.Name = "Label225"
        Me.Label225.Size = New System.Drawing.Size(32, 16)
        Me.Label225.TabIndex = 159
        Me.Label225.Text = "號機"
        '
        'Spec_CarGong_VONIC_Only_TextBox
        '
        Me.Spec_CarGong_VONIC_Only_TextBox.Enabled = False
        Me.Spec_CarGong_VONIC_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_VONIC_Only_TextBox.Location = New System.Drawing.Point(432, 119)
        Me.Spec_CarGong_VONIC_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CarGong_VONIC_Only_TextBox.MaxLength = 50
        Me.Spec_CarGong_VONIC_Only_TextBox.Name = "Spec_CarGong_VONIC_Only_TextBox"
        Me.Spec_CarGong_VONIC_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_CarGong_VONIC_Only_TextBox.TabIndex = 158
        '
        'Spec_CarGong_COB_Only_CheckBox
        '
        Me.Spec_CarGong_COB_Only_CheckBox.AutoSize = True
        Me.Spec_CarGong_COB_Only_CheckBox.Enabled = False
        Me.Spec_CarGong_COB_Only_CheckBox.Location = New System.Drawing.Point(377, 86)
        Me.Spec_CarGong_COB_Only_CheckBox.Name = "Spec_CarGong_COB_Only_CheckBox"
        Me.Spec_CarGong_COB_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_CarGong_COB_Only_CheckBox.TabIndex = 154
        Me.Spec_CarGong_COB_Only_CheckBox.Text = "Only"
        Me.Spec_CarGong_COB_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label224
        '
        Me.Label224.AutoSize = True
        Me.Label224.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label224.Location = New System.Drawing.Point(544, 88)
        Me.Label224.Name = "Label224"
        Me.Label224.Size = New System.Drawing.Size(32, 16)
        Me.Label224.TabIndex = 156
        Me.Label224.Text = "號機"
        '
        'Spec_CarGong_COB_Only_TextBox
        '
        Me.Spec_CarGong_COB_Only_TextBox.Enabled = False
        Me.Spec_CarGong_COB_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_COB_Only_TextBox.Location = New System.Drawing.Point(432, 85)
        Me.Spec_CarGong_COB_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CarGong_COB_Only_TextBox.MaxLength = 50
        Me.Spec_CarGong_COB_Only_TextBox.Name = "Spec_CarGong_COB_Only_TextBox"
        Me.Spec_CarGong_COB_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_CarGong_COB_Only_TextBox.TabIndex = 155
        '
        'Spec_CarGong_TopBtm_Only_CheckBox
        '
        Me.Spec_CarGong_TopBtm_Only_CheckBox.AutoSize = True
        Me.Spec_CarGong_TopBtm_Only_CheckBox.Enabled = False
        Me.Spec_CarGong_TopBtm_Only_CheckBox.Location = New System.Drawing.Point(377, 48)
        Me.Spec_CarGong_TopBtm_Only_CheckBox.Name = "Spec_CarGong_TopBtm_Only_CheckBox"
        Me.Spec_CarGong_TopBtm_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_CarGong_TopBtm_Only_CheckBox.TabIndex = 151
        Me.Spec_CarGong_TopBtm_Only_CheckBox.Text = "Only"
        Me.Spec_CarGong_TopBtm_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label79
        '
        Me.Label79.AutoSize = True
        Me.Label79.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label79.Location = New System.Drawing.Point(544, 50)
        Me.Label79.Name = "Label79"
        Me.Label79.Size = New System.Drawing.Size(32, 16)
        Me.Label79.TabIndex = 153
        Me.Label79.Text = "號機"
        '
        'Spec_CarGong_TopBtm_Only_TextBox
        '
        Me.Spec_CarGong_TopBtm_Only_TextBox.Enabled = False
        Me.Spec_CarGong_TopBtm_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_TopBtm_Only_TextBox.Location = New System.Drawing.Point(432, 47)
        Me.Spec_CarGong_TopBtm_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CarGong_TopBtm_Only_TextBox.MaxLength = 50
        Me.Spec_CarGong_TopBtm_Only_TextBox.Name = "Spec_CarGong_TopBtm_Only_TextBox"
        Me.Spec_CarGong_TopBtm_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_CarGong_TopBtm_Only_TextBox.TabIndex = 152
        '
        'Spec_CarGong_VONIC_CheckBox
        '
        Me.Spec_CarGong_VONIC_CheckBox.AutoSize = True
        Me.Spec_CarGong_VONIC_CheckBox.Location = New System.Drawing.Point(203, 123)
        Me.Spec_CarGong_VONIC_CheckBox.Name = "Spec_CarGong_VONIC_CheckBox"
        Me.Spec_CarGong_VONIC_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Spec_CarGong_VONIC_CheckBox.TabIndex = 150
        Me.Spec_CarGong_VONIC_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_CarGong_COB_CheckBox
        '
        Me.Spec_CarGong_COB_CheckBox.AutoSize = True
        Me.Spec_CarGong_COB_CheckBox.Location = New System.Drawing.Point(203, 89)
        Me.Spec_CarGong_COB_CheckBox.Name = "Spec_CarGong_COB_CheckBox"
        Me.Spec_CarGong_COB_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Spec_CarGong_COB_CheckBox.TabIndex = 149
        Me.Spec_CarGong_COB_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_CarGong_TopBtm_CheckBox
        '
        Me.Spec_CarGong_TopBtm_CheckBox.AutoSize = True
        Me.Spec_CarGong_TopBtm_CheckBox.Location = New System.Drawing.Point(203, 51)
        Me.Spec_CarGong_TopBtm_CheckBox.Name = "Spec_CarGong_TopBtm_CheckBox"
        Me.Spec_CarGong_TopBtm_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Spec_CarGong_TopBtm_CheckBox.TabIndex = 148
        Me.Spec_CarGong_TopBtm_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_CarGong_Top_CheckBox
        '
        Me.Spec_CarGong_Top_CheckBox.AutoSize = True
        Me.Spec_CarGong_Top_CheckBox.Location = New System.Drawing.Point(203, 13)
        Me.Spec_CarGong_Top_CheckBox.Name = "Spec_CarGong_Top_CheckBox"
        Me.Spec_CarGong_Top_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Spec_CarGong_Top_CheckBox.TabIndex = 147
        Me.Spec_CarGong_Top_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_CarGong_Top_Only_CheckBox
        '
        Me.Spec_CarGong_Top_Only_CheckBox.AutoSize = True
        Me.Spec_CarGong_Top_Only_CheckBox.Enabled = False
        Me.Spec_CarGong_Top_Only_CheckBox.Location = New System.Drawing.Point(377, 10)
        Me.Spec_CarGong_Top_Only_CheckBox.Name = "Spec_CarGong_Top_Only_CheckBox"
        Me.Spec_CarGong_Top_Only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_CarGong_Top_Only_CheckBox.TabIndex = 144
        Me.Spec_CarGong_Top_Only_CheckBox.Text = "Only"
        Me.Spec_CarGong_Top_Only_CheckBox.UseVisualStyleBackColor = True
        '
        'Label223
        '
        Me.Label223.AutoSize = True
        Me.Label223.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label223.Location = New System.Drawing.Point(544, 12)
        Me.Label223.Name = "Label223"
        Me.Label223.Size = New System.Drawing.Size(32, 16)
        Me.Label223.TabIndex = 146
        Me.Label223.Text = "號機"
        '
        'Spec_CarGong_Top_Only_TextBox
        '
        Me.Spec_CarGong_Top_Only_TextBox.Enabled = False
        Me.Spec_CarGong_Top_Only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_Top_Only_TextBox.Location = New System.Drawing.Point(432, 9)
        Me.Spec_CarGong_Top_Only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_CarGong_Top_Only_TextBox.MaxLength = 50
        Me.Spec_CarGong_Top_Only_TextBox.Name = "Spec_CarGong_Top_Only_TextBox"
        Me.Spec_CarGong_Top_Only_TextBox.Size = New System.Drawing.Size(106, 23)
        Me.Spec_CarGong_Top_Only_TextBox.TabIndex = 145
        '
        'Spec_CarGong_Label
        '
        Me.Spec_CarGong_Label.AutoSize = True
        Me.Spec_CarGong_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CarGong_Label.Location = New System.Drawing.Point(27, 12)
        Me.Spec_CarGong_Label.Name = "Spec_CarGong_Label"
        Me.Spec_CarGong_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_CarGong_Label.TabIndex = 70
        Me.Spec_CarGong_Label.Text = "車廂上到著鈴"
        '
        'Spec_CarGong_ComboBox
        '
        Me.Spec_CarGong_ComboBox.FormattingEnabled = True
        Me.Spec_CarGong_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CarGong_ComboBox.Location = New System.Drawing.Point(147, 8)
        Me.Spec_CarGong_ComboBox.Name = "Spec_CarGong_ComboBox"
        Me.Spec_CarGong_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CarGong_ComboBox.TabIndex = 71
        '
        'Spec_CRD_Panel
        '
        Me.Spec_CRD_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDType_Label)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDType_ComboBox)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDID5_Label)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRD_Label)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRD_ComboBox)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDSpec_Label)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDSpec_ComboBox)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDCancell_Label)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDCancell_ComboBox)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDNuisance_Label)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDNuisance_ComboBox)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDReg_Label)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDReg_ComboBox)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDID4_Label)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDID4_ComboBox)
        Me.Spec_CRD_Panel.Controls.Add(Me.Spec_CRDID5_ComboBox)
        Me.Spec_CRD_Panel.Location = New System.Drawing.Point(3, 164)
        Me.Spec_CRD_Panel.Name = "Spec_CRD_Panel"
        Me.Spec_CRD_Panel.Size = New System.Drawing.Size(580, 115)
        Me.Spec_CRD_Panel.TabIndex = 190
        '
        'Spec_CRDType_Label
        '
        Me.Spec_CRDType_Label.AutoSize = True
        Me.Spec_CRDType_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CRDType_Label.Location = New System.Drawing.Point(204, 9)
        Me.Spec_CRDType_Label.Name = "Spec_CRDType_Label"
        Me.Spec_CRDType_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_CRDType_Label.TabIndex = 93
        Me.Spec_CRDType_Label.Text = "分層?"
        '
        'Spec_CRDType_ComboBox
        '
        Me.Spec_CRDType_ComboBox.FormattingEnabled = True
        Me.Spec_CRDType_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CRDType_ComboBox.Location = New System.Drawing.Point(248, 5)
        Me.Spec_CRDType_ComboBox.Name = "Spec_CRDType_ComboBox"
        Me.Spec_CRDType_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CRDType_ComboBox.TabIndex = 92
        '
        'Spec_CRDID5_Label
        '
        Me.Spec_CRDID5_Label.AutoSize = True
        Me.Spec_CRDID5_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CRDID5_Label.Location = New System.Drawing.Point(410, 84)
        Me.Spec_CRDID5_Label.Name = "Spec_CRDID5_Label"
        Me.Spec_CRDID5_Label.Size = New System.Drawing.Size(84, 16)
        Me.Spec_CRDID5_Label.TabIndex = 91
        Me.Spec_CRDID5_Label.Text = "ID : 5 >>>>>"
        '
        'Spec_CRD_Label
        '
        Me.Spec_CRD_Label.AutoSize = True
        Me.Spec_CRD_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CRD_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_CRD_Label.Name = "Spec_CRD_Label"
        Me.Spec_CRD_Label.Size = New System.Drawing.Size(44, 16)
        Me.Spec_CRD_Label.TabIndex = 78
        Me.Spec_CRD_Label.Text = "刷卡機"
        '
        'Spec_CRD_ComboBox
        '
        Me.Spec_CRD_ComboBox.FormattingEnabled = True
        Me.Spec_CRD_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CRD_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_CRD_ComboBox.Name = "Spec_CRD_ComboBox"
        Me.Spec_CRD_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CRD_ComboBox.TabIndex = 79
        '
        'Spec_CRDSpec_Label
        '
        Me.Spec_CRDSpec_Label.AutoSize = True
        Me.Spec_CRDSpec_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CRDSpec_Label.Location = New System.Drawing.Point(299, 9)
        Me.Spec_CRDSpec_Label.Name = "Spec_CRDSpec_Label"
        Me.Spec_CRDSpec_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_CRDSpec_Label.TabIndex = 80
        Me.Spec_CRDSpec_Label.Text = "仕樣 :"
        '
        'Spec_CRDSpec_ComboBox
        '
        Me.Spec_CRDSpec_ComboBox.FormattingEnabled = True
        Me.Spec_CRDSpec_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CRDSpec_ComboBox.Location = New System.Drawing.Point(351, 5)
        Me.Spec_CRDSpec_ComboBox.Name = "Spec_CRDSpec_ComboBox"
        Me.Spec_CRDSpec_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CRDSpec_ComboBox.TabIndex = 81
        '
        'Spec_CRDCancell_Label
        '
        Me.Spec_CRDCancell_Label.AutoSize = True
        Me.Spec_CRDCancell_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CRDCancell_Label.Location = New System.Drawing.Point(410, 9)
        Me.Spec_CRDCancell_Label.Name = "Spec_CRDCancell_Label"
        Me.Spec_CRDCancell_Label.Size = New System.Drawing.Size(86, 16)
        Me.Spec_CRDCancell_Label.TabIndex = 82
        Me.Spec_CRDCancell_Label.Text = "逆向呼叫無效 :"
        '
        'Spec_CRDCancell_ComboBox
        '
        Me.Spec_CRDCancell_ComboBox.FormattingEnabled = True
        Me.Spec_CRDCancell_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CRDCancell_ComboBox.Location = New System.Drawing.Point(510, 5)
        Me.Spec_CRDCancell_ComboBox.Name = "Spec_CRDCancell_ComboBox"
        Me.Spec_CRDCancell_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CRDCancell_ComboBox.TabIndex = 83
        '
        'Spec_CRDNuisance_Label
        '
        Me.Spec_CRDNuisance_Label.AutoSize = True
        Me.Spec_CRDNuisance_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CRDNuisance_Label.Location = New System.Drawing.Point(263, 84)
        Me.Spec_CRDNuisance_Label.Name = "Spec_CRDNuisance_Label"
        Me.Spec_CRDNuisance_Label.Size = New System.Drawing.Size(74, 16)
        Me.Spec_CRDNuisance_Label.TabIndex = 84
        Me.Spec_CRDNuisance_Label.Text = "防嬉戲呼叫 :"
        '
        'Spec_CRDNuisance_ComboBox
        '
        Me.Spec_CRDNuisance_ComboBox.FormattingEnabled = True
        Me.Spec_CRDNuisance_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CRDNuisance_ComboBox.Location = New System.Drawing.Point(351, 80)
        Me.Spec_CRDNuisance_ComboBox.Name = "Spec_CRDNuisance_ComboBox"
        Me.Spec_CRDNuisance_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CRDNuisance_ComboBox.TabIndex = 85
        '
        'Spec_CRDReg_Label
        '
        Me.Spec_CRDReg_Label.AutoSize = True
        Me.Spec_CRDReg_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CRDReg_Label.Location = New System.Drawing.Point(282, 46)
        Me.Spec_CRDReg_Label.Name = "Spec_CRDReg_Label"
        Me.Spec_CRDReg_Label.Size = New System.Drawing.Size(62, 16)
        Me.Spec_CRDReg_Label.TabIndex = 86
        Me.Spec_CRDReg_Label.Text = "自動登錄 :"
        '
        'Spec_CRDReg_ComboBox
        '
        Me.Spec_CRDReg_ComboBox.FormattingEnabled = True
        Me.Spec_CRDReg_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CRDReg_ComboBox.Location = New System.Drawing.Point(350, 42)
        Me.Spec_CRDReg_ComboBox.Name = "Spec_CRDReg_ComboBox"
        Me.Spec_CRDReg_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CRDReg_ComboBox.TabIndex = 87
        '
        'Spec_CRDID4_Label
        '
        Me.Spec_CRDID4_Label.AutoSize = True
        Me.Spec_CRDID4_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_CRDID4_Label.Location = New System.Drawing.Point(410, 46)
        Me.Spec_CRDID4_Label.Name = "Spec_CRDID4_Label"
        Me.Spec_CRDID4_Label.Size = New System.Drawing.Size(84, 16)
        Me.Spec_CRDID4_Label.TabIndex = 88
        Me.Spec_CRDID4_Label.Text = "ID : 4 >>>>>"
        '
        'Spec_CRDID4_ComboBox
        '
        Me.Spec_CRDID4_ComboBox.FormattingEnabled = True
        Me.Spec_CRDID4_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CRDID4_ComboBox.Location = New System.Drawing.Point(510, 43)
        Me.Spec_CRDID4_ComboBox.Name = "Spec_CRDID4_ComboBox"
        Me.Spec_CRDID4_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CRDID4_ComboBox.TabIndex = 89
        '
        'Spec_CRDID5_ComboBox
        '
        Me.Spec_CRDID5_ComboBox.FormattingEnabled = True
        Me.Spec_CRDID5_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CRDID5_ComboBox.Location = New System.Drawing.Point(510, 80)
        Me.Spec_CRDID5_ComboBox.Name = "Spec_CRDID5_ComboBox"
        Me.Spec_CRDID5_ComboBox.Size = New System.Drawing.Size(44, 24)
        Me.Spec_CRDID5_ComboBox.TabIndex = 90
        '
        'TabPage13
        '
        Me.TabPage13.Controls.Add(Me.Spec_TW_FlowLayoutPanel4)
        Me.TabPage13.Location = New System.Drawing.Point(4, 25)
        Me.TabPage13.Name = "TabPage13"
        Me.TabPage13.Size = New System.Drawing.Size(627, 469)
        Me.TabPage13.TabIndex = 4
        Me.TabPage13.Text = "Page4"
        Me.TabPage13.UseVisualStyleBackColor = True
        '
        'Spec_TW_FlowLayoutPanel4
        '
        Me.Spec_TW_FlowLayoutPanel4.AutoScroll = True
        Me.Spec_TW_FlowLayoutPanel4.Controls.Add(Me.Spec_VonicBz_Panel)
        Me.Spec_TW_FlowLayoutPanel4.Controls.Add(Me.Spec_DrHold_Panel)
        Me.Spec_TW_FlowLayoutPanel4.Controls.Add(Me.Spec_Landic_Panel)
        Me.Spec_TW_FlowLayoutPanel4.Controls.Add(Me.Spec_MFLReturn_Panel)
        Me.Spec_TW_FlowLayoutPanel4.Controls.Add(Me.Spec_Vonic_Panel)
        Me.Spec_TW_FlowLayoutPanel4.Controls.Add(Me.Spec_Emer_Panel)
        Me.Spec_TW_FlowLayoutPanel4.Enabled = False
        Me.Spec_TW_FlowLayoutPanel4.Location = New System.Drawing.Point(6, 6)
        Me.Spec_TW_FlowLayoutPanel4.Name = "Spec_TW_FlowLayoutPanel4"
        Me.Spec_TW_FlowLayoutPanel4.Size = New System.Drawing.Size(615, 457)
        Me.Spec_TW_FlowLayoutPanel4.TabIndex = 2
        '
        'Spec_VonicBz_Panel
        '
        Me.Spec_VonicBz_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_VonicBz_Panel.Controls.Add(Me.Spec_VonicBz_Label)
        Me.Spec_VonicBz_Panel.Controls.Add(Me.Spec_VonicBz_ComboBox)
        Me.Spec_VonicBz_Panel.Location = New System.Drawing.Point(3, 3)
        Me.Spec_VonicBz_Panel.Name = "Spec_VonicBz_Panel"
        Me.Spec_VonicBz_Panel.Size = New System.Drawing.Size(580, 47)
        Me.Spec_VonicBz_Panel.TabIndex = 213
        '
        'Spec_VonicBz_Label
        '
        Me.Spec_VonicBz_Label.AutoSize = True
        Me.Spec_VonicBz_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_VonicBz_Label.Location = New System.Drawing.Point(27, 16)
        Me.Spec_VonicBz_Label.Name = "Spec_VonicBz_Label"
        Me.Spec_VonicBz_Label.Size = New System.Drawing.Size(83, 16)
        Me.Spec_VonicBz_Label.TabIndex = 117
        Me.Spec_VonicBz_Label.Text = "VONIC蜂鳴器"
        '
        'Spec_VonicBz_ComboBox
        '
        Me.Spec_VonicBz_ComboBox.FormattingEnabled = True
        Me.Spec_VonicBz_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_VonicBz_ComboBox.Location = New System.Drawing.Point(147, 12)
        Me.Spec_VonicBz_ComboBox.Name = "Spec_VonicBz_ComboBox"
        Me.Spec_VonicBz_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_VonicBz_ComboBox.TabIndex = 118
        '
        'Spec_DrHold_Panel
        '
        Me.Spec_DrHold_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_DrHold_Panel.Controls.Add(Me.Spec_DrHold_Label)
        Me.Spec_DrHold_Panel.Controls.Add(Me.Spec_DrHold_ComboBox)
        Me.Spec_DrHold_Panel.Location = New System.Drawing.Point(3, 56)
        Me.Spec_DrHold_Panel.Name = "Spec_DrHold_Panel"
        Me.Spec_DrHold_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_DrHold_Panel.TabIndex = 209
        '
        'Spec_DrHold_Label
        '
        Me.Spec_DrHold_Label.AutoSize = True
        Me.Spec_DrHold_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_DrHold_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_DrHold_Label.Name = "Spec_DrHold_Label"
        Me.Spec_DrHold_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_DrHold_Label.TabIndex = 76
        Me.Spec_DrHold_Label.Text = "開門延長按鈕"
        '
        'Spec_DrHold_ComboBox
        '
        Me.Spec_DrHold_ComboBox.FormattingEnabled = True
        Me.Spec_DrHold_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_DrHold_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_DrHold_ComboBox.Name = "Spec_DrHold_ComboBox"
        Me.Spec_DrHold_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_DrHold_ComboBox.TabIndex = 77
        '
        'Spec_Landic_Panel
        '
        Me.Spec_Landic_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Landic_Panel.Controls.Add(Me.Spec_Landic_Label)
        Me.Spec_Landic_Panel.Controls.Add(Me.Spec_Landic_ComboBox)
        Me.Spec_Landic_Panel.Location = New System.Drawing.Point(3, 98)
        Me.Spec_Landic_Panel.Name = "Spec_Landic_Panel"
        Me.Spec_Landic_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_Landic_Panel.TabIndex = 210
        '
        'Spec_Landic_Label
        '
        Me.Spec_Landic_Label.AutoSize = True
        Me.Spec_Landic_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Landic_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Landic_Label.Name = "Spec_Landic_Label"
        Me.Spec_Landic_Label.Size = New System.Drawing.Size(52, 16)
        Me.Spec_Landic_Label.TabIndex = 113
        Me.Spec_Landic_Label.Text = "LANDIC"
        '
        'Spec_Landic_ComboBox
        '
        Me.Spec_Landic_ComboBox.FormattingEnabled = True
        Me.Spec_Landic_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Landic_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_Landic_ComboBox.Name = "Spec_Landic_ComboBox"
        Me.Spec_Landic_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Landic_ComboBox.TabIndex = 114
        '
        'Spec_MFLReturn_Panel
        '
        Me.Spec_MFLReturn_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_MFLReturn_Panel.Controls.Add(Me.Spec_MFLReturn_Label)
        Me.Spec_MFLReturn_Panel.Controls.Add(Me.Spec_MFLReturn_FL_TextBox)
        Me.Spec_MFLReturn_Panel.Controls.Add(Me.Spec_MFLReturn_ComboBox)
        Me.Spec_MFLReturn_Panel.Controls.Add(Me.Spec_MFLReturn_FL_Label)
        Me.Spec_MFLReturn_Panel.Location = New System.Drawing.Point(3, 140)
        Me.Spec_MFLReturn_Panel.Name = "Spec_MFLReturn_Panel"
        Me.Spec_MFLReturn_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_MFLReturn_Panel.TabIndex = 211
        '
        'Spec_MFLReturn_Label
        '
        Me.Spec_MFLReturn_Label.AutoSize = True
        Me.Spec_MFLReturn_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_MFLReturn_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_MFLReturn_Label.Name = "Spec_MFLReturn_Label"
        Me.Spec_MFLReturn_Label.Size = New System.Drawing.Size(68, 16)
        Me.Spec_MFLReturn_Label.TabIndex = 115
        Me.Spec_MFLReturn_Label.Text = "基準階復歸"
        '
        'Spec_MFLReturn_FL_TextBox
        '
        Me.Spec_MFLReturn_FL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_MFLReturn_FL_TextBox.Location = New System.Drawing.Point(271, 6)
        Me.Spec_MFLReturn_FL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_MFLReturn_FL_TextBox.MaxLength = 50
        Me.Spec_MFLReturn_FL_TextBox.Name = "Spec_MFLReturn_FL_TextBox"
        Me.Spec_MFLReturn_FL_TextBox.Size = New System.Drawing.Size(61, 23)
        Me.Spec_MFLReturn_FL_TextBox.TabIndex = 113
        '
        'Spec_MFLReturn_ComboBox
        '
        Me.Spec_MFLReturn_ComboBox.FormattingEnabled = True
        Me.Spec_MFLReturn_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_MFLReturn_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_MFLReturn_ComboBox.Name = "Spec_MFLReturn_ComboBox"
        Me.Spec_MFLReturn_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_MFLReturn_ComboBox.TabIndex = 116
        '
        'Spec_MFLReturn_FL_Label
        '
        Me.Spec_MFLReturn_FL_Label.AutoSize = True
        Me.Spec_MFLReturn_FL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_MFLReturn_FL_Label.Location = New System.Drawing.Point(203, 9)
        Me.Spec_MFLReturn_FL_Label.Name = "Spec_MFLReturn_FL_Label"
        Me.Spec_MFLReturn_FL_Label.Size = New System.Drawing.Size(50, 16)
        Me.Spec_MFLReturn_FL_Label.TabIndex = 112
        Me.Spec_MFLReturn_FL_Label.Text = "基準階 :"
        '
        'Spec_Vonic_Panel
        '
        Me.Spec_Vonic_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Vonic_Panel.Controls.Add(Me.Spec_Vonic_standard_Label)
        Me.Spec_Vonic_Panel.Controls.Add(Me.Spec_Vonic_standard_ComboBox)
        Me.Spec_Vonic_Panel.Controls.Add(Me.Spec_Vonic_Label)
        Me.Spec_Vonic_Panel.Controls.Add(Me.Spec_Vonic_ComboBox)
        Me.Spec_Vonic_Panel.Location = New System.Drawing.Point(3, 182)
        Me.Spec_Vonic_Panel.Name = "Spec_Vonic_Panel"
        Me.Spec_Vonic_Panel.Size = New System.Drawing.Size(580, 47)
        Me.Spec_Vonic_Panel.TabIndex = 212
        '
        'Spec_Vonic_standard_Label
        '
        Me.Spec_Vonic_standard_Label.AutoSize = True
        Me.Spec_Vonic_standard_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Vonic_standard_Label.Location = New System.Drawing.Point(204, 16)
        Me.Spec_Vonic_standard_Label.Name = "Spec_Vonic_standard_Label"
        Me.Spec_Vonic_standard_Label.Size = New System.Drawing.Size(35, 16)
        Me.Spec_Vonic_standard_Label.TabIndex = 119
        Me.Spec_Vonic_standard_Label.Text = "標準:"
        '
        'Spec_Vonic_standard_ComboBox
        '
        Me.Spec_Vonic_standard_ComboBox.FormattingEnabled = True
        Me.Spec_Vonic_standard_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Vonic_standard_ComboBox.Location = New System.Drawing.Point(271, 12)
        Me.Spec_Vonic_standard_ComboBox.Name = "Spec_Vonic_standard_ComboBox"
        Me.Spec_Vonic_standard_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Vonic_standard_ComboBox.TabIndex = 120
        '
        'Spec_Vonic_Label
        '
        Me.Spec_Vonic_Label.AutoSize = True
        Me.Spec_Vonic_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Vonic_Label.Location = New System.Drawing.Point(27, 8)
        Me.Spec_Vonic_Label.Name = "Spec_Vonic_Label"
        Me.Spec_Vonic_Label.Size = New System.Drawing.Size(68, 32)
        Me.Spec_Vonic_Label.TabIndex = 117
        Me.Spec_Vonic_Label.Text = "語音撥放器" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "VONIC" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Spec_Vonic_ComboBox
        '
        Me.Spec_Vonic_ComboBox.FormattingEnabled = True
        Me.Spec_Vonic_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Vonic_ComboBox.Location = New System.Drawing.Point(147, 12)
        Me.Spec_Vonic_ComboBox.Name = "Spec_Vonic_ComboBox"
        Me.Spec_Vonic_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Vonic_ComboBox.TabIndex = 118
        '
        'Spec_Emer_Panel
        '
        Me.Spec_Emer_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerNum_NumericUpDown)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerCapacity_Label)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerSignal_Label)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerAddress_ComboBox)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerInput_ComboBox)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerAddress_Label)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_emerGroup_TabControl)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerNum_Label)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_Emer_Label)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_Emer_ComboBox)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerInput_Label)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerCapacity_TextBox)
        Me.Spec_Emer_Panel.Controls.Add(Me.Spec_EmerSignal_ComboBox)
        Me.Spec_Emer_Panel.Location = New System.Drawing.Point(3, 235)
        Me.Spec_Emer_Panel.Name = "Spec_Emer_Panel"
        Me.Spec_Emer_Panel.Size = New System.Drawing.Size(580, 215)
        Me.Spec_Emer_Panel.TabIndex = 214
        '
        'Spec_EmerNum_NumericUpDown
        '
        Me.Spec_EmerNum_NumericUpDown.Location = New System.Drawing.Point(146, 43)
        Me.Spec_EmerNum_NumericUpDown.Name = "Spec_EmerNum_NumericUpDown"
        Me.Spec_EmerNum_NumericUpDown.ReadOnly = True
        Me.Spec_EmerNum_NumericUpDown.Size = New System.Drawing.Size(46, 23)
        Me.Spec_EmerNum_NumericUpDown.TabIndex = 123
        '
        'Spec_EmerCapacity_Label
        '
        Me.Spec_EmerCapacity_Label.AutoSize = True
        Me.Spec_EmerCapacity_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_EmerCapacity_Label.Location = New System.Drawing.Point(22, 116)
        Me.Spec_EmerCapacity_Label.Name = "Spec_EmerCapacity_Label"
        Me.Spec_EmerCapacity_Label.Size = New System.Drawing.Size(99, 16)
        Me.Spec_EmerCapacity_Label.TabIndex = 122
        Me.Spec_EmerCapacity_Label.Text = "緊急容量(台/群) :"
        '
        'Spec_EmerSignal_Label
        '
        Me.Spec_EmerSignal_Label.AutoSize = True
        Me.Spec_EmerSignal_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_EmerSignal_Label.Location = New System.Drawing.Point(21, 84)
        Me.Spec_EmerSignal_Label.Name = "Spec_EmerSignal_Label"
        Me.Spec_EmerSignal_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_EmerSignal_Label.TabIndex = 120
        Me.Spec_EmerSignal_Label.Text = "訊號 :"
        '
        'Spec_EmerAddress_ComboBox
        '
        Me.Spec_EmerAddress_ComboBox.FormattingEnabled = True
        Me.Spec_EmerAddress_ComboBox.Location = New System.Drawing.Point(87, 178)
        Me.Spec_EmerAddress_ComboBox.Name = "Spec_EmerAddress_ComboBox"
        Me.Spec_EmerAddress_ComboBox.Size = New System.Drawing.Size(106, 24)
        Me.Spec_EmerAddress_ComboBox.TabIndex = 119
        '
        'Spec_EmerInput_ComboBox
        '
        Me.Spec_EmerInput_ComboBox.FormattingEnabled = True
        Me.Spec_EmerInput_ComboBox.Location = New System.Drawing.Point(86, 145)
        Me.Spec_EmerInput_ComboBox.Name = "Spec_EmerInput_ComboBox"
        Me.Spec_EmerInput_ComboBox.Size = New System.Drawing.Size(106, 24)
        Me.Spec_EmerInput_ComboBox.TabIndex = 118
        '
        'Spec_EmerAddress_Label
        '
        Me.Spec_EmerAddress_Label.AutoSize = True
        Me.Spec_EmerAddress_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_EmerAddress_Label.Location = New System.Drawing.Point(22, 182)
        Me.Spec_EmerAddress_Label.Name = "Spec_EmerAddress_Label"
        Me.Spec_EmerAddress_Label.Size = New System.Drawing.Size(59, 16)
        Me.Spec_EmerAddress_Label.TabIndex = 117
        Me.Spec_EmerAddress_Label.Text = "Address :"
        '
        'Spec_emerGroup_TabControl
        '
        Me.Spec_emerGroup_TabControl.Controls.Add(Me.TabPage2)
        Me.Spec_emerGroup_TabControl.Location = New System.Drawing.Point(199, 5)
        Me.Spec_emerGroup_TabControl.Name = "Spec_emerGroup_TabControl"
        Me.Spec_emerGroup_TabControl.SelectedIndex = 0
        Me.Spec_emerGroup_TabControl.Size = New System.Drawing.Size(375, 197)
        Me.Spec_emerGroup_TabControl.TabIndex = 116
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(367, 168)
        Me.TabPage2.TabIndex = 0
        Me.TabPage2.Text = "A群"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Spec_EmerNum_Label
        '
        Me.Spec_EmerNum_Label.AutoSize = True
        Me.Spec_EmerNum_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_EmerNum_Label.Location = New System.Drawing.Point(106, 46)
        Me.Spec_EmerNum_Label.Name = "Spec_EmerNum_Label"
        Me.Spec_EmerNum_Label.Size = New System.Drawing.Size(35, 16)
        Me.Spec_EmerNum_Label.TabIndex = 113
        Me.Spec_EmerNum_Label.Text = "群數:"
        '
        'Spec_Emer_Label
        '
        Me.Spec_Emer_Label.AutoSize = True
        Me.Spec_Emer_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Emer_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Emer_Label.Name = "Spec_Emer_Label"
        Me.Spec_Emer_Label.Size = New System.Drawing.Size(44, 16)
        Me.Spec_Emer_Label.TabIndex = 91
        Me.Spec_Emer_Label.Text = "自家發"
        '
        'Spec_Emer_ComboBox
        '
        Me.Spec_Emer_ComboBox.FormattingEnabled = True
        Me.Spec_Emer_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Emer_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_Emer_ComboBox.Name = "Spec_Emer_ComboBox"
        Me.Spec_Emer_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Emer_ComboBox.TabIndex = 92
        '
        'Spec_EmerInput_Label
        '
        Me.Spec_EmerInput_Label.AutoSize = True
        Me.Spec_EmerInput_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_EmerInput_Label.Location = New System.Drawing.Point(21, 149)
        Me.Spec_EmerInput_Label.Name = "Spec_EmerInput_Label"
        Me.Spec_EmerInput_Label.Size = New System.Drawing.Size(50, 16)
        Me.Spec_EmerInput_Label.TabIndex = 109
        Me.Spec_EmerInput_Label.Text = "入力點 :"
        '
        'Spec_EmerCapacity_TextBox
        '
        Me.Spec_EmerCapacity_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_EmerCapacity_TextBox.Location = New System.Drawing.Point(157, 113)
        Me.Spec_EmerCapacity_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_EmerCapacity_TextBox.MaxLength = 50
        Me.Spec_EmerCapacity_TextBox.Name = "Spec_EmerCapacity_TextBox"
        Me.Spec_EmerCapacity_TextBox.Size = New System.Drawing.Size(35, 23)
        Me.Spec_EmerCapacity_TextBox.TabIndex = 107
        Me.Spec_EmerCapacity_TextBox.Text = "1"
        '
        'Spec_EmerSignal_ComboBox
        '
        Me.Spec_EmerSignal_ComboBox.FormattingEnabled = True
        Me.Spec_EmerSignal_ComboBox.Location = New System.Drawing.Point(109, 80)
        Me.Spec_EmerSignal_ComboBox.Name = "Spec_EmerSignal_ComboBox"
        Me.Spec_EmerSignal_ComboBox.Size = New System.Drawing.Size(84, 24)
        Me.Spec_EmerSignal_ComboBox.TabIndex = 104
        '
        'TabPage14
        '
        Me.TabPage14.Controls.Add(Me.Spec_TW_FlowLayoutPanel5)
        Me.TabPage14.Location = New System.Drawing.Point(4, 25)
        Me.TabPage14.Name = "TabPage14"
        Me.TabPage14.Size = New System.Drawing.Size(627, 469)
        Me.TabPage14.TabIndex = 5
        Me.TabPage14.Text = "Page5"
        Me.TabPage14.UseVisualStyleBackColor = True
        '
        'Spec_TW_FlowLayoutPanel5
        '
        Me.Spec_TW_FlowLayoutPanel5.AutoScroll = True
        Me.Spec_TW_FlowLayoutPanel5.Controls.Add(Me.Spec_Elvic_Panel)
        Me.Spec_TW_FlowLayoutPanel5.Controls.Add(Me.Spec_WCOB_Panel)
        Me.Spec_TW_FlowLayoutPanel5.Enabled = False
        Me.Spec_TW_FlowLayoutPanel5.Location = New System.Drawing.Point(6, 6)
        Me.Spec_TW_FlowLayoutPanel5.Name = "Spec_TW_FlowLayoutPanel5"
        Me.Spec_TW_FlowLayoutPanel5.Size = New System.Drawing.Size(615, 457)
        Me.Spec_TW_FlowLayoutPanel5.TabIndex = 3
        '
        'Spec_Elvic_Panel
        '
        Me.Spec_Elvic_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Elvic_Panel.Controls.Add(Me.Label9)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Label)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_ComboBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Parking_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_VIP_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Label202)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Indep_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_FloorLockOut_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Express_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_ReturnFL_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Traffic_Peak_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_MainFL_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Label203)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_FloorLockOut_GR_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Zoning_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_CarCall_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Traffic_Peak_ComboBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Fire_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_Wavic_CheckBox)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Label204)
        Me.Spec_Elvic_Panel.Controls.Add(Me.Spec_Elvic_CRD_CheckBox)
        Me.Spec_Elvic_Panel.Location = New System.Drawing.Point(3, 3)
        Me.Spec_Elvic_Panel.Name = "Spec_Elvic_Panel"
        Me.Spec_Elvic_Panel.Size = New System.Drawing.Size(580, 341)
        Me.Spec_Elvic_Panel.TabIndex = 195
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label9.Location = New System.Drawing.Point(109, 35)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(87, 16)
        Me.Label9.TabIndex = 145
        Me.Label9.Text = "PARKING OPE"
        Me.Label9.Visible = False
        '
        'Spec_Elvic_Label
        '
        Me.Spec_Elvic_Label.AutoSize = True
        Me.Spec_Elvic_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Elvic_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Elvic_Label.Name = "Spec_Elvic_Label"
        Me.Spec_Elvic_Label.Size = New System.Drawing.Size(40, 16)
        Me.Spec_Elvic_Label.TabIndex = 119
        Me.Spec_Elvic_Label.Text = "ELVIC"
        '
        'Spec_Elvic_ComboBox
        '
        Me.Spec_Elvic_ComboBox.FormattingEnabled = True
        Me.Spec_Elvic_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Elvic_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_Elvic_ComboBox.Name = "Spec_Elvic_ComboBox"
        Me.Spec_Elvic_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Elvic_ComboBox.TabIndex = 120
        '
        'Spec_Elvic_Parking_CheckBox
        '
        Me.Spec_Elvic_Parking_CheckBox.AutoSize = True
        Me.Spec_Elvic_Parking_CheckBox.Location = New System.Drawing.Point(202, 32)
        Me.Spec_Elvic_Parking_CheckBox.Name = "Spec_Elvic_Parking_CheckBox"
        Me.Spec_Elvic_Parking_CheckBox.Size = New System.Drawing.Size(106, 20)
        Me.Spec_Elvic_Parking_CheckBox.TabIndex = 121
        Me.Spec_Elvic_Parking_CheckBox.Text = "PARKING OPE"
        Me.Spec_Elvic_Parking_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_VIP_CheckBox
        '
        Me.Spec_Elvic_VIP_CheckBox.AutoSize = True
        Me.Spec_Elvic_VIP_CheckBox.Location = New System.Drawing.Point(202, 59)
        Me.Spec_Elvic_VIP_CheckBox.Name = "Spec_Elvic_VIP_CheckBox"
        Me.Spec_Elvic_VIP_CheckBox.Size = New System.Drawing.Size(72, 20)
        Me.Spec_Elvic_VIP_CheckBox.TabIndex = 122
        Me.Spec_Elvic_VIP_CheckBox.Text = "VIP OPE"
        Me.Spec_Elvic_VIP_CheckBox.UseVisualStyleBackColor = True
        '
        'Label202
        '
        Me.Label202.AutoSize = True
        Me.Label202.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label202.Location = New System.Drawing.Point(202, 9)
        Me.Label202.Name = "Label202"
        Me.Label202.Size = New System.Drawing.Size(162, 16)
        Me.Label202.TabIndex = 123
        Me.Label202.Text = "1.COMMAND(ELEVATOR) :"
        '
        'Spec_Elvic_Indep_CheckBox
        '
        Me.Spec_Elvic_Indep_CheckBox.AutoSize = True
        Me.Spec_Elvic_Indep_CheckBox.Location = New System.Drawing.Point(202, 86)
        Me.Spec_Elvic_Indep_CheckBox.Name = "Spec_Elvic_Indep_CheckBox"
        Me.Spec_Elvic_Indep_CheckBox.Size = New System.Drawing.Size(140, 20)
        Me.Spec_Elvic_Indep_CheckBox.TabIndex = 124
        Me.Spec_Elvic_Indep_CheckBox.Text = "INDEPENDENT OPE"
        Me.Spec_Elvic_Indep_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_FloorLockOut_CheckBox
        '
        Me.Spec_Elvic_FloorLockOut_CheckBox.AutoSize = True
        Me.Spec_Elvic_FloorLockOut_CheckBox.Location = New System.Drawing.Point(359, 31)
        Me.Spec_Elvic_FloorLockOut_CheckBox.Name = "Spec_Elvic_FloorLockOut_CheckBox"
        Me.Spec_Elvic_FloorLockOut_CheckBox.Size = New System.Drawing.Size(130, 20)
        Me.Spec_Elvic_FloorLockOut_CheckBox.TabIndex = 125
        Me.Spec_Elvic_FloorLockOut_CheckBox.Text = "FLOOR LOCK OUT"
        Me.Spec_Elvic_FloorLockOut_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_Express_CheckBox
        '
        Me.Spec_Elvic_Express_CheckBox.AutoSize = True
        Me.Spec_Elvic_Express_CheckBox.Location = New System.Drawing.Point(359, 58)
        Me.Spec_Elvic_Express_CheckBox.Name = "Spec_Elvic_Express_CheckBox"
        Me.Spec_Elvic_Express_CheckBox.Size = New System.Drawing.Size(129, 20)
        Me.Spec_Elvic_Express_CheckBox.TabIndex = 126
        Me.Spec_Elvic_Express_CheckBox.Text = "EXPRESS SERVICE"
        Me.Spec_Elvic_Express_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_ReturnFL_CheckBox
        '
        Me.Spec_Elvic_ReturnFL_CheckBox.AutoSize = True
        Me.Spec_Elvic_ReturnFL_CheckBox.Location = New System.Drawing.Point(359, 85)
        Me.Spec_Elvic_ReturnFL_CheckBox.Name = "Spec_Elvic_ReturnFL_CheckBox"
        Me.Spec_Elvic_ReturnFL_CheckBox.Size = New System.Drawing.Size(218, 20)
        Me.Spec_Elvic_ReturnFL_CheckBox.TabIndex = 127
        Me.Spec_Elvic_ReturnFL_CheckBox.Text = "RETURN TO DESIGNATED FLOOR"
        Me.Spec_Elvic_ReturnFL_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_Traffic_Peak_CheckBox
        '
        Me.Spec_Elvic_Traffic_Peak_CheckBox.AutoSize = True
        Me.Spec_Elvic_Traffic_Peak_CheckBox.Location = New System.Drawing.Point(202, 142)
        Me.Spec_Elvic_Traffic_Peak_CheckBox.Name = "Spec_Elvic_Traffic_Peak_CheckBox"
        Me.Spec_Elvic_Traffic_Peak_CheckBox.Size = New System.Drawing.Size(193, 20)
        Me.Spec_Elvic_Traffic_Peak_CheckBox.TabIndex = 128
        Me.Spec_Elvic_Traffic_Peak_CheckBox.Text = "CHANGE TRAFFIC PATTERN : "
        Me.Spec_Elvic_Traffic_Peak_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_MainFL_CheckBox
        '
        Me.Spec_Elvic_MainFL_CheckBox.AutoSize = True
        Me.Spec_Elvic_MainFL_CheckBox.Location = New System.Drawing.Point(202, 169)
        Me.Spec_Elvic_MainFL_CheckBox.Name = "Spec_Elvic_MainFL_CheckBox"
        Me.Spec_Elvic_MainFL_CheckBox.Size = New System.Drawing.Size(157, 20)
        Me.Spec_Elvic_MainFL_CheckBox.TabIndex = 129
        Me.Spec_Elvic_MainFL_CheckBox.Text = "CHANGE MAIN FLOOR"
        Me.Spec_Elvic_MainFL_CheckBox.UseVisualStyleBackColor = True
        '
        'Label203
        '
        Me.Label203.AutoSize = True
        Me.Label203.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label203.Location = New System.Drawing.Point(202, 119)
        Me.Label203.Name = "Label203"
        Me.Label203.Size = New System.Drawing.Size(144, 16)
        Me.Label203.TabIndex = 130
        Me.Label203.Text = "2.COMMAND(GROUP) :"
        '
        'Spec_Elvic_FloorLockOut_GR_CheckBox
        '
        Me.Spec_Elvic_FloorLockOut_GR_CheckBox.AutoSize = True
        Me.Spec_Elvic_FloorLockOut_GR_CheckBox.Location = New System.Drawing.Point(202, 196)
        Me.Spec_Elvic_FloorLockOut_GR_CheckBox.Name = "Spec_Elvic_FloorLockOut_GR_CheckBox"
        Me.Spec_Elvic_FloorLockOut_GR_CheckBox.Size = New System.Drawing.Size(130, 20)
        Me.Spec_Elvic_FloorLockOut_GR_CheckBox.TabIndex = 131
        Me.Spec_Elvic_FloorLockOut_GR_CheckBox.Text = "FLOOR LOCK OUT"
        Me.Spec_Elvic_FloorLockOut_GR_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_Zoning_CheckBox
        '
        Me.Spec_Elvic_Zoning_CheckBox.AutoSize = True
        Me.Spec_Elvic_Zoning_CheckBox.Location = New System.Drawing.Point(399, 169)
        Me.Spec_Elvic_Zoning_CheckBox.Name = "Spec_Elvic_Zoning_CheckBox"
        Me.Spec_Elvic_Zoning_CheckBox.Size = New System.Drawing.Size(184, 20)
        Me.Spec_Elvic_Zoning_CheckBox.TabIndex = 133
        Me.Spec_Elvic_Zoning_CheckBox.Text = "ZONING FOR EXPRESS OPE"
        Me.Spec_Elvic_Zoning_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_CarCall_CheckBox
        '
        Me.Spec_Elvic_CarCall_CheckBox.AutoSize = True
        Me.Spec_Elvic_CarCall_CheckBox.Location = New System.Drawing.Point(399, 196)
        Me.Spec_Elvic_CarCall_CheckBox.Name = "Spec_Elvic_CarCall_CheckBox"
        Me.Spec_Elvic_CarCall_CheckBox.Size = New System.Drawing.Size(164, 20)
        Me.Spec_Elvic_CarCall_CheckBox.TabIndex = 134
        Me.Spec_Elvic_CarCall_CheckBox.Text = "CAR CALL DISCONNECT"
        Me.Spec_Elvic_CarCall_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_Traffic_Peak_ComboBox
        '
        Me.Spec_Elvic_Traffic_Peak_ComboBox.FormattingEnabled = True
        Me.Spec_Elvic_Traffic_Peak_ComboBox.Items.AddRange(New Object() {"UP PEAK", "DOWN PEAK", "LUNCH TIME"})
        Me.Spec_Elvic_Traffic_Peak_ComboBox.Location = New System.Drawing.Point(418, 140)
        Me.Spec_Elvic_Traffic_Peak_ComboBox.Name = "Spec_Elvic_Traffic_Peak_ComboBox"
        Me.Spec_Elvic_Traffic_Peak_ComboBox.Size = New System.Drawing.Size(88, 24)
        Me.Spec_Elvic_Traffic_Peak_ComboBox.TabIndex = 135
        '
        'Spec_Elvic_Fire_CheckBox
        '
        Me.Spec_Elvic_Fire_CheckBox.AutoSize = True
        Me.Spec_Elvic_Fire_CheckBox.Location = New System.Drawing.Point(202, 256)
        Me.Spec_Elvic_Fire_CheckBox.Name = "Spec_Elvic_Fire_CheckBox"
        Me.Spec_Elvic_Fire_CheckBox.Size = New System.Drawing.Size(153, 20)
        Me.Spec_Elvic_Fire_CheckBox.TabIndex = 136
        Me.Spec_Elvic_Fire_CheckBox.Text = "FIRE OPE. COMMAND"
        Me.Spec_Elvic_Fire_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_Elvic_Wavic_CheckBox
        '
        Me.Spec_Elvic_Wavic_CheckBox.AutoSize = True
        Me.Spec_Elvic_Wavic_CheckBox.Location = New System.Drawing.Point(202, 283)
        Me.Spec_Elvic_Wavic_CheckBox.Name = "Spec_Elvic_Wavic_CheckBox"
        Me.Spec_Elvic_Wavic_CheckBox.Size = New System.Drawing.Size(168, 20)
        Me.Spec_Elvic_Wavic_CheckBox.TabIndex = 137
        Me.Spec_Elvic_Wavic_CheckBox.Text = "WAVIC OPE. COMMAND"
        Me.Spec_Elvic_Wavic_CheckBox.UseVisualStyleBackColor = True
        '
        'Label204
        '
        Me.Label204.AutoSize = True
        Me.Label204.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label204.Location = New System.Drawing.Point(202, 233)
        Me.Label204.Name = "Label204"
        Me.Label204.Size = New System.Drawing.Size(137, 16)
        Me.Label204.TabIndex = 138
        Me.Label204.Text = "3.OTHER COMMAND :"
        '
        'Spec_Elvic_CRD_CheckBox
        '
        Me.Spec_Elvic_CRD_CheckBox.AutoSize = True
        Me.Spec_Elvic_CRD_CheckBox.Location = New System.Drawing.Point(202, 310)
        Me.Spec_Elvic_CRD_CheckBox.Name = "Spec_Elvic_CRD_CheckBox"
        Me.Spec_Elvic_CRD_CheckBox.Size = New System.Drawing.Size(182, 20)
        Me.Spec_Elvic_CRD_CheckBox.TabIndex = 139
        Me.Spec_Elvic_CRD_CheckBox.Text = "CARD READER COMMAND"
        Me.Spec_Elvic_CRD_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_WCOB_Panel
        '
        Me.Spec_WCOB_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WSCOB_only_CheckBox)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WSCOB_only_TextBox)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Label227)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WCOB_only_CheckBox)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WCOB_only_TextBox)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Label123)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WCOB_Label)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WCOB_ComboBox)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WSCOB_Label)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WSCOB_ComboBox)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WCOB_Ring_Label)
        Me.Spec_WCOB_Panel.Controls.Add(Me.Spec_WCOB_Ring_ComboBox)
        Me.Spec_WCOB_Panel.Location = New System.Drawing.Point(3, 350)
        Me.Spec_WCOB_Panel.Name = "Spec_WCOB_Panel"
        Me.Spec_WCOB_Panel.Size = New System.Drawing.Size(580, 100)
        Me.Spec_WCOB_Panel.TabIndex = 198
        '
        'Spec_WSCOB_only_CheckBox
        '
        Me.Spec_WSCOB_only_CheckBox.AutoSize = True
        Me.Spec_WSCOB_only_CheckBox.Location = New System.Drawing.Point(350, 39)
        Me.Spec_WSCOB_only_CheckBox.Name = "Spec_WSCOB_only_CheckBox"
        Me.Spec_WSCOB_only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_WSCOB_only_CheckBox.TabIndex = 161
        Me.Spec_WSCOB_only_CheckBox.Text = "Only"
        Me.Spec_WSCOB_only_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_WSCOB_only_TextBox
        '
        Me.Spec_WSCOB_only_TextBox.Enabled = False
        Me.Spec_WSCOB_only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_WSCOB_only_TextBox.Location = New System.Drawing.Point(409, 38)
        Me.Spec_WSCOB_only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_WSCOB_only_TextBox.MaxLength = 50
        Me.Spec_WSCOB_only_TextBox.Name = "Spec_WSCOB_only_TextBox"
        Me.Spec_WSCOB_only_TextBox.Size = New System.Drawing.Size(84, 23)
        Me.Spec_WSCOB_only_TextBox.TabIndex = 160
        '
        'Label227
        '
        Me.Label227.AutoSize = True
        Me.Label227.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label227.Location = New System.Drawing.Point(499, 41)
        Me.Label227.Name = "Label227"
        Me.Label227.Size = New System.Drawing.Size(32, 16)
        Me.Label227.TabIndex = 159
        Me.Label227.Text = "號機"
        '
        'Spec_WCOB_only_CheckBox
        '
        Me.Spec_WCOB_only_CheckBox.AutoSize = True
        Me.Spec_WCOB_only_CheckBox.Location = New System.Drawing.Point(202, 7)
        Me.Spec_WCOB_only_CheckBox.Name = "Spec_WCOB_only_CheckBox"
        Me.Spec_WCOB_only_CheckBox.Size = New System.Drawing.Size(53, 20)
        Me.Spec_WCOB_only_CheckBox.TabIndex = 158
        Me.Spec_WCOB_only_CheckBox.Text = "Only"
        Me.Spec_WCOB_only_CheckBox.UseVisualStyleBackColor = True
        '
        'Spec_WCOB_only_TextBox
        '
        Me.Spec_WCOB_only_TextBox.Enabled = False
        Me.Spec_WCOB_only_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_WCOB_only_TextBox.Location = New System.Drawing.Point(261, 6)
        Me.Spec_WCOB_only_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_WCOB_only_TextBox.MaxLength = 50
        Me.Spec_WCOB_only_TextBox.Name = "Spec_WCOB_only_TextBox"
        Me.Spec_WCOB_only_TextBox.Size = New System.Drawing.Size(84, 23)
        Me.Spec_WCOB_only_TextBox.TabIndex = 155
        '
        'Label123
        '
        Me.Label123.AutoSize = True
        Me.Label123.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label123.Location = New System.Drawing.Point(351, 9)
        Me.Label123.Name = "Label123"
        Me.Label123.Size = New System.Drawing.Size(32, 16)
        Me.Label123.TabIndex = 154
        Me.Label123.Text = "號機"
        '
        'Spec_WCOB_Label
        '
        Me.Spec_WCOB_Label.AutoSize = True
        Me.Spec_WCOB_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_WCOB_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_WCOB_Label.Name = "Spec_WCOB_Label"
        Me.Spec_WCOB_Label.Size = New System.Drawing.Size(56, 16)
        Me.Spec_WCOB_Label.TabIndex = 142
        Me.Spec_WCOB_Label.Text = "殘障仕樣"
        '
        'Spec_WCOB_ComboBox
        '
        Me.Spec_WCOB_ComboBox.FormattingEnabled = True
        Me.Spec_WCOB_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WCOB_ComboBox.Location = New System.Drawing.Point(148, 5)
        Me.Spec_WCOB_ComboBox.Name = "Spec_WCOB_ComboBox"
        Me.Spec_WCOB_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WCOB_ComboBox.TabIndex = 143
        '
        'Spec_WSCOB_Label
        '
        Me.Spec_WSCOB_Label.AutoSize = True
        Me.Spec_WSCOB_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_WSCOB_Label.Location = New System.Drawing.Point(202, 41)
        Me.Spec_WSCOB_Label.Name = "Spec_WSCOB_Label"
        Me.Spec_WSCOB_Label.Size = New System.Drawing.Size(70, 16)
        Me.Spec_WSCOB_Label.TabIndex = 144
        Me.Spec_WSCOB_Label.Text = "殘障SCOB :"
        '
        'Spec_WSCOB_ComboBox
        '
        Me.Spec_WSCOB_ComboBox.FormattingEnabled = True
        Me.Spec_WSCOB_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WSCOB_ComboBox.Location = New System.Drawing.Point(276, 37)
        Me.Spec_WSCOB_ComboBox.Name = "Spec_WSCOB_ComboBox"
        Me.Spec_WSCOB_ComboBox.Size = New System.Drawing.Size(55, 24)
        Me.Spec_WSCOB_ComboBox.TabIndex = 145
        '
        'Spec_WCOB_Ring_Label
        '
        Me.Spec_WCOB_Ring_Label.AutoSize = True
        Me.Spec_WCOB_Ring_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_WCOB_Ring_Label.Location = New System.Drawing.Point(202, 71)
        Me.Spec_WCOB_Ring_Label.Name = "Spec_WCOB_Ring_Label"
        Me.Spec_WCOB_Ring_Label.Size = New System.Drawing.Size(38, 16)
        Me.Spec_WCOB_Ring_Label.TabIndex = 146
        Me.Spec_WCOB_Ring_Label.Text = "鳴動 :"
        '
        'Spec_WCOB_Ring_ComboBox
        '
        Me.Spec_WCOB_Ring_ComboBox.FormattingEnabled = True
        Me.Spec_WCOB_Ring_ComboBox.Items.AddRange(New Object() {"鳴動", "不鳴動"})
        Me.Spec_WCOB_Ring_ComboBox.Location = New System.Drawing.Point(276, 67)
        Me.Spec_WCOB_Ring_ComboBox.Name = "Spec_WCOB_Ring_ComboBox"
        Me.Spec_WCOB_Ring_ComboBox.Size = New System.Drawing.Size(55, 24)
        Me.Spec_WCOB_Ring_ComboBox.TabIndex = 147
        '
        'TabPage15
        '
        Me.TabPage15.Controls.Add(Me.Spec_TW_FlowLayoutPanel6)
        Me.TabPage15.Location = New System.Drawing.Point(4, 25)
        Me.TabPage15.Name = "TabPage15"
        Me.TabPage15.Size = New System.Drawing.Size(627, 469)
        Me.TabPage15.TabIndex = 6
        Me.TabPage15.Text = "Page6"
        Me.TabPage15.UseVisualStyleBackColor = True
        '
        'Spec_TW_FlowLayoutPanel6
        '
        Me.Spec_TW_FlowLayoutPanel6.AutoScroll = True
        Me.Spec_TW_FlowLayoutPanel6.Controls.Add(Me.Spec_HLL_Panel)
        Me.Spec_TW_FlowLayoutPanel6.Controls.Add(Me.Spec_ATT_Panel)
        Me.Spec_TW_FlowLayoutPanel6.Controls.Add(Me.Spec_Flood_Panel)
        Me.Spec_TW_FlowLayoutPanel6.Controls.Add(Me.Spec_LS1M_Panel)
        Me.Spec_TW_FlowLayoutPanel6.Controls.Add(Me.Spec_PRU_Panel)
        Me.Spec_TW_FlowLayoutPanel6.Controls.Add(Me.Spec_LoadCell_Panel)
        Me.Spec_TW_FlowLayoutPanel6.Controls.Add(Me.Spec_FrontRearDr_Panel)
        Me.Spec_TW_FlowLayoutPanel6.Controls.Add(Me.Spec_OpeSw_Panel)
        Me.Spec_TW_FlowLayoutPanel6.Enabled = False
        Me.Spec_TW_FlowLayoutPanel6.Location = New System.Drawing.Point(6, 6)
        Me.Spec_TW_FlowLayoutPanel6.Name = "Spec_TW_FlowLayoutPanel6"
        Me.Spec_TW_FlowLayoutPanel6.Size = New System.Drawing.Size(615, 457)
        Me.Spec_TW_FlowLayoutPanel6.TabIndex = 4
        '
        'Spec_HLL_Panel
        '
        Me.Spec_HLL_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_HLL_Panel.Controls.Add(Me.Spec_HLL_Label)
        Me.Spec_HLL_Panel.Controls.Add(Me.Spec_HLL_ComboBox)
        Me.Spec_HLL_Panel.Location = New System.Drawing.Point(3, 3)
        Me.Spec_HLL_Panel.Name = "Spec_HLL_Panel"
        Me.Spec_HLL_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_HLL_Panel.TabIndex = 198
        '
        'Spec_HLL_Label
        '
        Me.Spec_HLL_Label.AutoSize = True
        Me.Spec_HLL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_HLL_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_HLL_Label.Name = "Spec_HLL_Label"
        Me.Spec_HLL_Label.Size = New System.Drawing.Size(56, 16)
        Me.Spec_HLL_Label.TabIndex = 140
        Me.Spec_HLL_Label.Text = "乘場廳燈"
        '
        'Spec_HLL_ComboBox
        '
        Me.Spec_HLL_ComboBox.FormattingEnabled = True
        Me.Spec_HLL_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_HLL_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_HLL_ComboBox.Name = "Spec_HLL_ComboBox"
        Me.Spec_HLL_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_HLL_ComboBox.TabIndex = 141
        '
        'Spec_ATT_Panel
        '
        Me.Spec_ATT_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_ATT_Panel.Controls.Add(Me.Spec_ATT_Label)
        Me.Spec_ATT_Panel.Controls.Add(Me.Spec_ATT_ComboBox)
        Me.Spec_ATT_Panel.Location = New System.Drawing.Point(3, 45)
        Me.Spec_ATT_Panel.Name = "Spec_ATT_Panel"
        Me.Spec_ATT_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_ATT_Panel.TabIndex = 202
        '
        'Spec_ATT_Label
        '
        Me.Spec_ATT_Label.AutoSize = True
        Me.Spec_ATT_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_ATT_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_ATT_Label.Name = "Spec_ATT_Label"
        Me.Spec_ATT_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_ATT_Label.TabIndex = 148
        Me.Spec_ATT_Label.Text = "運轉手盤運轉"
        '
        'Spec_ATT_ComboBox
        '
        Me.Spec_ATT_ComboBox.FormattingEnabled = True
        Me.Spec_ATT_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_ATT_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_ATT_ComboBox.Name = "Spec_ATT_ComboBox"
        Me.Spec_ATT_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_ATT_ComboBox.TabIndex = 149
        '
        'Spec_Flood_Panel
        '
        Me.Spec_Flood_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Flood_Panel.Controls.Add(Me.Spec_Flood_Label)
        Me.Spec_Flood_Panel.Controls.Add(Me.Spec_Flood_ComboBox)
        Me.Spec_Flood_Panel.Controls.Add(Me.Spec_Flood_FL_TextBox)
        Me.Spec_Flood_Panel.Controls.Add(Me.Spec_Flood_FL_Label)
        Me.Spec_Flood_Panel.Location = New System.Drawing.Point(3, 87)
        Me.Spec_Flood_Panel.Name = "Spec_Flood_Panel"
        Me.Spec_Flood_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_Flood_Panel.TabIndex = 203
        '
        'Spec_Flood_Label
        '
        Me.Spec_Flood_Label.AutoSize = True
        Me.Spec_Flood_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Flood_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Flood_Label.Name = "Spec_Flood_Label"
        Me.Spec_Flood_Label.Size = New System.Drawing.Size(80, 16)
        Me.Spec_Flood_Label.TabIndex = 150
        Me.Spec_Flood_Label.Text = "浸水管制運轉"
        '
        'Spec_Flood_ComboBox
        '
        Me.Spec_Flood_ComboBox.FormattingEnabled = True
        Me.Spec_Flood_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Flood_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_Flood_ComboBox.Name = "Spec_Flood_ComboBox"
        Me.Spec_Flood_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Flood_ComboBox.TabIndex = 151
        '
        'Spec_Flood_FL_TextBox
        '
        Me.Spec_Flood_FL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Flood_FL_TextBox.Location = New System.Drawing.Point(270, 6)
        Me.Spec_Flood_FL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_Flood_FL_TextBox.MaxLength = 50
        Me.Spec_Flood_FL_TextBox.Name = "Spec_Flood_FL_TextBox"
        Me.Spec_Flood_FL_TextBox.Size = New System.Drawing.Size(125, 23)
        Me.Spec_Flood_FL_TextBox.TabIndex = 153
        Me.Spec_Flood_FL_TextBox.Text = "1FL(1th)"
        '
        'Spec_Flood_FL_Label
        '
        Me.Spec_Flood_FL_Label.AutoSize = True
        Me.Spec_Flood_FL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Flood_FL_Label.Location = New System.Drawing.Point(202, 9)
        Me.Spec_Flood_FL_Label.Name = "Spec_Flood_FL_Label"
        Me.Spec_Flood_FL_Label.Size = New System.Drawing.Size(50, 16)
        Me.Spec_Flood_FL_Label.TabIndex = 152
        Me.Spec_Flood_FL_Label.Text = "停止階 :"
        '
        'Spec_LS1M_Panel
        '
        Me.Spec_LS1M_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_LS1M_Panel.Controls.Add(Me.Spec_LS1M_Label)
        Me.Spec_LS1M_Panel.Controls.Add(Me.Spec_LS1M_ComboBox)
        Me.Spec_LS1M_Panel.Location = New System.Drawing.Point(3, 129)
        Me.Spec_LS1M_Panel.Name = "Spec_LS1M_Panel"
        Me.Spec_LS1M_Panel.Size = New System.Drawing.Size(580, 51)
        Me.Spec_LS1M_Panel.TabIndex = 204
        '
        'Spec_LS1M_Label
        '
        Me.Spec_LS1M_Label.AutoSize = True
        Me.Spec_LS1M_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_LS1M_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_LS1M_Label.Name = "Spec_LS1M_Label"
        Me.Spec_LS1M_Label.Size = New System.Drawing.Size(104, 32)
        Me.Spec_LS1M_Label.TabIndex = 154
        Me.Spec_LS1M_Label.Text = "頂部緊急停止開關" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "LS1M" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Spec_LS1M_ComboBox
        '
        Me.Spec_LS1M_ComboBox.FormattingEnabled = True
        Me.Spec_LS1M_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_LS1M_ComboBox.Location = New System.Drawing.Point(147, 13)
        Me.Spec_LS1M_ComboBox.Name = "Spec_LS1M_ComboBox"
        Me.Spec_LS1M_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_LS1M_ComboBox.TabIndex = 155
        '
        'Spec_PRU_Panel
        '
        Me.Spec_PRU_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_PRU_Panel.Controls.Add(Me.Spec_PRU_Label)
        Me.Spec_PRU_Panel.Controls.Add(Me.Spec_PRU_ComboBox)
        Me.Spec_PRU_Panel.Location = New System.Drawing.Point(3, 186)
        Me.Spec_PRU_Panel.Name = "Spec_PRU_Panel"
        Me.Spec_PRU_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_PRU_Panel.TabIndex = 205
        '
        'Spec_PRU_Label
        '
        Me.Spec_PRU_Label.AutoSize = True
        Me.Spec_PRU_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_PRU_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_PRU_Label.Name = "Spec_PRU_Label"
        Me.Spec_PRU_Label.Size = New System.Drawing.Size(88, 16)
        Me.Spec_PRU_Label.TabIndex = 156
        Me.Spec_PRU_Label.Text = "電力回升(PRU)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Spec_PRU_ComboBox
        '
        Me.Spec_PRU_ComboBox.FormattingEnabled = True
        Me.Spec_PRU_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_PRU_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_PRU_ComboBox.Name = "Spec_PRU_ComboBox"
        Me.Spec_PRU_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_PRU_ComboBox.TabIndex = 157
        '
        'Spec_LoadCell_Panel
        '
        Me.Spec_LoadCell_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_LoadCell_Panel.Controls.Add(Me.Spec_LoadCellPos_ComboBox)
        Me.Spec_LoadCell_Panel.Controls.Add(Me.Spec_LoadCellPos_Label)
        Me.Spec_LoadCell_Panel.Controls.Add(Me.Spec_LoadCell_Label)
        Me.Spec_LoadCell_Panel.Controls.Add(Me.Spec_LoadCell_ComboBox)
        Me.Spec_LoadCell_Panel.Location = New System.Drawing.Point(3, 228)
        Me.Spec_LoadCell_Panel.Name = "Spec_LoadCell_Panel"
        Me.Spec_LoadCell_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_LoadCell_Panel.TabIndex = 206
        '
        'Spec_LoadCellPos_ComboBox
        '
        Me.Spec_LoadCellPos_ComboBox.FormattingEnabled = True
        Me.Spec_LoadCellPos_ComboBox.Items.AddRange(New Object() {"CAR BTM", "MR"})
        Me.Spec_LoadCellPos_ComboBox.Location = New System.Drawing.Point(271, 5)
        Me.Spec_LoadCellPos_ComboBox.Name = "Spec_LoadCellPos_ComboBox"
        Me.Spec_LoadCellPos_ComboBox.Size = New System.Drawing.Size(61, 24)
        Me.Spec_LoadCellPos_ComboBox.TabIndex = 155
        '
        'Spec_LoadCellPos_Label
        '
        Me.Spec_LoadCellPos_Label.AutoSize = True
        Me.Spec_LoadCellPos_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_LoadCellPos_Label.Location = New System.Drawing.Point(203, 9)
        Me.Spec_LoadCellPos_Label.Name = "Spec_LoadCellPos_Label"
        Me.Spec_LoadCellPos_Label.Size = New System.Drawing.Size(50, 16)
        Me.Spec_LoadCellPos_Label.TabIndex = 154
        Me.Spec_LoadCellPos_Label.Text = "裝置在 :"
        '
        'Spec_LoadCell_Label
        '
        Me.Spec_LoadCell_Label.AutoSize = True
        Me.Spec_LoadCell_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_LoadCell_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_LoadCell_Label.Name = "Spec_LoadCell_Label"
        Me.Spec_LoadCell_Label.Size = New System.Drawing.Size(61, 16)
        Me.Spec_LoadCell_Label.TabIndex = 76
        Me.Spec_LoadCell_Label.Text = "Load Cell"
        '
        'Spec_LoadCell_ComboBox
        '
        Me.Spec_LoadCell_ComboBox.FormattingEnabled = True
        Me.Spec_LoadCell_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_LoadCell_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_LoadCell_ComboBox.Name = "Spec_LoadCell_ComboBox"
        Me.Spec_LoadCell_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_LoadCell_ComboBox.TabIndex = 77
        '
        'Spec_FrontRearDr_Panel
        '
        Me.Spec_FrontRearDr_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_FrontRearDr_Panel.Controls.Add(Me.Spec_FrontRearDr_Label)
        Me.Spec_FrontRearDr_Panel.Controls.Add(Me.Spec_FrontRearDr_ComboBox)
        Me.Spec_FrontRearDr_Panel.Location = New System.Drawing.Point(3, 270)
        Me.Spec_FrontRearDr_Panel.Name = "Spec_FrontRearDr_Panel"
        Me.Spec_FrontRearDr_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_FrontRearDr_Panel.TabIndex = 211
        '
        'Spec_FrontRearDr_Label
        '
        Me.Spec_FrontRearDr_Label.AutoSize = True
        Me.Spec_FrontRearDr_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_FrontRearDr_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_FrontRearDr_Label.Name = "Spec_FrontRearDr_Label"
        Me.Spec_FrontRearDr_Label.Size = New System.Drawing.Size(44, 16)
        Me.Spec_FrontRearDr_Label.TabIndex = 76
        Me.Spec_FrontRearDr_Label.Text = "正背門"
        '
        'Spec_FrontRearDr_ComboBox
        '
        Me.Spec_FrontRearDr_ComboBox.FormattingEnabled = True
        Me.Spec_FrontRearDr_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_FrontRearDr_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_FrontRearDr_ComboBox.Name = "Spec_FrontRearDr_ComboBox"
        Me.Spec_FrontRearDr_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_FrontRearDr_ComboBox.TabIndex = 77
        '
        'Spec_OpeSw_Panel
        '
        Me.Spec_OpeSw_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_OpeSw_Panel.Controls.Add(Me.Label7)
        Me.Spec_OpeSw_Panel.Controls.Add(Me.Spec_OpeSw_InputPos_ComboBox)
        Me.Spec_OpeSw_Panel.Controls.Add(Me.Spec_OpeSw_InputAddress_TextBox)
        Me.Spec_OpeSw_Panel.Controls.Add(Me.Spec_OpeSw_InputPos_Label)
        Me.Spec_OpeSw_Panel.Controls.Add(Me.Spec_OpeSw_DevicePos_TextBox)
        Me.Spec_OpeSw_Panel.Controls.Add(Me.Spec_OpeSw_DevicePos_Label)
        Me.Spec_OpeSw_Panel.Controls.Add(Me.Spec_OpeSw_Label)
        Me.Spec_OpeSw_Panel.Controls.Add(Me.Spec_OpeSw_ComboBox)
        Me.Spec_OpeSw_Panel.Location = New System.Drawing.Point(3, 312)
        Me.Spec_OpeSw_Panel.Name = "Spec_OpeSw_Panel"
        Me.Spec_OpeSw_Panel.Size = New System.Drawing.Size(580, 115)
        Me.Spec_OpeSw_Panel.TabIndex = 212
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label7.Location = New System.Drawing.Point(204, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(59, 16)
        Me.Label7.TabIndex = 160
        Me.Label7.Text = "Address :"
        '
        'Spec_OpeSw_InputPos_ComboBox
        '
        Me.Spec_OpeSw_InputPos_ComboBox.FormattingEnabled = True
        Me.Spec_OpeSw_InputPos_ComboBox.Items.AddRange(New Object() {"MR", "GSP"})
        Me.Spec_OpeSw_InputPos_ComboBox.Location = New System.Drawing.Point(271, 42)
        Me.Spec_OpeSw_InputPos_ComboBox.Name = "Spec_OpeSw_InputPos_ComboBox"
        Me.Spec_OpeSw_InputPos_ComboBox.Size = New System.Drawing.Size(61, 24)
        Me.Spec_OpeSw_InputPos_ComboBox.TabIndex = 159
        '
        'Spec_OpeSw_InputAddress_TextBox
        '
        Me.Spec_OpeSw_InputAddress_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_OpeSw_InputAddress_TextBox.Location = New System.Drawing.Point(271, 77)
        Me.Spec_OpeSw_InputAddress_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_OpeSw_InputAddress_TextBox.MaxLength = 50
        Me.Spec_OpeSw_InputAddress_TextBox.Name = "Spec_OpeSw_InputAddress_TextBox"
        Me.Spec_OpeSw_InputAddress_TextBox.Size = New System.Drawing.Size(125, 23)
        Me.Spec_OpeSw_InputAddress_TextBox.TabIndex = 158
        Me.Spec_OpeSw_InputAddress_TextBox.Text = "$XXXX Bit X"
        '
        'Spec_OpeSw_InputPos_Label
        '
        Me.Spec_OpeSw_InputPos_Label.AutoSize = True
        Me.Spec_OpeSw_InputPos_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_OpeSw_InputPos_Label.Location = New System.Drawing.Point(204, 46)
        Me.Spec_OpeSw_InputPos_Label.Name = "Spec_OpeSw_InputPos_Label"
        Me.Spec_OpeSw_InputPos_Label.Size = New System.Drawing.Size(50, 16)
        Me.Spec_OpeSw_InputPos_Label.TabIndex = 157
        Me.Spec_OpeSw_InputPos_Label.Text = "入力點 :"
        '
        'Spec_OpeSw_DevicePos_TextBox
        '
        Me.Spec_OpeSw_DevicePos_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_OpeSw_DevicePos_TextBox.Location = New System.Drawing.Point(270, 6)
        Me.Spec_OpeSw_DevicePos_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Spec_OpeSw_DevicePos_TextBox.MaxLength = 50
        Me.Spec_OpeSw_DevicePos_TextBox.Name = "Spec_OpeSw_DevicePos_TextBox"
        Me.Spec_OpeSw_DevicePos_TextBox.Size = New System.Drawing.Size(125, 23)
        Me.Spec_OpeSw_DevicePos_TextBox.TabIndex = 156
        Me.Spec_OpeSw_DevicePos_TextBox.Text = "運轉手盤"
        '
        'Spec_OpeSw_DevicePos_Label
        '
        Me.Spec_OpeSw_DevicePos_Label.AutoSize = True
        Me.Spec_OpeSw_DevicePos_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_OpeSw_DevicePos_Label.Location = New System.Drawing.Point(203, 9)
        Me.Spec_OpeSw_DevicePos_Label.Name = "Spec_OpeSw_DevicePos_Label"
        Me.Spec_OpeSw_DevicePos_Label.Size = New System.Drawing.Size(50, 16)
        Me.Spec_OpeSw_DevicePos_Label.TabIndex = 155
        Me.Spec_OpeSw_DevicePos_Label.Text = "裝置在 :"
        '
        'Spec_OpeSw_Label
        '
        Me.Spec_OpeSw_Label.AutoSize = True
        Me.Spec_OpeSw_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_OpeSw_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_OpeSw_Label.Name = "Spec_OpeSw_Label"
        Me.Spec_OpeSw_Label.Size = New System.Drawing.Size(68, 16)
        Me.Spec_OpeSw_Label.TabIndex = 76
        Me.Spec_OpeSw_Label.Text = "單群控切換"
        '
        'Spec_OpeSw_ComboBox
        '
        Me.Spec_OpeSw_ComboBox.FormattingEnabled = True
        Me.Spec_OpeSw_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_OpeSw_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_OpeSw_ComboBox.Name = "Spec_OpeSw_ComboBox"
        Me.Spec_OpeSw_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_OpeSw_ComboBox.TabIndex = 77
        '
        'TabPage11
        '
        Me.TabPage11.Controls.Add(Me.Spec_TW_unUse_FlowLayoutPanel)
        Me.TabPage11.Location = New System.Drawing.Point(4, 25)
        Me.TabPage11.Name = "TabPage11"
        Me.TabPage11.Size = New System.Drawing.Size(627, 469)
        Me.TabPage11.TabIndex = 2
        Me.TabPage11.Text = "不使用"
        Me.TabPage11.UseVisualStyleBackColor = True
        '
        'Spec_TW_unUse_FlowLayoutPanel
        '
        Me.Spec_TW_unUse_FlowLayoutPanel.AutoScroll = True
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Panel42)
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Panel43)
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Panel54)
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Panel66)
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Spec_WTB_Panel)
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Spec_IF79x_Panel)
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Spec_EachStop_Panel)
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Panel115)
        Me.Spec_TW_unUse_FlowLayoutPanel.Controls.Add(Me.Spec_Operation_Panel)
        Me.Spec_TW_unUse_FlowLayoutPanel.Enabled = False
        Me.Spec_TW_unUse_FlowLayoutPanel.Location = New System.Drawing.Point(6, 6)
        Me.Spec_TW_unUse_FlowLayoutPanel.Name = "Spec_TW_unUse_FlowLayoutPanel"
        Me.Spec_TW_unUse_FlowLayoutPanel.Size = New System.Drawing.Size(615, 457)
        Me.Spec_TW_unUse_FlowLayoutPanel.TabIndex = 1
        '
        'Panel42
        '
        Me.Panel42.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel42.Controls.Add(Me.Label155)
        Me.Panel42.Controls.Add(Me.Spec_CancellBehind_ComboBox)
        Me.Panel42.Location = New System.Drawing.Point(3, 3)
        Me.Panel42.Name = "Panel42"
        Me.Panel42.Size = New System.Drawing.Size(580, 36)
        Me.Panel42.TabIndex = 165
        '
        'Label155
        '
        Me.Label155.AutoSize = True
        Me.Label155.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label155.Location = New System.Drawing.Point(27, 9)
        Me.Label155.Name = "Label155"
        Me.Label155.Size = New System.Drawing.Size(56, 16)
        Me.Label155.TabIndex = 16
        Me.Label155.Text = "逆呼無效"
        '
        'Spec_CancellBehind_ComboBox
        '
        Me.Spec_CancellBehind_ComboBox.FormattingEnabled = True
        Me.Spec_CancellBehind_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CancellBehind_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_CancellBehind_ComboBox.Name = "Spec_CancellBehind_ComboBox"
        Me.Spec_CancellBehind_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CancellBehind_ComboBox.TabIndex = 32
        '
        'Panel43
        '
        Me.Panel43.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel43.Controls.Add(Me.Label156)
        Me.Panel43.Controls.Add(Me.Spec_LampChk_ComboBox)
        Me.Panel43.Location = New System.Drawing.Point(3, 45)
        Me.Panel43.Name = "Panel43"
        Me.Panel43.Size = New System.Drawing.Size(580, 36)
        Me.Panel43.TabIndex = 166
        '
        'Label156
        '
        Me.Label156.AutoSize = True
        Me.Label156.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label156.Location = New System.Drawing.Point(27, 9)
        Me.Label156.Name = "Label156"
        Me.Label156.Size = New System.Drawing.Size(68, 16)
        Me.Label156.TabIndex = 17
        Me.Label156.Text = "燈點檢模式"
        '
        'Spec_LampChk_ComboBox
        '
        Me.Spec_LampChk_ComboBox.FormattingEnabled = True
        Me.Spec_LampChk_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_LampChk_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_LampChk_ComboBox.Name = "Spec_LampChk_ComboBox"
        Me.Spec_LampChk_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_LampChk_ComboBox.TabIndex = 33
        '
        'Panel54
        '
        Me.Panel54.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel54.Controls.Add(Me.Label163)
        Me.Panel54.Controls.Add(Me.Spec_CCCancell_ComboBox)
        Me.Panel54.Location = New System.Drawing.Point(3, 87)
        Me.Panel54.Name = "Panel54"
        Me.Panel54.Size = New System.Drawing.Size(580, 36)
        Me.Panel54.TabIndex = 172
        '
        'Label163
        '
        Me.Label163.AutoSize = True
        Me.Label163.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label163.Location = New System.Drawing.Point(27, 9)
        Me.Label163.Name = "Label163"
        Me.Label163.Size = New System.Drawing.Size(80, 16)
        Me.Label163.TabIndex = 41
        Me.Label163.Text = "車廂呼叫取消"
        '
        'Spec_CCCancell_ComboBox
        '
        Me.Spec_CCCancell_ComboBox.FormattingEnabled = True
        Me.Spec_CCCancell_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_CCCancell_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_CCCancell_ComboBox.Name = "Spec_CCCancell_ComboBox"
        Me.Spec_CCCancell_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_CCCancell_ComboBox.TabIndex = 42
        '
        'Panel66
        '
        Me.Panel66.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel66.Controls.Add(Me.Spec_UCMP_ComboBox)
        Me.Panel66.Controls.Add(Me.Label169)
        Me.Panel66.Location = New System.Drawing.Point(3, 129)
        Me.Panel66.Name = "Panel66"
        Me.Panel66.Size = New System.Drawing.Size(580, 36)
        Me.Panel66.TabIndex = 178
        '
        'Spec_UCMP_ComboBox
        '
        Me.Spec_UCMP_ComboBox.FormattingEnabled = True
        Me.Spec_UCMP_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_UCMP_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_UCMP_ComboBox.Name = "Spec_UCMP_ComboBox"
        Me.Spec_UCMP_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_UCMP_ComboBox.TabIndex = 55
        '
        'Label169
        '
        Me.Label169.AutoSize = True
        Me.Label169.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label169.Location = New System.Drawing.Point(27, 9)
        Me.Label169.Name = "Label169"
        Me.Label169.Size = New System.Drawing.Size(124, 16)
        Me.Label169.TabIndex = 54
        Me.Label169.Text = "戶開行走保護(UCMP)"
        '
        'Spec_WTB_Panel
        '
        Me.Spec_WTB_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_WTB_Panel.Controls.Add(Me.Label144)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_EQMac_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label143)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_EQIND_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label142)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_Indep_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label141)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_EQ_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label140)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_Alart_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label137)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_BZSW_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label138)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_EQSW_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label139)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_PKSW_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label133)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_EmerPow_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label134)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_FO_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label135)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_Urgent_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label136)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_Normal_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label108)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_ChkSW_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label105)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_FM_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label102)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_Stop_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label98)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_Error_ComboBox)
        Me.Spec_WTB_Panel.Controls.Add(Me.Label68)
        Me.Spec_WTB_Panel.Controls.Add(Me.Spec_WTB_ComboBox)
        Me.Spec_WTB_Panel.Location = New System.Drawing.Point(3, 171)
        Me.Spec_WTB_Panel.Name = "Spec_WTB_Panel"
        Me.Spec_WTB_Panel.Size = New System.Drawing.Size(580, 194)
        Me.Spec_WTB_Panel.TabIndex = 208
        '
        'Label144
        '
        Me.Label144.AutoSize = True
        Me.Label144.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label144.Location = New System.Drawing.Point(204, 163)
        Me.Label144.Name = "Label144"
        Me.Label144.Size = New System.Drawing.Size(56, 16)
        Me.Label144.TabIndex = 155
        Me.Label144.Text = "地震強度"
        '
        'Spec_WTB_EQMac_ComboBox
        '
        Me.Spec_WTB_EQMac_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_EQMac_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_EQMac_ComboBox.Location = New System.Drawing.Point(263, 159)
        Me.Spec_WTB_EQMac_ComboBox.Name = "Spec_WTB_EQMac_ComboBox"
        Me.Spec_WTB_EQMac_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_EQMac_ComboBox.TabIndex = 156
        '
        'Label143
        '
        Me.Label143.AutoSize = True
        Me.Label143.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label143.Location = New System.Drawing.Point(445, 133)
        Me.Label143.Name = "Label143"
        Me.Label143.Size = New System.Drawing.Size(68, 16)
        Me.Label143.TabIndex = 153
        Me.Label143.Text = "地震指示器"
        '
        'Spec_WTB_EQIND_ComboBox
        '
        Me.Spec_WTB_EQIND_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_EQIND_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_EQIND_ComboBox.Location = New System.Drawing.Point(518, 129)
        Me.Spec_WTB_EQIND_ComboBox.Name = "Spec_WTB_EQIND_ComboBox"
        Me.Spec_WTB_EQIND_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_EQIND_ComboBox.TabIndex = 154
        '
        'Label142
        '
        Me.Label142.AutoSize = True
        Me.Label142.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label142.Location = New System.Drawing.Point(204, 102)
        Me.Label142.Name = "Label142"
        Me.Label142.Size = New System.Drawing.Size(44, 16)
        Me.Label142.TabIndex = 151
        Me.Label142.Text = "專用燈"
        '
        'Spec_WTB_Indep_ComboBox
        '
        Me.Spec_WTB_Indep_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_Indep_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_Indep_ComboBox.Location = New System.Drawing.Point(263, 98)
        Me.Spec_WTB_Indep_ComboBox.Name = "Spec_WTB_Indep_ComboBox"
        Me.Spec_WTB_Indep_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_Indep_ComboBox.TabIndex = 152
        '
        'Label141
        '
        Me.Label141.AutoSize = True
        Me.Label141.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label141.Location = New System.Drawing.Point(445, 71)
        Me.Label141.Name = "Label141"
        Me.Label141.Size = New System.Drawing.Size(44, 16)
        Me.Label141.TabIndex = 149
        Me.Label141.Text = "地震燈"
        '
        'Spec_WTB_EQ_ComboBox
        '
        Me.Spec_WTB_EQ_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_EQ_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_EQ_ComboBox.Location = New System.Drawing.Point(518, 67)
        Me.Spec_WTB_EQ_ComboBox.Name = "Spec_WTB_EQ_ComboBox"
        Me.Spec_WTB_EQ_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_EQ_ComboBox.TabIndex = 150
        '
        'Label140
        '
        Me.Label140.AutoSize = True
        Me.Label140.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label140.Location = New System.Drawing.Point(320, 71)
        Me.Label140.Name = "Label140"
        Me.Label140.Size = New System.Drawing.Size(44, 16)
        Me.Label140.TabIndex = 147
        Me.Label140.Text = "警示燈"
        '
        'Spec_WTB_Alart_ComboBox
        '
        Me.Spec_WTB_Alart_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_Alart_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_Alart_ComboBox.Location = New System.Drawing.Point(389, 67)
        Me.Spec_WTB_Alart_ComboBox.Name = "Spec_WTB_Alart_ComboBox"
        Me.Spec_WTB_Alart_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_Alart_ComboBox.TabIndex = 148
        '
        'Label137
        '
        Me.Label137.AutoSize = True
        Me.Label137.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label137.Location = New System.Drawing.Point(445, 102)
        Me.Label137.Name = "Label137"
        Me.Label137.Size = New System.Drawing.Size(70, 16)
        Me.Label137.TabIndex = 145
        Me.Label137.Text = "BZ解除開關"
        '
        'Spec_WTB_BZSW_ComboBox
        '
        Me.Spec_WTB_BZSW_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_BZSW_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_BZSW_ComboBox.Location = New System.Drawing.Point(518, 97)
        Me.Spec_WTB_BZSW_ComboBox.Name = "Spec_WTB_BZSW_ComboBox"
        Me.Spec_WTB_BZSW_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_BZSW_ComboBox.TabIndex = 146
        '
        'Label138
        '
        Me.Label138.AutoSize = True
        Me.Label138.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label138.Location = New System.Drawing.Point(320, 102)
        Me.Label138.Name = "Label138"
        Me.Label138.Size = New System.Drawing.Size(56, 16)
        Me.Label138.TabIndex = 143
        Me.Label138.Text = "地震開關"
        '
        'Spec_WTB_EQSW_ComboBox
        '
        Me.Spec_WTB_EQSW_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_EQSW_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_EQSW_ComboBox.Location = New System.Drawing.Point(389, 98)
        Me.Spec_WTB_EQSW_ComboBox.Name = "Spec_WTB_EQSW_ComboBox"
        Me.Spec_WTB_EQSW_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_EQSW_ComboBox.TabIndex = 144
        '
        'Label139
        '
        Me.Label139.AutoSize = True
        Me.Label139.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label139.Location = New System.Drawing.Point(320, 133)
        Me.Label139.Name = "Label139"
        Me.Label139.Size = New System.Drawing.Size(56, 16)
        Me.Label139.TabIndex = 141
        Me.Label139.Text = "停車開關"
        '
        'Spec_WTB_PKSW_ComboBox
        '
        Me.Spec_WTB_PKSW_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_PKSW_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_PKSW_ComboBox.Location = New System.Drawing.Point(389, 129)
        Me.Spec_WTB_PKSW_ComboBox.Name = "Spec_WTB_PKSW_ComboBox"
        Me.Spec_WTB_PKSW_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_PKSW_ComboBox.TabIndex = 142
        '
        'Label133
        '
        Me.Label133.AutoSize = True
        Me.Label133.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label133.Location = New System.Drawing.Point(204, 71)
        Me.Label133.Name = "Label133"
        Me.Label133.Size = New System.Drawing.Size(56, 16)
        Me.Label133.TabIndex = 139
        Me.Label133.Text = "自家發燈"
        '
        'Spec_WTB_EmerPow_ComboBox
        '
        Me.Spec_WTB_EmerPow_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_EmerPow_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_EmerPow_ComboBox.Location = New System.Drawing.Point(263, 67)
        Me.Spec_WTB_EmerPow_ComboBox.Name = "Spec_WTB_EmerPow_ComboBox"
        Me.Spec_WTB_EmerPow_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_EmerPow_ComboBox.TabIndex = 140
        '
        'Label134
        '
        Me.Label134.AutoSize = True
        Me.Label134.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label134.Location = New System.Drawing.Point(445, 40)
        Me.Label134.Name = "Label134"
        Me.Label134.Size = New System.Drawing.Size(44, 16)
        Me.Label134.TabIndex = 137
        Me.Label134.Text = "火災燈"
        '
        'Spec_WTB_FO_ComboBox
        '
        Me.Spec_WTB_FO_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_FO_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_FO_ComboBox.Location = New System.Drawing.Point(518, 36)
        Me.Spec_WTB_FO_ComboBox.Name = "Spec_WTB_FO_ComboBox"
        Me.Spec_WTB_FO_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_FO_ComboBox.TabIndex = 138
        '
        'Label135
        '
        Me.Label135.AutoSize = True
        Me.Label135.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label135.Location = New System.Drawing.Point(320, 40)
        Me.Label135.Name = "Label135"
        Me.Label135.Size = New System.Drawing.Size(68, 16)
        Me.Label135.TabIndex = 135
        Me.Label135.Text = "緊急電源燈"
        '
        'Spec_WTB_Urgent_ComboBox
        '
        Me.Spec_WTB_Urgent_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_Urgent_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_Urgent_ComboBox.Location = New System.Drawing.Point(389, 36)
        Me.Spec_WTB_Urgent_ComboBox.Name = "Spec_WTB_Urgent_ComboBox"
        Me.Spec_WTB_Urgent_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_Urgent_ComboBox.TabIndex = 136
        '
        'Label136
        '
        Me.Label136.AutoSize = True
        Me.Label136.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label136.Location = New System.Drawing.Point(204, 40)
        Me.Label136.Name = "Label136"
        Me.Label136.Size = New System.Drawing.Size(44, 16)
        Me.Label136.TabIndex = 133
        Me.Label136.Text = "正常燈"
        '
        'Spec_WTB_Normal_ComboBox
        '
        Me.Spec_WTB_Normal_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_Normal_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_Normal_ComboBox.Location = New System.Drawing.Point(263, 36)
        Me.Spec_WTB_Normal_ComboBox.Name = "Spec_WTB_Normal_ComboBox"
        Me.Spec_WTB_Normal_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_Normal_ComboBox.TabIndex = 134
        '
        'Label108
        '
        Me.Label108.AutoSize = True
        Me.Label108.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label108.Location = New System.Drawing.Point(204, 133)
        Me.Label108.Name = "Label108"
        Me.Label108.Size = New System.Drawing.Size(66, 16)
        Me.Label108.TabIndex = 131
        Me.Label108.Text = "Check開關"
        '
        'Spec_WTB_ChkSW_ComboBox
        '
        Me.Spec_WTB_ChkSW_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_ChkSW_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_ChkSW_ComboBox.Location = New System.Drawing.Point(263, 129)
        Me.Spec_WTB_ChkSW_ComboBox.Name = "Spec_WTB_ChkSW_ComboBox"
        Me.Spec_WTB_ChkSW_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_ChkSW_ComboBox.TabIndex = 132
        '
        'Label105
        '
        Me.Label105.AutoSize = True
        Me.Label105.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label105.Location = New System.Drawing.Point(445, 9)
        Me.Label105.Name = "Label105"
        Me.Label105.Size = New System.Drawing.Size(44, 16)
        Me.Label105.TabIndex = 129
        Me.Label105.Text = "消防燈"
        '
        'Spec_WTB_FM_ComboBox
        '
        Me.Spec_WTB_FM_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_FM_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_FM_ComboBox.Location = New System.Drawing.Point(518, 5)
        Me.Spec_WTB_FM_ComboBox.Name = "Spec_WTB_FM_ComboBox"
        Me.Spec_WTB_FM_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_FM_ComboBox.TabIndex = 130
        '
        'Label102
        '
        Me.Label102.AutoSize = True
        Me.Label102.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label102.Location = New System.Drawing.Point(320, 9)
        Me.Label102.Name = "Label102"
        Me.Label102.Size = New System.Drawing.Size(44, 16)
        Me.Label102.TabIndex = 127
        Me.Label102.Text = "休止燈"
        '
        'Spec_WTB_Stop_ComboBox
        '
        Me.Spec_WTB_Stop_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_Stop_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_Stop_ComboBox.Location = New System.Drawing.Point(389, 5)
        Me.Spec_WTB_Stop_ComboBox.Name = "Spec_WTB_Stop_ComboBox"
        Me.Spec_WTB_Stop_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_Stop_ComboBox.TabIndex = 128
        '
        'Label98
        '
        Me.Label98.AutoSize = True
        Me.Label98.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label98.Location = New System.Drawing.Point(204, 9)
        Me.Label98.Name = "Label98"
        Me.Label98.Size = New System.Drawing.Size(44, 16)
        Me.Label98.TabIndex = 125
        Me.Label98.Text = "故障燈"
        '
        'Spec_WTB_Error_ComboBox
        '
        Me.Spec_WTB_Error_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_Error_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_Error_ComboBox.Location = New System.Drawing.Point(263, 5)
        Me.Spec_WTB_Error_ComboBox.Name = "Spec_WTB_Error_ComboBox"
        Me.Spec_WTB_Error_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_Error_ComboBox.TabIndex = 126
        '
        'Label68
        '
        Me.Label68.AutoSize = True
        Me.Label68.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label68.Location = New System.Drawing.Point(27, 9)
        Me.Label68.Name = "Label68"
        Me.Label68.Size = New System.Drawing.Size(34, 16)
        Me.Label68.TabIndex = 76
        Me.Label68.Text = "WTB"
        '
        'Spec_WTB_ComboBox
        '
        Me.Spec_WTB_ComboBox.FormattingEnabled = True
        Me.Spec_WTB_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_WTB_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_WTB_ComboBox.Name = "Spec_WTB_ComboBox"
        Me.Spec_WTB_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_WTB_ComboBox.TabIndex = 77
        '
        'Spec_IF79x_Panel
        '
        Me.Spec_IF79x_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_IF79x_Panel.Controls.Add(Me.Label120)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Spec_IF79x_IDM0_ComboBox)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Label121)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Spec_IF79x_ID12_ComboBox)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Label119)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Spec_IF79x_ID5_ComboBox)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Label118)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Spec_IF79x_ID4_ComboBox)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Label117)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Spec_IF79x_ID0_ComboBox)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Label69)
        Me.Spec_IF79x_Panel.Controls.Add(Me.Spec_IF79x_ComboBox)
        Me.Spec_IF79x_Panel.Location = New System.Drawing.Point(3, 371)
        Me.Spec_IF79x_Panel.Name = "Spec_IF79x_Panel"
        Me.Spec_IF79x_Panel.Size = New System.Drawing.Size(580, 68)
        Me.Spec_IF79x_Panel.TabIndex = 209
        '
        'Label120
        '
        Me.Label120.AutoSize = True
        Me.Label120.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label120.Location = New System.Drawing.Point(313, 43)
        Me.Label120.Name = "Label120"
        Me.Label120.Size = New System.Drawing.Size(60, 16)
        Me.Label120.TabIndex = 90
        Me.Label120.Text = "ID = M0 :"
        '
        'Spec_IF79x_IDM0_ComboBox
        '
        Me.Spec_IF79x_IDM0_ComboBox.FormattingEnabled = True
        Me.Spec_IF79x_IDM0_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_IF79x_IDM0_ComboBox.Location = New System.Drawing.Point(378, 39)
        Me.Spec_IF79x_IDM0_ComboBox.Name = "Spec_IF79x_IDM0_ComboBox"
        Me.Spec_IF79x_IDM0_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_IF79x_IDM0_ComboBox.TabIndex = 91
        '
        'Label121
        '
        Me.Label121.AutoSize = True
        Me.Label121.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label121.Location = New System.Drawing.Point(203, 43)
        Me.Label121.Name = "Label121"
        Me.Label121.Size = New System.Drawing.Size(55, 16)
        Me.Label121.TabIndex = 88
        Me.Label121.Text = "ID = 12 :"
        '
        'Spec_IF79x_ID12_ComboBox
        '
        Me.Spec_IF79x_ID12_ComboBox.FormattingEnabled = True
        Me.Spec_IF79x_ID12_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_IF79x_ID12_ComboBox.Location = New System.Drawing.Point(263, 39)
        Me.Spec_IF79x_ID12_ComboBox.Name = "Spec_IF79x_ID12_ComboBox"
        Me.Spec_IF79x_ID12_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_IF79x_ID12_ComboBox.TabIndex = 89
        '
        'Label119
        '
        Me.Label119.AutoSize = True
        Me.Label119.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label119.Location = New System.Drawing.Point(433, 9)
        Me.Label119.Name = "Label119"
        Me.Label119.Size = New System.Drawing.Size(48, 16)
        Me.Label119.TabIndex = 86
        Me.Label119.Text = "ID = 5 :"
        '
        'Spec_IF79x_ID5_ComboBox
        '
        Me.Spec_IF79x_ID5_ComboBox.FormattingEnabled = True
        Me.Spec_IF79x_ID5_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_IF79x_ID5_ComboBox.Location = New System.Drawing.Point(492, 5)
        Me.Spec_IF79x_ID5_ComboBox.Name = "Spec_IF79x_ID5_ComboBox"
        Me.Spec_IF79x_ID5_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_IF79x_ID5_ComboBox.TabIndex = 87
        '
        'Label118
        '
        Me.Label118.AutoSize = True
        Me.Label118.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label118.Location = New System.Drawing.Point(318, 9)
        Me.Label118.Name = "Label118"
        Me.Label118.Size = New System.Drawing.Size(48, 16)
        Me.Label118.TabIndex = 84
        Me.Label118.Text = "ID = 4 :"
        '
        'Spec_IF79x_ID4_ComboBox
        '
        Me.Spec_IF79x_ID4_ComboBox.FormattingEnabled = True
        Me.Spec_IF79x_ID4_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_IF79x_ID4_ComboBox.Location = New System.Drawing.Point(378, 5)
        Me.Spec_IF79x_ID4_ComboBox.Name = "Spec_IF79x_ID4_ComboBox"
        Me.Spec_IF79x_ID4_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_IF79x_ID4_ComboBox.TabIndex = 85
        '
        'Label117
        '
        Me.Label117.AutoSize = True
        Me.Label117.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label117.Location = New System.Drawing.Point(203, 9)
        Me.Label117.Name = "Label117"
        Me.Label117.Size = New System.Drawing.Size(48, 16)
        Me.Label117.TabIndex = 82
        Me.Label117.Text = "ID = 0 :"
        '
        'Spec_IF79x_ID0_ComboBox
        '
        Me.Spec_IF79x_ID0_ComboBox.FormattingEnabled = True
        Me.Spec_IF79x_ID0_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_IF79x_ID0_ComboBox.Location = New System.Drawing.Point(263, 5)
        Me.Spec_IF79x_ID0_ComboBox.Name = "Spec_IF79x_ID0_ComboBox"
        Me.Spec_IF79x_ID0_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_IF79x_ID0_ComboBox.TabIndex = 83
        '
        'Label69
        '
        Me.Label69.AutoSize = True
        Me.Label69.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label69.Location = New System.Drawing.Point(27, 9)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(97, 16)
        Me.Label69.TabIndex = 76
        Me.Label69.Text = "IF79x入出力位置"
        '
        'Spec_IF79x_ComboBox
        '
        Me.Spec_IF79x_ComboBox.FormattingEnabled = True
        Me.Spec_IF79x_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_IF79x_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_IF79x_ComboBox.Name = "Spec_IF79x_ComboBox"
        Me.Spec_IF79x_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_IF79x_ComboBox.TabIndex = 77
        '
        'Spec_EachStop_Panel
        '
        Me.Spec_EachStop_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_EachStop_Panel.Controls.Add(Me.Label71)
        Me.Spec_EachStop_Panel.Controls.Add(Me.Spec_EachStop_ComboBox)
        Me.Spec_EachStop_Panel.Location = New System.Drawing.Point(3, 445)
        Me.Spec_EachStop_Panel.Name = "Spec_EachStop_Panel"
        Me.Spec_EachStop_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_EachStop_Panel.TabIndex = 211
        '
        'Label71
        '
        Me.Label71.AutoSize = True
        Me.Label71.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label71.Location = New System.Drawing.Point(27, 9)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(56, 16)
        Me.Label71.TabIndex = 76
        Me.Label71.Text = "各停開關"
        '
        'Spec_EachStop_ComboBox
        '
        Me.Spec_EachStop_ComboBox.FormattingEnabled = True
        Me.Spec_EachStop_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_EachStop_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_EachStop_ComboBox.Name = "Spec_EachStop_ComboBox"
        Me.Spec_EachStop_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_EachStop_ComboBox.TabIndex = 77
        '
        'Panel115
        '
        Me.Panel115.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel115.Controls.Add(Me.Label_SPEC_INSTALL_OPE)
        Me.Panel115.Controls.Add(Me.Spec_install_ope_ComboBox)
        Me.Panel115.Location = New System.Drawing.Point(3, 487)
        Me.Panel115.Name = "Panel115"
        Me.Panel115.Size = New System.Drawing.Size(580, 36)
        Me.Panel115.TabIndex = 212
        '
        'Label_SPEC_INSTALL_OPE
        '
        Me.Label_SPEC_INSTALL_OPE.AutoSize = True
        Me.Label_SPEC_INSTALL_OPE.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label_SPEC_INSTALL_OPE.Location = New System.Drawing.Point(27, 9)
        Me.Label_SPEC_INSTALL_OPE.Name = "Label_SPEC_INSTALL_OPE"
        Me.Label_SPEC_INSTALL_OPE.Size = New System.Drawing.Size(56, 16)
        Me.Label_SPEC_INSTALL_OPE.TabIndex = 76
        Me.Label_SPEC_INSTALL_OPE.Text = "拒付運轉"
        '
        'Spec_install_ope_ComboBox
        '
        Me.Spec_install_ope_ComboBox.FormattingEnabled = True
        Me.Spec_install_ope_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_install_ope_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_install_ope_ComboBox.Name = "Spec_install_ope_ComboBox"
        Me.Spec_install_ope_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_install_ope_ComboBox.TabIndex = 77
        '
        'Spec_Operation_Panel
        '
        Me.Spec_Operation_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Spec_Operation_Panel.Controls.Add(Me.Spec_Operation_Label)
        Me.Spec_Operation_Panel.Controls.Add(Me.Spec_Operation_ComboBox)
        Me.Spec_Operation_Panel.Location = New System.Drawing.Point(3, 529)
        Me.Spec_Operation_Panel.Name = "Spec_Operation_Panel"
        Me.Spec_Operation_Panel.Size = New System.Drawing.Size(580, 36)
        Me.Spec_Operation_Panel.TabIndex = 213
        '
        'Spec_Operation_Label
        '
        Me.Spec_Operation_Label.AutoSize = True
        Me.Spec_Operation_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Spec_Operation_Label.Location = New System.Drawing.Point(27, 9)
        Me.Spec_Operation_Label.Name = "Spec_Operation_Label"
        Me.Spec_Operation_Label.Size = New System.Drawing.Size(56, 16)
        Me.Spec_Operation_Label.TabIndex = 47
        Me.Spec_Operation_Label.Text = "操作方式"
        '
        'Spec_Operation_ComboBox
        '
        Me.Spec_Operation_ComboBox.FormattingEnabled = True
        Me.Spec_Operation_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.Spec_Operation_ComboBox.Location = New System.Drawing.Point(147, 5)
        Me.Spec_Operation_ComboBox.Name = "Spec_Operation_ComboBox"
        Me.Spec_Operation_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.Spec_Operation_ComboBox.TabIndex = 48
        '
        'Use_SpecTWFP17_CheckBox
        '
        Me.Use_SpecTWFP17_CheckBox.AutoSize = True
        Me.Use_SpecTWFP17_CheckBox.Enabled = False
        Me.Use_SpecTWFP17_CheckBox.Location = New System.Drawing.Point(102, 16)
        Me.Use_SpecTWFP17_CheckBox.Name = "Use_SpecTWFP17_CheckBox"
        Me.Use_SpecTWFP17_CheckBox.Size = New System.Drawing.Size(59, 20)
        Me.Use_SpecTWFP17_CheckBox.TabIndex = 17
        Me.Use_SpecTWFP17_CheckBox.Text = "FP-17"
        Me.Use_SpecTWFP17_CheckBox.UseVisualStyleBackColor = True
        '
        'Use_SpecTWIDU_CheckBox
        '
        Me.Use_SpecTWIDU_CheckBox.AutoSize = True
        Me.Use_SpecTWIDU_CheckBox.Enabled = False
        Me.Use_SpecTWIDU_CheckBox.Location = New System.Drawing.Point(9, 16)
        Me.Use_SpecTWIDU_CheckBox.Name = "Use_SpecTWIDU_CheckBox"
        Me.Use_SpecTWIDU_CheckBox.Size = New System.Drawing.Size(85, 20)
        Me.Use_SpecTWIDU_CheckBox.TabIndex = 15
        Me.Use_SpecTWIDU_CheckBox.Text = "Z/REXIA-T"
        Me.Use_SpecTWIDU_CheckBox.UseVisualStyleBackColor = True
        '
        'DWG_TabPage
        '
        Me.DWG_TabPage.Controls.Add(Me.DWG_GroupBox)
        Me.DWG_TabPage.Controls.Add(Me.Use_prk_CheckBox)
        Me.DWG_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.DWG_TabPage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DWG_TabPage.Name = "DWG_TabPage"
        Me.DWG_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.DWG_TabPage.TabIndex = 2
        Me.DWG_TabPage.Text = "送狀"
        Me.DWG_TabPage.UseVisualStyleBackColor = True
        '
        'DWG_GroupBox
        '
        Me.DWG_GroupBox.Controls.Add(Me.Label194)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_VonicStd_ComboBox)
        Me.DWG_GroupBox.Controls.Add(Me.Label193)
        Me.DWG_GroupBox.Controls.Add(Me.Label192)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_Produce_CheckedListBox)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_Construction_CheckedListBox)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_StdPage_Button)
        Me.DWG_GroupBox.Controls.Add(Me.Label58)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_Page_ChkAllButton)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_PageNum_TextBox)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_Page_CheckedListBox)
        Me.DWG_GroupBox.Controls.Add(Me.Label60)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_Page_unChkAllButton)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_PrkName_ComboBox)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_Page_SubButton)
        Me.DWG_GroupBox.Controls.Add(Me.Label59)
        Me.DWG_GroupBox.Controls.Add(Me.DWG_Page_AddButton)
        Me.DWG_GroupBox.Enabled = False
        Me.DWG_GroupBox.Location = New System.Drawing.Point(15, 20)
        Me.DWG_GroupBox.Name = "DWG_GroupBox"
        Me.DWG_GroupBox.Size = New System.Drawing.Size(630, 543)
        Me.DWG_GroupBox.TabIndex = 60
        Me.DWG_GroupBox.TabStop = False
        '
        'Label194
        '
        Me.Label194.AutoSize = True
        Me.Label194.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label194.Location = New System.Drawing.Point(6, 45)
        Me.Label194.Name = "Label194"
        Me.Label194.Size = New System.Drawing.Size(70, 16)
        Me.Label194.TabIndex = 121
        Me.Label194.Text = "Vonic標準 :"
        '
        'DWG_VonicStd_ComboBox
        '
        Me.DWG_VonicStd_ComboBox.FormattingEnabled = True
        Me.DWG_VonicStd_ComboBox.Items.AddRange(New Object() {"○", "×"})
        Me.DWG_VonicStd_ComboBox.Location = New System.Drawing.Point(82, 42)
        Me.DWG_VonicStd_ComboBox.Name = "DWG_VonicStd_ComboBox"
        Me.DWG_VonicStd_ComboBox.Size = New System.Drawing.Size(45, 24)
        Me.DWG_VonicStd_ComboBox.TabIndex = 122
        '
        'Label193
        '
        Me.Label193.AutoSize = True
        Me.Label193.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label193.Location = New System.Drawing.Point(459, 19)
        Me.Label193.Name = "Label193"
        Me.Label193.Size = New System.Drawing.Size(32, 16)
        Me.Label193.TabIndex = 62
        Me.Label193.Text = "製造"
        '
        'Label192
        '
        Me.Label192.AutoSize = True
        Me.Label192.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label192.Location = New System.Drawing.Point(394, 19)
        Me.Label192.Name = "Label192"
        Me.Label192.Size = New System.Drawing.Size(59, 16)
        Me.Label192.TabIndex = 61
        Me.Label192.Text = "工務,現場"
        '
        'DWG_Produce_CheckedListBox
        '
        Me.DWG_Produce_CheckedListBox.CheckOnClick = True
        Me.DWG_Produce_CheckedListBox.FormattingEnabled = True
        Me.DWG_Produce_CheckedListBox.Location = New System.Drawing.Point(458, 42)
        Me.DWG_Produce_CheckedListBox.Name = "DWG_Produce_CheckedListBox"
        Me.DWG_Produce_CheckedListBox.Size = New System.Drawing.Size(20, 490)
        Me.DWG_Produce_CheckedListBox.TabIndex = 60
        '
        'DWG_Construction_CheckedListBox
        '
        Me.DWG_Construction_CheckedListBox.CheckOnClick = True
        Me.DWG_Construction_CheckedListBox.FormattingEnabled = True
        Me.DWG_Construction_CheckedListBox.Location = New System.Drawing.Point(397, 42)
        Me.DWG_Construction_CheckedListBox.Name = "DWG_Construction_CheckedListBox"
        Me.DWG_Construction_CheckedListBox.Size = New System.Drawing.Size(20, 490)
        Me.DWG_Construction_CheckedListBox.TabIndex = 59
        '
        'DWG_StdPage_Button
        '
        Me.DWG_StdPage_Button.Location = New System.Drawing.Point(6, 75)
        Me.DWG_StdPage_Button.Name = "DWG_StdPage_Button"
        Me.DWG_StdPage_Button.Size = New System.Drawing.Size(161, 24)
        Me.DWG_StdPage_Button.TabIndex = 58
        Me.DWG_StdPage_Button.Text = "基本版型套用"
        Me.DWG_StdPage_Button.UseVisualStyleBackColor = True
        '
        'Label58
        '
        Me.Label58.AutoSize = True
        Me.Label58.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label58.Location = New System.Drawing.Point(6, 106)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(68, 16)
        Me.Label58.TabIndex = 55
        Me.Label58.Text = "PRK Name"
        '
        'DWG_Page_ChkAllButton
        '
        Me.DWG_Page_ChkAllButton.Location = New System.Drawing.Point(486, 42)
        Me.DWG_Page_ChkAllButton.Name = "DWG_Page_ChkAllButton"
        Me.DWG_Page_ChkAllButton.Size = New System.Drawing.Size(100, 23)
        Me.DWG_Page_ChkAllButton.TabIndex = 53
        Me.DWG_Page_ChkAllButton.Text = "v Check All"
        Me.DWG_Page_ChkAllButton.UseVisualStyleBackColor = True
        '
        'DWG_PageNum_TextBox
        '
        Me.DWG_PageNum_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.DWG_PageNum_TextBox.Location = New System.Drawing.Point(98, 129)
        Me.DWG_PageNum_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DWG_PageNum_TextBox.MaxLength = 50
        Me.DWG_PageNum_TextBox.Name = "DWG_PageNum_TextBox"
        Me.DWG_PageNum_TextBox.Size = New System.Drawing.Size(38, 23)
        Me.DWG_PageNum_TextBox.TabIndex = 48
        Me.DWG_PageNum_TextBox.Text = "6"
        '
        'DWG_Page_CheckedListBox
        '
        Me.DWG_Page_CheckedListBox.CheckOnClick = True
        Me.DWG_Page_CheckedListBox.FormattingEnabled = True
        Me.DWG_Page_CheckedListBox.Location = New System.Drawing.Point(173, 42)
        Me.DWG_Page_CheckedListBox.Name = "DWG_Page_CheckedListBox"
        Me.DWG_Page_CheckedListBox.Size = New System.Drawing.Size(206, 490)
        Me.DWG_Page_CheckedListBox.TabIndex = 52
        '
        'Label60
        '
        Me.Label60.AutoSize = True
        Me.Label60.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label60.Location = New System.Drawing.Point(173, 19)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(116, 16)
        Me.Label60.TabIndex = 57
        Me.Label60.Text = "v輸出項目必要打勾v"
        '
        'DWG_Page_unChkAllButton
        '
        Me.DWG_Page_unChkAllButton.Location = New System.Drawing.Point(486, 71)
        Me.DWG_Page_unChkAllButton.Name = "DWG_Page_unChkAllButton"
        Me.DWG_Page_unChkAllButton.Size = New System.Drawing.Size(100, 24)
        Me.DWG_Page_unChkAllButton.TabIndex = 54
        Me.DWG_Page_unChkAllButton.Text = "x Uncheck All"
        Me.DWG_Page_unChkAllButton.UseVisualStyleBackColor = True
        '
        'DWG_PrkName_ComboBox
        '
        Me.DWG_PrkName_ComboBox.FormattingEnabled = True
        Me.DWG_PrkName_ComboBox.Items.AddRange(New Object() {""})
        Me.DWG_PrkName_ComboBox.Location = New System.Drawing.Point(6, 129)
        Me.DWG_PrkName_ComboBox.Name = "DWG_PrkName_ComboBox"
        Me.DWG_PrkName_ComboBox.Size = New System.Drawing.Size(83, 24)
        Me.DWG_PrkName_ComboBox.TabIndex = 49
        '
        'DWG_Page_SubButton
        '
        Me.DWG_Page_SubButton.AccessibleDescription = "右側勾選後可刪除"
        Me.DWG_Page_SubButton.Location = New System.Drawing.Point(142, 159)
        Me.DWG_Page_SubButton.Name = "DWG_Page_SubButton"
        Me.DWG_Page_SubButton.Size = New System.Drawing.Size(25, 23)
        Me.DWG_Page_SubButton.TabIndex = 51
        Me.DWG_Page_SubButton.Text = "-"
        Me.DWG_Page_SubButton.UseVisualStyleBackColor = True
        '
        'Label59
        '
        Me.Label59.AutoSize = True
        Me.Label59.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label59.Location = New System.Drawing.Point(98, 106)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(37, 16)
        Me.Label59.TabIndex = 56
        Me.Label59.Text = "Page"
        '
        'DWG_Page_AddButton
        '
        Me.DWG_Page_AddButton.AccessibleDescription = "左側名稱數量選取好後即可新增"
        Me.DWG_Page_AddButton.Location = New System.Drawing.Point(142, 129)
        Me.DWG_Page_AddButton.Name = "DWG_Page_AddButton"
        Me.DWG_Page_AddButton.Size = New System.Drawing.Size(25, 23)
        Me.DWG_Page_AddButton.TabIndex = 50
        Me.DWG_Page_AddButton.Text = "+"
        Me.DWG_Page_AddButton.UseVisualStyleBackColor = True
        '
        'Use_prk_CheckBox
        '
        Me.Use_prk_CheckBox.AutoSize = True
        Me.Use_prk_CheckBox.Enabled = False
        Me.Use_prk_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_prk_CheckBox.Name = "Use_prk_CheckBox"
        Me.Use_prk_CheckBox.Size = New System.Drawing.Size(83, 20)
        Me.Use_prk_CheckBox.TabIndex = 58
        Me.Use_prk_CheckBox.Text = "(暫停使用)"
        Me.Use_prk_CheckBox.UseVisualStyleBackColor = True
        '
        'ProgramChange_TabPage
        '
        Me.ProgramChange_TabPage.Controls.Add(Me.TabControl3)
        Me.ProgramChange_TabPage.Controls.Add(Me.Use_Program_CheckBox)
        Me.ProgramChange_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.ProgramChange_TabPage.Name = "ProgramChange_TabPage"
        Me.ProgramChange_TabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.ProgramChange_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.ProgramChange_TabPage.TabIndex = 6
        Me.ProgramChange_TabPage.Text = "程式變更"
        Me.ProgramChange_TabPage.UseVisualStyleBackColor = True
        '
        'TabControl3
        '
        Me.TabControl3.Controls.Add(Me.TabPage5)
        Me.TabControl3.Controls.Add(Me.TabPage6)
        Me.TabControl3.Location = New System.Drawing.Point(6, 23)
        Me.TabControl3.Name = "TabControl3"
        Me.TabControl3.SelectedIndex = 0
        Me.TabControl3.Size = New System.Drawing.Size(652, 555)
        Me.TabControl3.TabIndex = 163
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.ProgramChange_FlowLayoutPanel)
        Me.TabPage5.Location = New System.Drawing.Point(4, 25)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(644, 526)
        Me.TabPage5.TabIndex = 0
        Me.TabPage5.Text = "Page1"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'ProgramChange_FlowLayoutPanel
        '
        Me.ProgramChange_FlowLayoutPanel.AutoScroll = True
        Me.ProgramChange_FlowLayoutPanel.Controls.Add(Me.use_ProgramChg_Panel1)
        Me.ProgramChange_FlowLayoutPanel.Controls.Add(Me.use_ProgramChg_Panel2)
        Me.ProgramChange_FlowLayoutPanel.Controls.Add(Me.use_ProgramChg_Panel3)
        Me.ProgramChange_FlowLayoutPanel.Controls.Add(Me.use_ProgramChg_Panel5)
        Me.ProgramChange_FlowLayoutPanel.Enabled = False
        Me.ProgramChange_FlowLayoutPanel.Location = New System.Drawing.Point(6, 6)
        Me.ProgramChange_FlowLayoutPanel.Name = "ProgramChange_FlowLayoutPanel"
        Me.ProgramChange_FlowLayoutPanel.Size = New System.Drawing.Size(632, 514)
        Me.ProgramChange_FlowLayoutPanel.TabIndex = 162
        '
        'use_ProgramChg_Panel1
        '
        Me.use_ProgramChg_Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.use_ProgramChg_Panel1.Controls.Add(Me.Label33)
        Me.use_ProgramChg_Panel1.Controls.Add(Me.Label32)
        Me.use_ProgramChg_Panel1.Controls.Add(Me.PrmList_1_reason_TextBox)
        Me.use_ProgramChg_Panel1.Location = New System.Drawing.Point(3, 3)
        Me.use_ProgramChg_Panel1.Name = "use_ProgramChg_Panel1"
        Me.use_ProgramChg_Panel1.Size = New System.Drawing.Size(600, 105)
        Me.use_ProgramChg_Panel1.TabIndex = 0
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label33.Location = New System.Drawing.Point(3, 11)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(72, 16)
        Me.Label33.TabIndex = 42
        Me.Label33.Text = "1.對象ROM"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label32.Location = New System.Drawing.Point(17, 36)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(62, 16)
        Me.Label32.TabIndex = 40
        Me.Label32.Text = "變更理由 :"
        '
        'PrmList_1_reason_TextBox
        '
        Me.PrmList_1_reason_TextBox.Location = New System.Drawing.Point(85, 36)
        Me.PrmList_1_reason_TextBox.Multiline = True
        Me.PrmList_1_reason_TextBox.Name = "PrmList_1_reason_TextBox"
        Me.PrmList_1_reason_TextBox.Size = New System.Drawing.Size(425, 57)
        Me.PrmList_1_reason_TextBox.TabIndex = 1
        '
        'use_ProgramChg_Panel2
        '
        Me.use_ProgramChg_Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.use_ProgramChg_Panel2.Controls.Add(Me.PrmList_2_Other_CheckBox)
        Me.use_ProgramChg_Panel2.Controls.Add(Me.PrmList_2_Tower_CheckBox)
        Me.use_ProgramChg_Panel2.Controls.Add(Me.PrmList_2_COP_CheckBox)
        Me.use_ProgramChg_Panel2.Controls.Add(Me.PrmList_2_test_CheckBox)
        Me.use_ProgramChg_Panel2.Controls.Add(Me.PrmList_2_test_TextBox)
        Me.use_ProgramChg_Panel2.Controls.Add(Me.PrmList_2_COP_TextBox)
        Me.use_ProgramChg_Panel2.Controls.Add(Me.PrmList_2_tower_TextBox)
        Me.use_ProgramChg_Panel2.Controls.Add(Me.PrmList_2_other_TextBox)
        Me.use_ProgramChg_Panel2.Controls.Add(Me.Label34)
        Me.use_ProgramChg_Panel2.Location = New System.Drawing.Point(3, 114)
        Me.use_ProgramChg_Panel2.Name = "use_ProgramChg_Panel2"
        Me.use_ProgramChg_Panel2.Size = New System.Drawing.Size(600, 149)
        Me.use_ProgramChg_Panel2.TabIndex = 159
        '
        'PrmList_2_Other_CheckBox
        '
        Me.PrmList_2_Other_CheckBox.AutoSize = True
        Me.PrmList_2_Other_CheckBox.Location = New System.Drawing.Point(20, 112)
        Me.PrmList_2_Other_CheckBox.Name = "PrmList_2_Other_CheckBox"
        Me.PrmList_2_Other_CheckBox.Size = New System.Drawing.Size(51, 20)
        Me.PrmList_2_Other_CheckBox.TabIndex = 5
        Me.PrmList_2_Other_CheckBox.Text = "其他"
        Me.PrmList_2_Other_CheckBox.UseVisualStyleBackColor = True
        '
        'PrmList_2_Tower_CheckBox
        '
        Me.PrmList_2_Tower_CheckBox.AutoSize = True
        Me.PrmList_2_Tower_CheckBox.Location = New System.Drawing.Point(20, 85)
        Me.PrmList_2_Tower_CheckBox.Name = "PrmList_2_Tower_CheckBox"
        Me.PrmList_2_Tower_CheckBox.Size = New System.Drawing.Size(87, 20)
        Me.PrmList_2_Tower_CheckBox.TabIndex = 4
        Me.PrmList_2_Tower_CheckBox.Text = "研修測試塔"
        Me.PrmList_2_Tower_CheckBox.UseVisualStyleBackColor = True
        '
        'PrmList_2_COP_CheckBox
        '
        Me.PrmList_2_COP_CheckBox.AutoSize = True
        Me.PrmList_2_COP_CheckBox.Location = New System.Drawing.Point(20, 57)
        Me.PrmList_2_COP_CheckBox.Name = "PrmList_2_COP_CheckBox"
        Me.PrmList_2_COP_CheckBox.Size = New System.Drawing.Size(63, 20)
        Me.PrmList_2_COP_CheckBox.TabIndex = 3
        Me.PrmList_2_COP_CheckBox.Text = "控制盤"
        Me.PrmList_2_COP_CheckBox.UseVisualStyleBackColor = True
        '
        'PrmList_2_test_CheckBox
        '
        Me.PrmList_2_test_CheckBox.AutoSize = True
        Me.PrmList_2_test_CheckBox.Location = New System.Drawing.Point(20, 28)
        Me.PrmList_2_test_CheckBox.Name = "PrmList_2_test_CheckBox"
        Me.PrmList_2_test_CheckBox.Size = New System.Drawing.Size(75, 20)
        Me.PrmList_2_test_CheckBox.TabIndex = 2
        Me.PrmList_2_test_CheckBox.Text = "測試裝置"
        Me.PrmList_2_test_CheckBox.UseVisualStyleBackColor = True
        '
        'PrmList_2_test_TextBox
        '
        Me.PrmList_2_test_TextBox.Location = New System.Drawing.Point(122, 25)
        Me.PrmList_2_test_TextBox.Name = "PrmList_2_test_TextBox"
        Me.PrmList_2_test_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.PrmList_2_test_TextBox.TabIndex = 6
        '
        'PrmList_2_COP_TextBox
        '
        Me.PrmList_2_COP_TextBox.Location = New System.Drawing.Point(122, 54)
        Me.PrmList_2_COP_TextBox.Name = "PrmList_2_COP_TextBox"
        Me.PrmList_2_COP_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.PrmList_2_COP_TextBox.TabIndex = 7
        '
        'PrmList_2_tower_TextBox
        '
        Me.PrmList_2_tower_TextBox.Location = New System.Drawing.Point(122, 83)
        Me.PrmList_2_tower_TextBox.Name = "PrmList_2_tower_TextBox"
        Me.PrmList_2_tower_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.PrmList_2_tower_TextBox.TabIndex = 8
        '
        'PrmList_2_other_TextBox
        '
        Me.PrmList_2_other_TextBox.Location = New System.Drawing.Point(122, 112)
        Me.PrmList_2_other_TextBox.Name = "PrmList_2_other_TextBox"
        Me.PrmList_2_other_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.PrmList_2_other_TextBox.TabIndex = 9
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label34.Location = New System.Drawing.Point(3, 9)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(134, 16)
        Me.Label34.TabIndex = 43
        Me.Label34.Text = "2.使用裝置(擔當者記入)"
        '
        'use_ProgramChg_Panel3
        '
        Me.use_ProgramChg_Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.use_ProgramChg_Panel3.Controls.Add(Me.PrmList_3_debug_CheckBox)
        Me.use_ProgramChg_Panel3.Controls.Add(Me.PrmList_3_excute_CheckBox)
        Me.use_ProgramChg_Panel3.Controls.Add(Me.PrmList_3_confirm_CheckBox)
        Me.use_ProgramChg_Panel3.Controls.Add(Me.PrmList_3_other_Checkbox)
        Me.use_ProgramChg_Panel3.Controls.Add(Me.PrmList_3_test_CheckBox)
        Me.use_ProgramChg_Panel3.Controls.Add(Me.PrmList_3_other_TextBox)
        Me.use_ProgramChg_Panel3.Controls.Add(Me.Label35)
        Me.use_ProgramChg_Panel3.Location = New System.Drawing.Point(3, 269)
        Me.use_ProgramChg_Panel3.Name = "use_ProgramChg_Panel3"
        Me.use_ProgramChg_Panel3.Size = New System.Drawing.Size(600, 124)
        Me.use_ProgramChg_Panel3.TabIndex = 160
        '
        'PrmList_3_debug_CheckBox
        '
        Me.PrmList_3_debug_CheckBox.AutoSize = True
        Me.PrmList_3_debug_CheckBox.Location = New System.Drawing.Point(20, 27)
        Me.PrmList_3_debug_CheckBox.Name = "PrmList_3_debug_CheckBox"
        Me.PrmList_3_debug_CheckBox.Size = New System.Drawing.Size(175, 20)
        Me.PrmList_3_debug_CheckBox.TabIndex = 10
        Me.PrmList_3_debug_CheckBox.Text = "使用程式LIST的上機DEBUG"
        Me.PrmList_3_debug_CheckBox.UseVisualStyleBackColor = True
        '
        'PrmList_3_excute_CheckBox
        '
        Me.PrmList_3_excute_CheckBox.AutoSize = True
        Me.PrmList_3_excute_CheckBox.Location = New System.Drawing.Point(259, 59)
        Me.PrmList_3_excute_CheckBox.Name = "PrmList_3_excute_CheckBox"
        Me.PrmList_3_excute_CheckBox.Size = New System.Drawing.Size(123, 20)
        Me.PrmList_3_excute_CheckBox.TabIndex = 13
        Me.PrmList_3_excute_CheckBox.Text = "確認程式實際執行"
        Me.PrmList_3_excute_CheckBox.UseVisualStyleBackColor = True
        '
        'PrmList_3_confirm_CheckBox
        '
        Me.PrmList_3_confirm_CheckBox.AutoSize = True
        Me.PrmList_3_confirm_CheckBox.Location = New System.Drawing.Point(20, 59)
        Me.PrmList_3_confirm_CheckBox.Name = "PrmList_3_confirm_CheckBox"
        Me.PrmList_3_confirm_CheckBox.Size = New System.Drawing.Size(111, 20)
        Me.PrmList_3_confirm_CheckBox.TabIndex = 12
        Me.PrmList_3_confirm_CheckBox.Text = "一般動作的確認"
        Me.PrmList_3_confirm_CheckBox.UseVisualStyleBackColor = True
        '
        'PrmList_3_other_Checkbox
        '
        Me.PrmList_3_other_Checkbox.AutoSize = True
        Me.PrmList_3_other_Checkbox.Location = New System.Drawing.Point(20, 91)
        Me.PrmList_3_other_Checkbox.Name = "PrmList_3_other_Checkbox"
        Me.PrmList_3_other_Checkbox.Size = New System.Drawing.Size(51, 20)
        Me.PrmList_3_other_Checkbox.TabIndex = 14
        Me.PrmList_3_other_Checkbox.Text = "其他"
        Me.PrmList_3_other_Checkbox.UseVisualStyleBackColor = True
        '
        'PrmList_3_test_CheckBox
        '
        Me.PrmList_3_test_CheckBox.AutoSize = True
        Me.PrmList_3_test_CheckBox.Location = New System.Drawing.Point(259, 27)
        Me.PrmList_3_test_CheckBox.Name = "PrmList_3_test_CheckBox"
        Me.PrmList_3_test_CheckBox.Size = New System.Drawing.Size(111, 20)
        Me.PrmList_3_test_CheckBox.TabIndex = 11
        Me.PrmList_3_test_CheckBox.Text = "設計内容的測試"
        Me.PrmList_3_test_CheckBox.UseVisualStyleBackColor = True
        '
        'PrmList_3_other_TextBox
        '
        Me.PrmList_3_other_TextBox.Location = New System.Drawing.Point(122, 90)
        Me.PrmList_3_other_TextBox.Name = "PrmList_3_other_TextBox"
        Me.PrmList_3_other_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.PrmList_3_other_TextBox.TabIndex = 15
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label35.Location = New System.Drawing.Point(3, 6)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(222, 16)
        Me.Label35.TabIndex = 99
        Me.Label35.Text = "3.檢查方法（擔當者有實施的部分打勾）"
        '
        'use_ProgramChg_Panel5
        '
        Me.use_ProgramChg_Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.use_ProgramChg_Panel5.Controls.Add(Me.Label52)
        Me.use_ProgramChg_Panel5.Controls.Add(Me.PrmList_5_review_CheckBox)
        Me.use_ProgramChg_Panel5.Location = New System.Drawing.Point(3, 399)
        Me.use_ProgramChg_Panel5.Name = "use_ProgramChg_Panel5"
        Me.use_ProgramChg_Panel5.Size = New System.Drawing.Size(600, 111)
        Me.use_ProgramChg_Panel5.TabIndex = 161
        '
        'Label52
        '
        Me.Label52.AutoSize = True
        Me.Label52.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label52.Location = New System.Drawing.Point(3, 12)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(126, 16)
        Me.Label52.TabIndex = 149
        Me.Label52.Text = "5.覆核（覆核者記入）"
        '
        'PrmList_5_review_CheckBox
        '
        Me.PrmList_5_review_CheckBox.AutoSize = True
        Me.PrmList_5_review_CheckBox.Location = New System.Drawing.Point(19, 39)
        Me.PrmList_5_review_CheckBox.Name = "PrmList_5_review_CheckBox"
        Me.PrmList_5_review_CheckBox.Size = New System.Drawing.Size(279, 20)
        Me.PrmList_5_review_CheckBox.TabIndex = 16
        Me.PrmList_5_review_CheckBox.Text = "確認上記的確認方法、檢查結果都必要且充分。"
        Me.PrmList_5_review_CheckBox.UseVisualStyleBackColor = True
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.FlowLayoutPanel1)
        Me.TabPage6.Location = New System.Drawing.Point(4, 25)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage6.Size = New System.Drawing.Size(644, 526)
        Me.TabPage6.TabIndex = 1
        Me.TabPage6.Text = "Page2"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.AutoScroll = True
        Me.FlowLayoutPanel1.Controls.Add(Me.use_ProgramChg_Panel4)
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(6, 6)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(630, 514)
        Me.FlowLayoutPanel1.TabIndex = 0
        '
        'use_ProgramChg_Panel4
        '
        Me.use_ProgramChg_Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label36)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel11)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label37)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel7)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label38)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel12)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label39)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel6)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label40)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel13)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label41)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel8)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label42)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel14)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label43)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel5)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label48)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel9)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label47)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel4)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label46)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel10)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label45)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Panel3)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label44)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label51)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.PrmList_4_content12_TextBox)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label50)
        Me.use_ProgramChg_Panel4.Controls.Add(Me.Label49)
        Me.use_ProgramChg_Panel4.Location = New System.Drawing.Point(3, 3)
        Me.use_ProgramChg_Panel4.Name = "use_ProgramChg_Panel4"
        Me.use_ProgramChg_Panel4.Size = New System.Drawing.Size(600, 603)
        Me.use_ProgramChg_Panel4.TabIndex = 160
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label36.Location = New System.Drawing.Point(3, 16)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(456, 16)
        Me.Label36.TabIndex = 107
        Me.Label36.Text = "4.檢查結果（擔當者記入）　（ＯＫ的話 ，在''要''打勾 、沒關聯的時候在''否''打勾）"
        '
        'Panel11
        '
        Me.Panel11.Controls.Add(Me.PrmList_4_yes12_RadioButton)
        Me.Panel11.Controls.Add(Me.PrmList_4_no12_RadioButton)
        Me.Panel11.Location = New System.Drawing.Point(14, 450)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(56, 20)
        Me.Panel11.TabIndex = 156
        '
        'PrmList_4_yes12_RadioButton
        '
        Me.PrmList_4_yes12_RadioButton.AutoSize = True
        Me.PrmList_4_yes12_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes12_RadioButton.Name = "PrmList_4_yes12_RadioButton"
        Me.PrmList_4_yes12_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes12_RadioButton.TabIndex = 23
        Me.PrmList_4_yes12_RadioButton.TabStop = True
        Me.PrmList_4_yes12_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no12_RadioButton
        '
        Me.PrmList_4_no12_RadioButton.AutoSize = True
        Me.PrmList_4_no12_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no12_RadioButton.Name = "PrmList_4_no12_RadioButton"
        Me.PrmList_4_no12_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no12_RadioButton.TabIndex = 24
        Me.PrmList_4_no12_RadioButton.TabStop = True
        Me.PrmList_4_no12_RadioButton.UseVisualStyleBackColor = True
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label37.Location = New System.Drawing.Point(14, 42)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(20, 16)
        Me.Label37.TabIndex = 108
        Me.Label37.Text = "要"
        '
        'Panel7
        '
        Me.Panel7.Controls.Add(Me.PrmList_4_yes8_RadioButton)
        Me.Panel7.Controls.Add(Me.PrmList_4_no8_RadioButton)
        Me.Panel7.Location = New System.Drawing.Point(14, 298)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(56, 20)
        Me.Panel7.TabIndex = 156
        '
        'PrmList_4_yes8_RadioButton
        '
        Me.PrmList_4_yes8_RadioButton.AutoSize = True
        Me.PrmList_4_yes8_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes8_RadioButton.Name = "PrmList_4_yes8_RadioButton"
        Me.PrmList_4_yes8_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes8_RadioButton.TabIndex = 15
        Me.PrmList_4_yes8_RadioButton.TabStop = True
        Me.PrmList_4_yes8_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no8_RadioButton
        '
        Me.PrmList_4_no8_RadioButton.AutoSize = True
        Me.PrmList_4_no8_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no8_RadioButton.Name = "PrmList_4_no8_RadioButton"
        Me.PrmList_4_no8_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no8_RadioButton.TabIndex = 16
        Me.PrmList_4_no8_RadioButton.TabStop = True
        Me.PrmList_4_no8_RadioButton.UseVisualStyleBackColor = True
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label38.Location = New System.Drawing.Point(50, 42)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(20, 16)
        Me.Label38.TabIndex = 109
        Me.Label38.Text = "否"
        '
        'Panel12
        '
        Me.Panel12.Controls.Add(Me.PrmList_4_yes11_RadioButton)
        Me.Panel12.Controls.Add(Me.PrmList_4_no11_RadioButton)
        Me.Panel12.Location = New System.Drawing.Point(14, 416)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(56, 20)
        Me.Panel12.TabIndex = 157
        '
        'PrmList_4_yes11_RadioButton
        '
        Me.PrmList_4_yes11_RadioButton.AutoSize = True
        Me.PrmList_4_yes11_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes11_RadioButton.Name = "PrmList_4_yes11_RadioButton"
        Me.PrmList_4_yes11_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes11_RadioButton.TabIndex = 21
        Me.PrmList_4_yes11_RadioButton.TabStop = True
        Me.PrmList_4_yes11_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no11_RadioButton
        '
        Me.PrmList_4_no11_RadioButton.AutoSize = True
        Me.PrmList_4_no11_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no11_RadioButton.Name = "PrmList_4_no11_RadioButton"
        Me.PrmList_4_no11_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no11_RadioButton.TabIndex = 22
        Me.PrmList_4_no11_RadioButton.TabStop = True
        Me.PrmList_4_no11_RadioButton.UseVisualStyleBackColor = True
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label39.Location = New System.Drawing.Point(88, 62)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(138, 16)
        Me.Label39.TabIndex = 112
        Me.Label39.Text = "1.可以手動、全自動運轉"
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.PrmList_4_yes4_RadioButton)
        Me.Panel6.Controls.Add(Me.PrmList_4_no4_RadioButton)
        Me.Panel6.Location = New System.Drawing.Point(14, 162)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(56, 20)
        Me.Panel6.TabIndex = 154
        '
        'PrmList_4_yes4_RadioButton
        '
        Me.PrmList_4_yes4_RadioButton.AutoSize = True
        Me.PrmList_4_yes4_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes4_RadioButton.Name = "PrmList_4_yes4_RadioButton"
        Me.PrmList_4_yes4_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes4_RadioButton.TabIndex = 7
        Me.PrmList_4_yes4_RadioButton.TabStop = True
        Me.PrmList_4_yes4_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no4_RadioButton
        '
        Me.PrmList_4_no4_RadioButton.AutoSize = True
        Me.PrmList_4_no4_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no4_RadioButton.Name = "PrmList_4_no4_RadioButton"
        Me.PrmList_4_no4_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no4_RadioButton.TabIndex = 8
        Me.PrmList_4_no4_RadioButton.TabStop = True
        Me.PrmList_4_no4_RadioButton.UseVisualStyleBackColor = True
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label40.Location = New System.Drawing.Point(88, 96)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(174, 16)
        Me.Label40.TabIndex = 115
        Me.Label40.Text = "2.入出力點和電氣設計内容一致"
        '
        'Panel13
        '
        Me.Panel13.Controls.Add(Me.PrmList_4_yes10_RadioButton)
        Me.Panel13.Controls.Add(Me.PrmList_4_no10_RadioButton)
        Me.Panel13.Location = New System.Drawing.Point(14, 374)
        Me.Panel13.Name = "Panel13"
        Me.Panel13.Size = New System.Drawing.Size(56, 20)
        Me.Panel13.TabIndex = 158
        '
        'PrmList_4_yes10_RadioButton
        '
        Me.PrmList_4_yes10_RadioButton.AutoSize = True
        Me.PrmList_4_yes10_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes10_RadioButton.Name = "PrmList_4_yes10_RadioButton"
        Me.PrmList_4_yes10_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes10_RadioButton.TabIndex = 19
        Me.PrmList_4_yes10_RadioButton.TabStop = True
        Me.PrmList_4_yes10_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no10_RadioButton
        '
        Me.PrmList_4_no10_RadioButton.AutoSize = True
        Me.PrmList_4_no10_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no10_RadioButton.Name = "PrmList_4_no10_RadioButton"
        Me.PrmList_4_no10_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no10_RadioButton.TabIndex = 20
        Me.PrmList_4_no10_RadioButton.TabStop = True
        Me.PrmList_4_no10_RadioButton.UseVisualStyleBackColor = True
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label41.Location = New System.Drawing.Point(88, 130)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(126, 16)
        Me.Label41.TabIndex = 118
        Me.Label41.Text = "3.變數有明示的初始化"
        '
        'Panel8
        '
        Me.Panel8.Controls.Add(Me.PrmList_4_yes7_RadioButton)
        Me.Panel8.Controls.Add(Me.PrmList_4_no7_RadioButton)
        Me.Panel8.Location = New System.Drawing.Point(14, 264)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(56, 20)
        Me.Panel8.TabIndex = 157
        '
        'PrmList_4_yes7_RadioButton
        '
        Me.PrmList_4_yes7_RadioButton.AutoSize = True
        Me.PrmList_4_yes7_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes7_RadioButton.Name = "PrmList_4_yes7_RadioButton"
        Me.PrmList_4_yes7_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes7_RadioButton.TabIndex = 13
        Me.PrmList_4_yes7_RadioButton.TabStop = True
        Me.PrmList_4_yes7_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no7_RadioButton
        '
        Me.PrmList_4_no7_RadioButton.AutoSize = True
        Me.PrmList_4_no7_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no7_RadioButton.Name = "PrmList_4_no7_RadioButton"
        Me.PrmList_4_no7_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no7_RadioButton.TabIndex = 14
        Me.PrmList_4_no7_RadioButton.TabStop = True
        Me.PrmList_4_no7_RadioButton.UseVisualStyleBackColor = True
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label42.Location = New System.Drawing.Point(88, 164)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(353, 16)
        Me.Label42.TabIndex = 121
        Me.Label42.Text = "4.內容是否有出現沒有OTHER的CASE（若沒有的話，確定可以）"
        '
        'Panel14
        '
        Me.Panel14.Controls.Add(Me.PrmList_4_yes9_RadioButton)
        Me.Panel14.Controls.Add(Me.PrmList_4_no9_RadioButton)
        Me.Panel14.Location = New System.Drawing.Point(14, 332)
        Me.Panel14.Name = "Panel14"
        Me.Panel14.Size = New System.Drawing.Size(56, 20)
        Me.Panel14.TabIndex = 155
        '
        'PrmList_4_yes9_RadioButton
        '
        Me.PrmList_4_yes9_RadioButton.AutoSize = True
        Me.PrmList_4_yes9_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes9_RadioButton.Name = "PrmList_4_yes9_RadioButton"
        Me.PrmList_4_yes9_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes9_RadioButton.TabIndex = 17
        Me.PrmList_4_yes9_RadioButton.TabStop = True
        Me.PrmList_4_yes9_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no9_RadioButton
        '
        Me.PrmList_4_no9_RadioButton.AutoSize = True
        Me.PrmList_4_no9_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no9_RadioButton.Name = "PrmList_4_no9_RadioButton"
        Me.PrmList_4_no9_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no9_RadioButton.TabIndex = 18
        Me.PrmList_4_no9_RadioButton.TabStop = True
        Me.PrmList_4_no9_RadioButton.UseVisualStyleBackColor = True
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label43.Location = New System.Drawing.Point(88, 198)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(318, 16)
        Me.Label43.TabIndex = 124
        Me.Label43.Text = "5.內容是否有出現沒有ELSE的IF（若沒有的話，確定可以）"
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.PrmList_4_yes3_RadioButton)
        Me.Panel5.Controls.Add(Me.PrmList_4_no3_RadioButton)
        Me.Panel5.Location = New System.Drawing.Point(14, 128)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(56, 20)
        Me.Panel5.TabIndex = 154
        '
        'PrmList_4_yes3_RadioButton
        '
        Me.PrmList_4_yes3_RadioButton.AutoSize = True
        Me.PrmList_4_yes3_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes3_RadioButton.Name = "PrmList_4_yes3_RadioButton"
        Me.PrmList_4_yes3_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes3_RadioButton.TabIndex = 5
        Me.PrmList_4_yes3_RadioButton.TabStop = True
        Me.PrmList_4_yes3_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no3_RadioButton
        '
        Me.PrmList_4_no3_RadioButton.AutoSize = True
        Me.PrmList_4_no3_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no3_RadioButton.Name = "PrmList_4_no3_RadioButton"
        Me.PrmList_4_no3_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no3_RadioButton.TabIndex = 6
        Me.PrmList_4_no3_RadioButton.TabStop = True
        Me.PrmList_4_no3_RadioButton.UseVisualStyleBackColor = True
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label48.Location = New System.Drawing.Point(88, 232)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(255, 16)
        Me.Label48.TabIndex = 127
        Me.Label48.Text = "6.使用反覆指令時，有沒有可能造成無限LOOP"
        '
        'Panel9
        '
        Me.Panel9.Controls.Add(Me.PrmList_4_yes6_RadioButton)
        Me.Panel9.Controls.Add(Me.PrmList_4_no6_RadioButton)
        Me.Panel9.Location = New System.Drawing.Point(14, 230)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(56, 20)
        Me.Panel9.TabIndex = 158
        '
        'PrmList_4_yes6_RadioButton
        '
        Me.PrmList_4_yes6_RadioButton.AutoSize = True
        Me.PrmList_4_yes6_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes6_RadioButton.Name = "PrmList_4_yes6_RadioButton"
        Me.PrmList_4_yes6_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes6_RadioButton.TabIndex = 11
        Me.PrmList_4_yes6_RadioButton.TabStop = True
        Me.PrmList_4_yes6_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no6_RadioButton
        '
        Me.PrmList_4_no6_RadioButton.AutoSize = True
        Me.PrmList_4_no6_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no6_RadioButton.Name = "PrmList_4_no6_RadioButton"
        Me.PrmList_4_no6_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no6_RadioButton.TabIndex = 12
        Me.PrmList_4_no6_RadioButton.TabStop = True
        Me.PrmList_4_no6_RadioButton.UseVisualStyleBackColor = True
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label47.Location = New System.Drawing.Point(88, 266)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(390, 16)
        Me.Label47.TabIndex = 130
        Me.Label47.Text = "7.所有的配列參照、所增加的文字、是不是在對應的次元所定義的範圍內"
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.PrmList_4_yes2_RadioButton)
        Me.Panel4.Controls.Add(Me.PrmList_4_no2_RadioButton)
        Me.Panel4.Location = New System.Drawing.Point(14, 94)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(56, 20)
        Me.Panel4.TabIndex = 154
        '
        'PrmList_4_yes2_RadioButton
        '
        Me.PrmList_4_yes2_RadioButton.AutoSize = True
        Me.PrmList_4_yes2_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes2_RadioButton.Name = "PrmList_4_yes2_RadioButton"
        Me.PrmList_4_yes2_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes2_RadioButton.TabIndex = 3
        Me.PrmList_4_yes2_RadioButton.TabStop = True
        Me.PrmList_4_yes2_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no2_RadioButton
        '
        Me.PrmList_4_no2_RadioButton.AutoSize = True
        Me.PrmList_4_no2_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no2_RadioButton.Name = "PrmList_4_no2_RadioButton"
        Me.PrmList_4_no2_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no2_RadioButton.TabIndex = 4
        Me.PrmList_4_no2_RadioButton.TabStop = True
        Me.PrmList_4_no2_RadioButton.UseVisualStyleBackColor = True
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label46.Location = New System.Drawing.Point(88, 300)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(358, 16)
        Me.Label46.TabIndex = 133
        Me.Label46.Text = "8.在CASTING時，確認沒有OVERFLOW或UNDERFLOW的可能性"
        '
        'Panel10
        '
        Me.Panel10.Controls.Add(Me.PrmList_4_yes5_RadioButton)
        Me.Panel10.Controls.Add(Me.PrmList_4_no5_RadioButton)
        Me.Panel10.Location = New System.Drawing.Point(14, 196)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(56, 20)
        Me.Panel10.TabIndex = 155
        '
        'PrmList_4_yes5_RadioButton
        '
        Me.PrmList_4_yes5_RadioButton.AutoSize = True
        Me.PrmList_4_yes5_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes5_RadioButton.Name = "PrmList_4_yes5_RadioButton"
        Me.PrmList_4_yes5_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes5_RadioButton.TabIndex = 9
        Me.PrmList_4_yes5_RadioButton.TabStop = True
        Me.PrmList_4_yes5_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no5_RadioButton
        '
        Me.PrmList_4_no5_RadioButton.AutoSize = True
        Me.PrmList_4_no5_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no5_RadioButton.Name = "PrmList_4_no5_RadioButton"
        Me.PrmList_4_no5_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no5_RadioButton.TabIndex = 10
        Me.PrmList_4_no5_RadioButton.TabStop = True
        Me.PrmList_4_no5_RadioButton.UseVisualStyleBackColor = True
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label45.Location = New System.Drawing.Point(88, 334)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(145, 16)
        Me.Label45.TabIndex = 136
        Me.Label45.Text = "9.確定沒有用0來除的式子"
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.PrmList_4_yes1_RadioButton)
        Me.Panel3.Controls.Add(Me.PrmList_4_no1_RadioButton)
        Me.Panel3.Location = New System.Drawing.Point(14, 60)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(56, 20)
        Me.Panel3.TabIndex = 153
        '
        'PrmList_4_yes1_RadioButton
        '
        Me.PrmList_4_yes1_RadioButton.AutoSize = True
        Me.PrmList_4_yes1_RadioButton.Location = New System.Drawing.Point(3, 3)
        Me.PrmList_4_yes1_RadioButton.Name = "PrmList_4_yes1_RadioButton"
        Me.PrmList_4_yes1_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_yes1_RadioButton.TabIndex = 1
        Me.PrmList_4_yes1_RadioButton.TabStop = True
        Me.PrmList_4_yes1_RadioButton.UseVisualStyleBackColor = True
        '
        'PrmList_4_no1_RadioButton
        '
        Me.PrmList_4_no1_RadioButton.AutoSize = True
        Me.PrmList_4_no1_RadioButton.Location = New System.Drawing.Point(39, 3)
        Me.PrmList_4_no1_RadioButton.Name = "PrmList_4_no1_RadioButton"
        Me.PrmList_4_no1_RadioButton.Size = New System.Drawing.Size(14, 13)
        Me.PrmList_4_no1_RadioButton.TabIndex = 2
        Me.PrmList_4_no1_RadioButton.TabStop = True
        Me.PrmList_4_no1_RadioButton.UseVisualStyleBackColor = True
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label44.Location = New System.Drawing.Point(88, 368)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(440, 32)
        Me.Label44.TabIndex = 139
        Me.Label44.Text = "10.包含2個以上的運算子的式子中、實行的運算子是否正確地如所期待的方式執行" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（確認有理解運算子的優先順序）"
        '
        'Label51
        '
        Me.Label51.AutoSize = True
        Me.Label51.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label51.Location = New System.Drawing.Point(88, 418)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(351, 16)
        Me.Label51.TabIndex = 140
        Me.Label51.Text = "11.變數有分配ADDRESS時、是否有其他的變數和ADDRESS重複"
        '
        'PrmList_4_content12_TextBox
        '
        Me.PrmList_4_content12_TextBox.Location = New System.Drawing.Point(91, 505)
        Me.PrmList_4_content12_TextBox.Multiline = True
        Me.PrmList_4_content12_TextBox.Name = "PrmList_4_content12_TextBox"
        Me.PrmList_4_content12_TextBox.Size = New System.Drawing.Size(425, 79)
        Me.PrmList_4_content12_TextBox.TabIndex = 25
        '
        'Label50
        '
        Me.Label50.AutoSize = True
        Me.Label50.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label50.Location = New System.Drawing.Point(88, 452)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(157, 16)
        Me.Label50.TabIndex = 141
        Me.Label50.Text = "12.有實現客戶所要求的仕樣"
        '
        'Label49
        '
        Me.Label49.AutoSize = True
        Me.Label49.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label49.Location = New System.Drawing.Point(88, 486)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(62, 16)
        Me.Label49.TabIndex = 142
        Me.Label49.Text = "測試內容 :"
        '
        'Use_Program_CheckBox
        '
        Me.Use_Program_CheckBox.AutoSize = True
        Me.Use_Program_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_Program_CheckBox.Name = "Use_Program_CheckBox"
        Me.Use_Program_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Use_Program_CheckBox.TabIndex = 161
        Me.Use_Program_CheckBox.UseVisualStyleBackColor = True
        '
        'CheckList
        '
        Me.CheckList.Controls.Add(Me.CheckList_GroupBox)
        Me.CheckList.Controls.Add(Me.Use_ChkList_CheckBox)
        Me.CheckList.Location = New System.Drawing.Point(4, 25)
        Me.CheckList.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CheckList.Name = "CheckList"
        Me.CheckList.Size = New System.Drawing.Size(664, 584)
        Me.CheckList.TabIndex = 4
        Me.CheckList.Text = "CheckList"
        Me.CheckList.UseVisualStyleBackColor = True
        '
        'CheckList_GroupBox
        '
        Me.CheckList_GroupBox.Controls.Add(Me.TabControl1)
        Me.CheckList_GroupBox.Controls.Add(Me.Label10)
        Me.CheckList_GroupBox.Controls.Add(Me.ChkList_PaSheet_CheckBox)
        Me.CheckList_GroupBox.Controls.Add(Me.Label11)
        Me.CheckList_GroupBox.Controls.Add(Me.ChkList_OS_CheckBox)
        Me.CheckList_GroupBox.Controls.Add(Me.ChkList_Elec_DateTimePicker)
        Me.CheckList_GroupBox.Controls.Add(Me.Label12)
        Me.CheckList_GroupBox.Controls.Add(Me.ChkList_Confirm_DateTimePicker)
        Me.CheckList_GroupBox.Controls.Add(Me.ChkList_Confirm_CheckBox)
        Me.CheckList_GroupBox.Controls.Add(Me.ChkList_OS_DateTimePicker)
        Me.CheckList_GroupBox.Controls.Add(Me.Label13)
        Me.CheckList_GroupBox.Controls.Add(Me.ChkList_PaSheet_DateTimePicker)
        Me.CheckList_GroupBox.Controls.Add(Me.ChkList_Elec_CheckBox)
        Me.CheckList_GroupBox.Enabled = False
        Me.CheckList_GroupBox.Location = New System.Drawing.Point(3, 24)
        Me.CheckList_GroupBox.Name = "CheckList_GroupBox"
        Me.CheckList_GroupBox.Size = New System.Drawing.Size(658, 555)
        Me.CheckList_GroupBox.TabIndex = 59
        Me.CheckList_GroupBox.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(6, 132)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(646, 417)
        Me.TabControl1.TabIndex = 61
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.CheckList_FlowLayoutPanel)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(638, 388)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Page1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'CheckList_FlowLayoutPanel
        '
        Me.CheckList_FlowLayoutPanel.AutoScroll = True
        Me.CheckList_FlowLayoutPanel.Controls.Add(Me.ChkList_1_Panel)
        Me.CheckList_FlowLayoutPanel.Controls.Add(Me.ChkList_2_Panel)
        Me.CheckList_FlowLayoutPanel.Controls.Add(Me.ChkList_3_Panel)
        Me.CheckList_FlowLayoutPanel.Controls.Add(Me.Button9)
        Me.CheckList_FlowLayoutPanel.Controls.Add(Me.Button6)
        Me.CheckList_FlowLayoutPanel.Location = New System.Drawing.Point(6, 7)
        Me.CheckList_FlowLayoutPanel.Name = "CheckList_FlowLayoutPanel"
        Me.CheckList_FlowLayoutPanel.Size = New System.Drawing.Size(626, 410)
        Me.CheckList_FlowLayoutPanel.TabIndex = 60
        '
        'ChkList_1_Panel
        '
        Me.ChkList_1_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_1_Panel.Controls.Add(Me.ChkList_1_no_RadioButton)
        Me.ChkList_1_Panel.Controls.Add(Me.ChkList_1_yes_RadioButton)
        Me.ChkList_1_Panel.Controls.Add(Me.ChkList_1_yes_Content_TextBox)
        Me.ChkList_1_Panel.Controls.Add(Me.ChkList_1_yes_result_TextBox)
        Me.ChkList_1_Panel.Controls.Add(Me.Label15)
        Me.ChkList_1_Panel.Controls.Add(Me.Label14)
        Me.ChkList_1_Panel.Location = New System.Drawing.Point(3, 3)
        Me.ChkList_1_Panel.Name = "ChkList_1_Panel"
        Me.ChkList_1_Panel.Size = New System.Drawing.Size(590, 100)
        Me.ChkList_1_Panel.TabIndex = 154
        '
        'ChkList_1_no_RadioButton
        '
        Me.ChkList_1_no_RadioButton.AutoSize = True
        Me.ChkList_1_no_RadioButton.Location = New System.Drawing.Point(19, 19)
        Me.ChkList_1_no_RadioButton.Name = "ChkList_1_no_RadioButton"
        Me.ChkList_1_no_RadioButton.Size = New System.Drawing.Size(38, 20)
        Me.ChkList_1_no_RadioButton.TabIndex = 5
        Me.ChkList_1_no_RadioButton.TabStop = True
        Me.ChkList_1_no_RadioButton.Text = "無"
        Me.ChkList_1_no_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_1_yes_RadioButton
        '
        Me.ChkList_1_yes_RadioButton.AutoSize = True
        Me.ChkList_1_yes_RadioButton.Location = New System.Drawing.Point(19, 42)
        Me.ChkList_1_yes_RadioButton.Name = "ChkList_1_yes_RadioButton"
        Me.ChkList_1_yes_RadioButton.Size = New System.Drawing.Size(90, 20)
        Me.ChkList_1_yes_RadioButton.TabIndex = 6
        Me.ChkList_1_yes_RadioButton.TabStop = True
        Me.ChkList_1_yes_RadioButton.Text = "有(討論內容"
        Me.ChkList_1_yes_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_1_yes_Content_TextBox
        '
        Me.ChkList_1_yes_Content_TextBox.Location = New System.Drawing.Point(113, 41)
        Me.ChkList_1_yes_Content_TextBox.Name = "ChkList_1_yes_Content_TextBox"
        Me.ChkList_1_yes_Content_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_1_yes_Content_TextBox.TabIndex = 7
        '
        'ChkList_1_yes_result_TextBox
        '
        Me.ChkList_1_yes_result_TextBox.Location = New System.Drawing.Point(113, 70)
        Me.ChkList_1_yes_result_TextBox.Name = "ChkList_1_yes_result_TextBox"
        Me.ChkList_1_yes_result_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_1_yes_result_TextBox.TabIndex = 8
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label15.Location = New System.Drawing.Point(41, 73)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(42, 16)
        Me.Label15.TabIndex = 38
        Me.Label15.Text = "(結果 :"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label14.Location = New System.Drawing.Point(3, 1)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(479, 16)
        Me.Label14.TabIndex = 34
        Me.Label14.Text = "1.主仕樣、工直仕樣、確認圖、OS中有沒有不清楚或有問題之處【確認事項都解決了嗎】"
        '
        'ChkList_2_Panel
        '
        Me.ChkList_2_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_2_Panel.Controls.Add(Me.ChkList_2_yes_RadioButton)
        Me.ChkList_2_Panel.Controls.Add(Me.ChkList_2_no_RadioButton)
        Me.ChkList_2_Panel.Controls.Add(Me.Label17)
        Me.ChkList_2_Panel.Controls.Add(Me.Label16)
        Me.ChkList_2_Panel.Controls.Add(Me.ChkList_2_yes_Result_TextBox)
        Me.ChkList_2_Panel.Controls.Add(Me.ChkList_2_yes_Content_TextBox)
        Me.ChkList_2_Panel.Location = New System.Drawing.Point(3, 109)
        Me.ChkList_2_Panel.Name = "ChkList_2_Panel"
        Me.ChkList_2_Panel.Size = New System.Drawing.Size(590, 110)
        Me.ChkList_2_Panel.TabIndex = 155
        '
        'ChkList_2_yes_RadioButton
        '
        Me.ChkList_2_yes_RadioButton.AutoSize = True
        Me.ChkList_2_yes_RadioButton.Location = New System.Drawing.Point(19, 45)
        Me.ChkList_2_yes_RadioButton.Name = "ChkList_2_yes_RadioButton"
        Me.ChkList_2_yes_RadioButton.Size = New System.Drawing.Size(96, 20)
        Me.ChkList_2_yes_RadioButton.TabIndex = 10
        Me.ChkList_2_yes_RadioButton.TabStop = True
        Me.ChkList_2_yes_RadioButton.Text = "有(指出內容 :"
        Me.ChkList_2_yes_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_2_no_RadioButton
        '
        Me.ChkList_2_no_RadioButton.AutoSize = True
        Me.ChkList_2_no_RadioButton.Location = New System.Drawing.Point(19, 19)
        Me.ChkList_2_no_RadioButton.Name = "ChkList_2_no_RadioButton"
        Me.ChkList_2_no_RadioButton.Size = New System.Drawing.Size(38, 20)
        Me.ChkList_2_no_RadioButton.TabIndex = 9
        Me.ChkList_2_no_RadioButton.TabStop = True
        Me.ChkList_2_no_RadioButton.Text = "無"
        Me.ChkList_2_no_RadioButton.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label17.Location = New System.Drawing.Point(7, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(330, 16)
        Me.Label17.TabIndex = 40
        Me.Label17.Text = "2.有沒有可能會在法規、安全、機能面上會發生問題的仕樣？"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label16.Location = New System.Drawing.Point(47, 74)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(42, 16)
        Me.Label16.TabIndex = 44
        Me.Label16.Text = "(結果 :"
        '
        'ChkList_2_yes_Result_TextBox
        '
        Me.ChkList_2_yes_Result_TextBox.Location = New System.Drawing.Point(113, 71)
        Me.ChkList_2_yes_Result_TextBox.Name = "ChkList_2_yes_Result_TextBox"
        Me.ChkList_2_yes_Result_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_2_yes_Result_TextBox.TabIndex = 12
        '
        'ChkList_2_yes_Content_TextBox
        '
        Me.ChkList_2_yes_Content_TextBox.Location = New System.Drawing.Point(113, 44)
        Me.ChkList_2_yes_Content_TextBox.Name = "ChkList_2_yes_Content_TextBox"
        Me.ChkList_2_yes_Content_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_2_yes_Content_TextBox.TabIndex = 11
        '
        'ChkList_3_Panel
        '
        Me.ChkList_3_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_3_Panel.Controls.Add(Me.ChkList_3_yes_RadioButton)
        Me.ChkList_3_Panel.Controls.Add(Me.ChkList_3_no_RadioButton)
        Me.ChkList_3_Panel.Controls.Add(Me.Label19)
        Me.ChkList_3_Panel.Controls.Add(Me.Label18)
        Me.ChkList_3_Panel.Controls.Add(Me.ChkList_3_yes_Man_TextBox)
        Me.ChkList_3_Panel.Controls.Add(Me.Label20)
        Me.ChkList_3_Panel.Controls.Add(Me.ChkList_3_yes_Content_TextBox)
        Me.ChkList_3_Panel.Controls.Add(Me.Label21)
        Me.ChkList_3_Panel.Controls.Add(Me.ChkList_3_yes_Result_TextBox)
        Me.ChkList_3_Panel.Location = New System.Drawing.Point(3, 225)
        Me.ChkList_3_Panel.Name = "ChkList_3_Panel"
        Me.ChkList_3_Panel.Size = New System.Drawing.Size(590, 165)
        Me.ChkList_3_Panel.TabIndex = 156
        '
        'ChkList_3_yes_RadioButton
        '
        Me.ChkList_3_yes_RadioButton.AutoSize = True
        Me.ChkList_3_yes_RadioButton.Location = New System.Drawing.Point(19, 45)
        Me.ChkList_3_yes_RadioButton.Name = "ChkList_3_yes_RadioButton"
        Me.ChkList_3_yes_RadioButton.Size = New System.Drawing.Size(154, 20)
        Me.ChkList_3_yes_RadioButton.TabIndex = 14
        Me.ChkList_3_yes_RadioButton.TabStop = True
        Me.ChkList_3_yes_RadioButton.Text = "有(和電氣設計擔當確認)"
        Me.ChkList_3_yes_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_3_no_RadioButton
        '
        Me.ChkList_3_no_RadioButton.AutoSize = True
        Me.ChkList_3_no_RadioButton.Location = New System.Drawing.Point(19, 19)
        Me.ChkList_3_no_RadioButton.Name = "ChkList_3_no_RadioButton"
        Me.ChkList_3_no_RadioButton.Size = New System.Drawing.Size(38, 20)
        Me.ChkList_3_no_RadioButton.TabIndex = 13
        Me.ChkList_3_no_RadioButton.TabStop = True
        Me.ChkList_3_no_RadioButton.Text = "無"
        Me.ChkList_3_no_RadioButton.UseVisualStyleBackColor = True
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label19.Location = New System.Drawing.Point(7, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(362, 16)
        Me.Label19.TabIndex = 46
        Me.Label19.Text = "3.氣圖面上有沒有不清楚之處？和電氣設計GR.間的介面是否有確實"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label18.Location = New System.Drawing.Point(53, 74)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(50, 16)
        Me.Label18.TabIndex = 50
        Me.Label18.Text = "討論者 :"
        '
        'ChkList_3_yes_Man_TextBox
        '
        Me.ChkList_3_yes_Man_TextBox.Location = New System.Drawing.Point(113, 71)
        Me.ChkList_3_yes_Man_TextBox.Name = "ChkList_3_yes_Man_TextBox"
        Me.ChkList_3_yes_Man_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_3_yes_Man_TextBox.TabIndex = 15
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label20.Location = New System.Drawing.Point(53, 103)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(38, 16)
        Me.Label20.TabIndex = 52
        Me.Label20.Text = "內容 :"
        '
        'ChkList_3_yes_Content_TextBox
        '
        Me.ChkList_3_yes_Content_TextBox.Location = New System.Drawing.Point(113, 100)
        Me.ChkList_3_yes_Content_TextBox.Name = "ChkList_3_yes_Content_TextBox"
        Me.ChkList_3_yes_Content_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_3_yes_Content_TextBox.TabIndex = 16
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label21.Location = New System.Drawing.Point(53, 132)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(38, 16)
        Me.Label21.TabIndex = 54
        Me.Label21.Text = "結論 :"
        '
        'ChkList_3_yes_Result_TextBox
        '
        Me.ChkList_3_yes_Result_TextBox.Location = New System.Drawing.Point(113, 129)
        Me.ChkList_3_yes_Result_TextBox.Name = "ChkList_3_yes_Result_TextBox"
        Me.ChkList_3_yes_Result_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_3_yes_Result_TextBox.TabIndex = 17
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(3, 396)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(600, 10)
        Me.Button9.TabIndex = 170
        Me.Button9.Text = "Button9"
        Me.Button9.UseVisualStyleBackColor = True
        Me.Button9.Visible = False
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(3, 412)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(615, 10)
        Me.Button6.TabIndex = 167
        Me.Button6.Text = "Button6"
        Me.Button6.UseVisualStyleBackColor = True
        Me.Button6.Visible = False
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.CheckList2_FlowLayoutPanel)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(638, 388)
        Me.TabPage3.TabIndex = 1
        Me.TabPage3.Text = "Page2"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'CheckList2_FlowLayoutPanel
        '
        Me.CheckList2_FlowLayoutPanel.Controls.Add(Me.ChkList_6_Panel)
        Me.CheckList2_FlowLayoutPanel.Controls.Add(Me.ChkList_4_Panel)
        Me.CheckList2_FlowLayoutPanel.Controls.Add(Me.ChkList_5_Panel)
        Me.CheckList2_FlowLayoutPanel.Location = New System.Drawing.Point(6, 6)
        Me.CheckList2_FlowLayoutPanel.Name = "CheckList2_FlowLayoutPanel"
        Me.CheckList2_FlowLayoutPanel.Size = New System.Drawing.Size(626, 376)
        Me.CheckList2_FlowLayoutPanel.TabIndex = 62
        '
        'ChkList_6_Panel
        '
        Me.ChkList_6_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_6_Panel.Controls.Add(Me.Panel24)
        Me.ChkList_6_Panel.Controls.Add(Me.ChkList_6_no_RadioButton)
        Me.ChkList_6_Panel.Controls.Add(Me.ChkList_6_yes_RadioButton)
        Me.ChkList_6_Panel.Controls.Add(Me.Label27)
        Me.ChkList_6_Panel.Location = New System.Drawing.Point(3, 3)
        Me.ChkList_6_Panel.Name = "ChkList_6_Panel"
        Me.ChkList_6_Panel.Size = New System.Drawing.Size(590, 142)
        Me.ChkList_6_Panel.TabIndex = 160
        '
        'Panel24
        '
        Me.Panel24.Controls.Add(Me.ChkList_6_yesItem_RadioButton)
        Me.Panel24.Controls.Add(Me.ChkList_6_yesChk_RadioButton)
        Me.Panel24.Controls.Add(Me.ChkList_6_yes_Content_TextBox)
        Me.Panel24.Location = New System.Drawing.Point(41, 45)
        Me.Panel24.Name = "Panel24"
        Me.Panel24.Size = New System.Drawing.Size(422, 57)
        Me.Panel24.TabIndex = 157
        '
        'ChkList_6_yesItem_RadioButton
        '
        Me.ChkList_6_yesItem_RadioButton.AutoSize = True
        Me.ChkList_6_yesItem_RadioButton.Location = New System.Drawing.Point(5, 29)
        Me.ChkList_6_yesItem_RadioButton.Name = "ChkList_6_yesItem_RadioButton"
        Me.ChkList_6_yesItem_RadioButton.Size = New System.Drawing.Size(80, 20)
        Me.ChkList_6_yesItem_RadioButton.TabIndex = 25
        Me.ChkList_6_yesItem_RadioButton.TabStop = True
        Me.ChkList_6_yesItem_RadioButton.Text = "檢驗項目 :"
        Me.ChkList_6_yesItem_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_6_yesChk_RadioButton
        '
        Me.ChkList_6_yesChk_RadioButton.AutoSize = True
        Me.ChkList_6_yesChk_RadioButton.Location = New System.Drawing.Point(5, 3)
        Me.ChkList_6_yesChk_RadioButton.Name = "ChkList_6_yesChk_RadioButton"
        Me.ChkList_6_yesChk_RadioButton.Size = New System.Drawing.Size(177, 20)
        Me.ChkList_6_yesChk_RadioButton.TabIndex = 24
        Me.ChkList_6_yesChk_RadioButton.TabStop = True
        Me.ChkList_6_yesChk_RadioButton.Text = "根據程式變更CHECK SHEET"
        Me.ChkList_6_yesChk_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_6_yes_Content_TextBox
        '
        Me.ChkList_6_yes_Content_TextBox.Location = New System.Drawing.Point(92, 29)
        Me.ChkList_6_yes_Content_TextBox.Name = "ChkList_6_yes_Content_TextBox"
        Me.ChkList_6_yes_Content_TextBox.Size = New System.Drawing.Size(320, 23)
        Me.ChkList_6_yes_Content_TextBox.TabIndex = 26
        '
        'ChkList_6_no_RadioButton
        '
        Me.ChkList_6_no_RadioButton.AutoSize = True
        Me.ChkList_6_no_RadioButton.Location = New System.Drawing.Point(19, 106)
        Me.ChkList_6_no_RadioButton.Name = "ChkList_6_no_RadioButton"
        Me.ChkList_6_no_RadioButton.Size = New System.Drawing.Size(221, 20)
        Me.ChkList_6_no_RadioButton.TabIndex = 27
        Me.ChkList_6_no_RadioButton.TabStop = True
        Me.ChkList_6_no_RadioButton.Text = " 無（因為是類似設計所以沒有必要）"
        Me.ChkList_6_no_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_6_yes_RadioButton
        '
        Me.ChkList_6_yes_RadioButton.AutoSize = True
        Me.ChkList_6_yes_RadioButton.Location = New System.Drawing.Point(19, 19)
        Me.ChkList_6_yes_RadioButton.Name = "ChkList_6_yes_RadioButton"
        Me.ChkList_6_yes_RadioButton.Size = New System.Drawing.Size(38, 20)
        Me.ChkList_6_yes_RadioButton.TabIndex = 23
        Me.ChkList_6_yes_RadioButton.TabStop = True
        Me.ChkList_6_yes_RadioButton.Text = "有"
        Me.ChkList_6_yes_RadioButton.UseVisualStyleBackColor = True
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label27.Location = New System.Drawing.Point(7, 0)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(126, 16)
        Me.Label27.TabIndex = 79
        Me.Label27.Text = "6.有執行動作確認嗎？"
        '
        'ChkList_4_Panel
        '
        Me.ChkList_4_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_4_Panel.Controls.Add(Me.ChkList_4_ObjName_TextBox)
        Me.ChkList_4_Panel.Controls.Add(Me.ChkList_4_ObjBase_TextBox)
        Me.ChkList_4_Panel.Controls.Add(Me.Label22)
        Me.ChkList_4_Panel.Controls.Add(Me.ChkList_4_SV_TextBox)
        Me.ChkList_4_Panel.Controls.Add(Me.Label23)
        Me.ChkList_4_Panel.Controls.Add(Me.Label24)
        Me.ChkList_4_Panel.Controls.Add(Me.Label26)
        Me.ChkList_4_Panel.Controls.Add(Me.Label25)
        Me.ChkList_4_Panel.Controls.Add(Me.ChkList_4_SVBase_TextBox)
        Me.ChkList_4_Panel.Location = New System.Drawing.Point(3, 151)
        Me.ChkList_4_Panel.Name = "ChkList_4_Panel"
        Me.ChkList_4_Panel.Size = New System.Drawing.Size(590, 91)
        Me.ChkList_4_Panel.TabIndex = 161
        '
        'ChkList_4_ObjName_TextBox
        '
        Me.ChkList_4_ObjName_TextBox.Location = New System.Drawing.Point(19, 50)
        Me.ChkList_4_ObjName_TextBox.Name = "ChkList_4_ObjName_TextBox"
        Me.ChkList_4_ObjName_TextBox.Size = New System.Drawing.Size(120, 23)
        Me.ChkList_4_ObjName_TextBox.TabIndex = 59
        Me.ChkList_4_ObjName_TextBox.Visible = False
        '
        'ChkList_4_ObjBase_TextBox
        '
        Me.ChkList_4_ObjBase_TextBox.Location = New System.Drawing.Point(160, 50)
        Me.ChkList_4_ObjBase_TextBox.Name = "ChkList_4_ObjBase_TextBox"
        Me.ChkList_4_ObjBase_TextBox.Size = New System.Drawing.Size(120, 23)
        Me.ChkList_4_ObjBase_TextBox.TabIndex = 66
        Me.ChkList_4_ObjBase_TextBox.Visible = False
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label22.Location = New System.Drawing.Point(7, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(213, 16)
        Me.Label22.TabIndex = 56
        Me.Label22.Text = "4.適用OBJECT(MMIC分頁會自動帶入)"
        '
        'ChkList_4_SV_TextBox
        '
        Me.ChkList_4_SV_TextBox.Location = New System.Drawing.Point(295, 50)
        Me.ChkList_4_SV_TextBox.Name = "ChkList_4_SV_TextBox"
        Me.ChkList_4_SV_TextBox.Size = New System.Drawing.Size(130, 23)
        Me.ChkList_4_SV_TextBox.TabIndex = 61
        Me.ChkList_4_SV_TextBox.Visible = False
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label23.Location = New System.Drawing.Point(19, 27)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(49, 16)
        Me.Label23.TabIndex = 62
        Me.Label23.Text = "MMIC :"
        Me.Label23.Visible = False
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label24.Location = New System.Drawing.Point(295, 27)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(29, 16)
        Me.Label24.TabIndex = 63
        Me.Label24.Text = "SV :"
        Me.Label24.Visible = False
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label26.Location = New System.Drawing.Point(160, 27)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(43, 16)
        Me.Label26.TabIndex = 64
        Me.Label26.Text = "BASE :"
        Me.Label26.Visible = False
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label25.Location = New System.Drawing.Point(440, 27)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(43, 16)
        Me.Label25.TabIndex = 65
        Me.Label25.Text = "BASE :"
        Me.Label25.Visible = False
        '
        'ChkList_4_SVBase_TextBox
        '
        Me.ChkList_4_SVBase_TextBox.Location = New System.Drawing.Point(440, 50)
        Me.ChkList_4_SVBase_TextBox.Name = "ChkList_4_SVBase_TextBox"
        Me.ChkList_4_SVBase_TextBox.Size = New System.Drawing.Size(130, 23)
        Me.ChkList_4_SVBase_TextBox.TabIndex = 67
        Me.ChkList_4_SVBase_TextBox.Visible = False
        '
        'ChkList_5_Panel
        '
        Me.ChkList_5_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_5_Panel.Controls.Add(Me.ChkList_5_nstd_RadioButton)
        Me.ChkList_5_Panel.Controls.Add(Me.ChkList_5_std_RadioButton)
        Me.ChkList_5_Panel.Controls.Add(Me.ChkList_5_no_RadioButton)
        Me.ChkList_5_Panel.Controls.Add(Me.Label30)
        Me.ChkList_5_Panel.Controls.Add(Me.ChkList_5_std_Content_TextBox)
        Me.ChkList_5_Panel.Controls.Add(Me.ChkList_5_nstd_Content_TextBox)
        Me.ChkList_5_Panel.Location = New System.Drawing.Point(3, 248)
        Me.ChkList_5_Panel.Name = "ChkList_5_Panel"
        Me.ChkList_5_Panel.Size = New System.Drawing.Size(590, 110)
        Me.ChkList_5_Panel.TabIndex = 162
        '
        'ChkList_5_nstd_RadioButton
        '
        Me.ChkList_5_nstd_RadioButton.AutoSize = True
        Me.ChkList_5_nstd_RadioButton.Location = New System.Drawing.Point(19, 75)
        Me.ChkList_5_nstd_RadioButton.Name = "ChkList_5_nstd_RadioButton"
        Me.ChkList_5_nstd_RadioButton.Size = New System.Drawing.Size(44, 20)
        Me.ChkList_5_nstd_RadioButton.TabIndex = 20
        Me.ChkList_5_nstd_RadioButton.TabStop = True
        Me.ChkList_5_nstd_RadioButton.Text = "工 :"
        Me.ChkList_5_nstd_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_5_std_RadioButton
        '
        Me.ChkList_5_std_RadioButton.AutoSize = True
        Me.ChkList_5_std_RadioButton.Location = New System.Drawing.Point(19, 47)
        Me.ChkList_5_std_RadioButton.Name = "ChkList_5_std_RadioButton"
        Me.ChkList_5_std_RadioButton.Size = New System.Drawing.Size(44, 20)
        Me.ChkList_5_std_RadioButton.TabIndex = 19
        Me.ChkList_5_std_RadioButton.TabStop = True
        Me.ChkList_5_std_RadioButton.Text = "標 :"
        Me.ChkList_5_std_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_5_no_RadioButton
        '
        Me.ChkList_5_no_RadioButton.AutoSize = True
        Me.ChkList_5_no_RadioButton.Location = New System.Drawing.Point(19, 19)
        Me.ChkList_5_no_RadioButton.Name = "ChkList_5_no_RadioButton"
        Me.ChkList_5_no_RadioButton.Size = New System.Drawing.Size(38, 20)
        Me.ChkList_5_no_RadioButton.TabIndex = 18
        Me.ChkList_5_no_RadioButton.TabStop = True
        Me.ChkList_5_no_RadioButton.Text = "無"
        Me.ChkList_5_no_RadioButton.UseVisualStyleBackColor = True
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label30.Location = New System.Drawing.Point(7, 0)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(57, 16)
        Me.Label30.TabIndex = 68
        Me.Label30.Text = "5.VONIC"
        '
        'ChkList_5_std_Content_TextBox
        '
        Me.ChkList_5_std_Content_TextBox.Location = New System.Drawing.Point(113, 46)
        Me.ChkList_5_std_Content_TextBox.Name = "ChkList_5_std_Content_TextBox"
        Me.ChkList_5_std_Content_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_5_std_Content_TextBox.TabIndex = 21
        '
        'ChkList_5_nstd_Content_TextBox
        '
        Me.ChkList_5_nstd_Content_TextBox.Location = New System.Drawing.Point(113, 74)
        Me.ChkList_5_nstd_Content_TextBox.Name = "ChkList_5_nstd_Content_TextBox"
        Me.ChkList_5_nstd_Content_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_5_nstd_Content_TextBox.TabIndex = 22
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.CheckList3_FlowLayoutPanel)
        Me.TabPage4.Location = New System.Drawing.Point(4, 25)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(638, 388)
        Me.TabPage4.TabIndex = 2
        Me.TabPage4.Text = "Page3"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'CheckList3_FlowLayoutPanel
        '
        Me.CheckList3_FlowLayoutPanel.Controls.Add(Me.ChkList_7_Panel)
        Me.CheckList3_FlowLayoutPanel.Controls.Add(Me.ChkList_8_Panel)
        Me.CheckList3_FlowLayoutPanel.Controls.Add(Me.ChkList_9_Panel)
        Me.CheckList3_FlowLayoutPanel.Location = New System.Drawing.Point(6, 6)
        Me.CheckList3_FlowLayoutPanel.Name = "CheckList3_FlowLayoutPanel"
        Me.CheckList3_FlowLayoutPanel.Size = New System.Drawing.Size(626, 376)
        Me.CheckList3_FlowLayoutPanel.TabIndex = 62
        '
        'ChkList_7_Panel
        '
        Me.ChkList_7_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_7_Panel.Controls.Add(Me.ChkList_7_yes_RadioButton)
        Me.ChkList_7_Panel.Controls.Add(Me.ChkList_7_no_RadioButton)
        Me.ChkList_7_Panel.Controls.Add(Me.Label28)
        Me.ChkList_7_Panel.Controls.Add(Me.ChkList_7_yes1_content_TextBox)
        Me.ChkList_7_Panel.Location = New System.Drawing.Point(3, 3)
        Me.ChkList_7_Panel.Name = "ChkList_7_Panel"
        Me.ChkList_7_Panel.Size = New System.Drawing.Size(590, 68)
        Me.ChkList_7_Panel.TabIndex = 161
        '
        'ChkList_7_yes_RadioButton
        '
        Me.ChkList_7_yes_RadioButton.AutoSize = True
        Me.ChkList_7_yes_RadioButton.Location = New System.Drawing.Point(19, 21)
        Me.ChkList_7_yes_RadioButton.Name = "ChkList_7_yes_RadioButton"
        Me.ChkList_7_yes_RadioButton.Size = New System.Drawing.Size(87, 20)
        Me.ChkList_7_yes_RadioButton.TabIndex = 28
        Me.ChkList_7_yes_RadioButton.TabStop = True
        Me.ChkList_7_yes_RadioButton.Text = "有(文書No."
        Me.ChkList_7_yes_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_7_no_RadioButton
        '
        Me.ChkList_7_no_RadioButton.AutoSize = True
        Me.ChkList_7_no_RadioButton.Location = New System.Drawing.Point(19, 46)
        Me.ChkList_7_no_RadioButton.Name = "ChkList_7_no_RadioButton"
        Me.ChkList_7_no_RadioButton.Size = New System.Drawing.Size(38, 20)
        Me.ChkList_7_no_RadioButton.TabIndex = 30
        Me.ChkList_7_no_RadioButton.TabStop = True
        Me.ChkList_7_no_RadioButton.Text = "無"
        Me.ChkList_7_no_RadioButton.UseVisualStyleBackColor = True
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label28.Location = New System.Drawing.Point(7, 0)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(276, 16)
        Me.Label28.TabIndex = 86
        Me.Label28.Text = "7.軟體工直設計的一般資料以外有其他參考資料嗎?" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ChkList_7_yes1_content_TextBox
        '
        Me.ChkList_7_yes1_content_TextBox.Location = New System.Drawing.Point(113, 20)
        Me.ChkList_7_yes1_content_TextBox.Name = "ChkList_7_yes1_content_TextBox"
        Me.ChkList_7_yes1_content_TextBox.Size = New System.Drawing.Size(248, 23)
        Me.ChkList_7_yes1_content_TextBox.TabIndex = 29
        '
        'ChkList_8_Panel
        '
        Me.ChkList_8_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_8_Panel.Controls.Add(Me.Panel1)
        Me.ChkList_8_Panel.Controls.Add(Me.ChkList_8Item_RadioButton)
        Me.ChkList_8_Panel.Controls.Add(Me.Label29)
        Me.ChkList_8_Panel.Location = New System.Drawing.Point(3, 77)
        Me.ChkList_8_Panel.Name = "ChkList_8_Panel"
        Me.ChkList_8_Panel.Size = New System.Drawing.Size(590, 72)
        Me.ChkList_8_Panel.TabIndex = 162
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ChkList_8_yes_RadioButton)
        Me.Panel1.Controls.Add(Me.ChkList_8_no_RadioButton)
        Me.Panel1.Location = New System.Drawing.Point(10, 19)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(86, 50)
        Me.Panel1.TabIndex = 161
        '
        'ChkList_8_yes_RadioButton
        '
        Me.ChkList_8_yes_RadioButton.AutoSize = True
        Me.ChkList_8_yes_RadioButton.Location = New System.Drawing.Point(7, 0)
        Me.ChkList_8_yes_RadioButton.Name = "ChkList_8_yes_RadioButton"
        Me.ChkList_8_yes_RadioButton.Size = New System.Drawing.Size(43, 20)
        Me.ChkList_8_yes_RadioButton.TabIndex = 31
        Me.ChkList_8_yes_RadioButton.TabStop = True
        Me.ChkList_8_yes_RadioButton.Text = "OK"
        Me.ChkList_8_yes_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_8_no_RadioButton
        '
        Me.ChkList_8_no_RadioButton.AutoSize = True
        Me.ChkList_8_no_RadioButton.Location = New System.Drawing.Point(7, 27)
        Me.ChkList_8_no_RadioButton.Name = "ChkList_8_no_RadioButton"
        Me.ChkList_8_no_RadioButton.Size = New System.Drawing.Size(46, 20)
        Me.ChkList_8_no_RadioButton.TabIndex = 32
        Me.ChkList_8_no_RadioButton.TabStop = True
        Me.ChkList_8_no_RadioButton.Text = "NO"
        Me.ChkList_8_no_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_8Item_RadioButton
        '
        Me.ChkList_8Item_RadioButton.AutoSize = True
        Me.ChkList_8Item_RadioButton.Location = New System.Drawing.Point(327, 19)
        Me.ChkList_8Item_RadioButton.Name = "ChkList_8Item_RadioButton"
        Me.ChkList_8Item_RadioButton.Size = New System.Drawing.Size(47, 20)
        Me.ChkList_8Item_RadioButton.TabIndex = 3
        Me.ChkList_8Item_RadioButton.TabStop = True
        Me.ChkList_8Item_RadioButton.Text = "YES"
        Me.ChkList_8Item_RadioButton.UseVisualStyleBackColor = True
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label29.Location = New System.Drawing.Point(7, 0)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(443, 16)
        Me.Label29.TabIndex = 90
        Me.Label29.Text = "8.品目明細書・確認圖・OS和軟體仕樣書之間的對照結果？【有滿足特記事項嗎】"
        '
        'ChkList_9_Panel
        '
        Me.ChkList_9_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ChkList_9_Panel.Controls.Add(Me.ChkList_9_no_RadioButton)
        Me.ChkList_9_Panel.Controls.Add(Me.ChkList_9_yes_RadioButton)
        Me.ChkList_9_Panel.Controls.Add(Me.Label31)
        Me.ChkList_9_Panel.Location = New System.Drawing.Point(3, 155)
        Me.ChkList_9_Panel.Name = "ChkList_9_Panel"
        Me.ChkList_9_Panel.Size = New System.Drawing.Size(590, 82)
        Me.ChkList_9_Panel.TabIndex = 163
        '
        'ChkList_9_no_RadioButton
        '
        Me.ChkList_9_no_RadioButton.AutoSize = True
        Me.ChkList_9_no_RadioButton.Location = New System.Drawing.Point(19, 46)
        Me.ChkList_9_no_RadioButton.Name = "ChkList_9_no_RadioButton"
        Me.ChkList_9_no_RadioButton.Size = New System.Drawing.Size(46, 20)
        Me.ChkList_9_no_RadioButton.TabIndex = 35
        Me.ChkList_9_no_RadioButton.TabStop = True
        Me.ChkList_9_no_RadioButton.Text = "NO"
        Me.ChkList_9_no_RadioButton.UseVisualStyleBackColor = True
        '
        'ChkList_9_yes_RadioButton
        '
        Me.ChkList_9_yes_RadioButton.AutoSize = True
        Me.ChkList_9_yes_RadioButton.Location = New System.Drawing.Point(19, 19)
        Me.ChkList_9_yes_RadioButton.Name = "ChkList_9_yes_RadioButton"
        Me.ChkList_9_yes_RadioButton.Size = New System.Drawing.Size(43, 20)
        Me.ChkList_9_yes_RadioButton.TabIndex = 34
        Me.ChkList_9_yes_RadioButton.TabStop = True
        Me.ChkList_9_yes_RadioButton.Text = "OK"
        Me.ChkList_9_yes_RadioButton.UseVisualStyleBackColor = True
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label31.Location = New System.Drawing.Point(7, 0)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(162, 16)
        Me.Label31.TabIndex = 93
        Me.Label31.Text = "9.自已檢查表有作成並確認？"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label10.Location = New System.Drawing.Point(6, 19)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(80, 16)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "品目明細日期"
        '
        'ChkList_PaSheet_CheckBox
        '
        Me.ChkList_PaSheet_CheckBox.AutoSize = True
        Me.ChkList_PaSheet_CheckBox.Location = New System.Drawing.Point(270, 17)
        Me.ChkList_PaSheet_CheckBox.Name = "ChkList_PaSheet_CheckBox"
        Me.ChkList_PaSheet_CheckBox.Size = New System.Drawing.Size(99, 20)
        Me.ChkList_PaSheet_CheckBox.TabIndex = 1
        Me.ChkList_PaSheet_CheckBox.Text = "無品目明細表"
        Me.ChkList_PaSheet_CheckBox.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label11.Location = New System.Drawing.Point(6, 48)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(82, 16)
        Me.Label11.TabIndex = 12
        Me.Label11.Text = "ORDER SPEC"
        '
        'ChkList_OS_CheckBox
        '
        Me.ChkList_OS_CheckBox.AutoSize = True
        Me.ChkList_OS_CheckBox.Location = New System.Drawing.Point(270, 46)
        Me.ChkList_OS_CheckBox.Name = "ChkList_OS_CheckBox"
        Me.ChkList_OS_CheckBox.Size = New System.Drawing.Size(113, 20)
        Me.ChkList_OS_CheckBox.TabIndex = 2
        Me.ChkList_OS_CheckBox.Text = "無ORDER SPEC"
        Me.ChkList_OS_CheckBox.UseVisualStyleBackColor = True
        '
        'ChkList_Elec_DateTimePicker
        '
        Me.ChkList_Elec_DateTimePicker.Location = New System.Drawing.Point(92, 103)
        Me.ChkList_Elec_DateTimePicker.Name = "ChkList_Elec_DateTimePicker"
        Me.ChkList_Elec_DateTimePicker.Size = New System.Drawing.Size(153, 23)
        Me.ChkList_Elec_DateTimePicker.TabIndex = 33
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label12.Location = New System.Drawing.Point(6, 77)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(44, 16)
        Me.Label12.TabIndex = 17
        Me.Label12.Text = "確認圖"
        '
        'ChkList_Confirm_DateTimePicker
        '
        Me.ChkList_Confirm_DateTimePicker.Location = New System.Drawing.Point(92, 74)
        Me.ChkList_Confirm_DateTimePicker.Name = "ChkList_Confirm_DateTimePicker"
        Me.ChkList_Confirm_DateTimePicker.Size = New System.Drawing.Size(153, 23)
        Me.ChkList_Confirm_DateTimePicker.TabIndex = 32
        '
        'ChkList_Confirm_CheckBox
        '
        Me.ChkList_Confirm_CheckBox.AutoSize = True
        Me.ChkList_Confirm_CheckBox.Location = New System.Drawing.Point(270, 75)
        Me.ChkList_Confirm_CheckBox.Name = "ChkList_Confirm_CheckBox"
        Me.ChkList_Confirm_CheckBox.Size = New System.Drawing.Size(75, 20)
        Me.ChkList_Confirm_CheckBox.TabIndex = 3
        Me.ChkList_Confirm_CheckBox.Text = "無確認圖"
        Me.ChkList_Confirm_CheckBox.UseVisualStyleBackColor = True
        '
        'ChkList_OS_DateTimePicker
        '
        Me.ChkList_OS_DateTimePicker.Location = New System.Drawing.Point(92, 45)
        Me.ChkList_OS_DateTimePicker.Name = "ChkList_OS_DateTimePicker"
        Me.ChkList_OS_DateTimePicker.Size = New System.Drawing.Size(153, 23)
        Me.ChkList_OS_DateTimePicker.TabIndex = 31
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label13.Location = New System.Drawing.Point(6, 106)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 16)
        Me.Label13.TabIndex = 22
        Me.Label13.Text = "電氣圖面"
        '
        'ChkList_PaSheet_DateTimePicker
        '
        Me.ChkList_PaSheet_DateTimePicker.Location = New System.Drawing.Point(92, 16)
        Me.ChkList_PaSheet_DateTimePicker.Name = "ChkList_PaSheet_DateTimePicker"
        Me.ChkList_PaSheet_DateTimePicker.Size = New System.Drawing.Size(153, 23)
        Me.ChkList_PaSheet_DateTimePicker.TabIndex = 30
        '
        'ChkList_Elec_CheckBox
        '
        Me.ChkList_Elec_CheckBox.AutoSize = True
        Me.ChkList_Elec_CheckBox.Location = New System.Drawing.Point(270, 104)
        Me.ChkList_Elec_CheckBox.Name = "ChkList_Elec_CheckBox"
        Me.ChkList_Elec_CheckBox.Size = New System.Drawing.Size(87, 20)
        Me.ChkList_Elec_CheckBox.TabIndex = 4
        Me.ChkList_Elec_CheckBox.Text = "無電氣圖面"
        Me.ChkList_Elec_CheckBox.UseVisualStyleBackColor = True
        '
        'Use_ChkList_CheckBox
        '
        Me.Use_ChkList_CheckBox.AutoSize = True
        Me.Use_ChkList_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_ChkList_CheckBox.Name = "Use_ChkList_CheckBox"
        Me.Use_ChkList_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Use_ChkList_CheckBox.TabIndex = 1
        Me.Use_ChkList_CheckBox.UseVisualStyleBackColor = True
        '
        'Basic_TabPage
        '
        Me.Basic_TabPage.Controls.Add(Me.Basic_GroupBox)
        Me.Basic_TabPage.Controls.Add(Me.Use_Basic_CheckBox)
        Me.Basic_TabPage.Controls.Add(Me.ReminderMarquee_Label)
        Me.Basic_TabPage.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.Basic_TabPage.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_TabPage.Name = "Basic_TabPage"
        Me.Basic_TabPage.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.Basic_TabPage.TabIndex = 0
        Me.Basic_TabPage.Text = "基本"
        Me.Basic_TabPage.UseVisualStyleBackColor = True
        '
        'Basic_GroupBox
        '
        Me.Basic_GroupBox.Controls.Add(Me.Basic_Local_Label)
        Me.Basic_GroupBox.Controls.Add(Me.NumericUpDown1)
        Me.Basic_GroupBox.Controls.Add(Me.Label2)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_JobNoOld_Label)
        Me.Basic_GroupBox.Controls.Add(Me.Label222)
        Me.Basic_GroupBox.Controls.Add(Me.Label3)
        Me.Basic_GroupBox.Controls.Add(Me.Label221)
        Me.Basic_GroupBox.Controls.Add(Me.Label4)
        Me.Basic_GroupBox.Controls.Add(Me.Label220)
        Me.Basic_GroupBox.Controls.Add(Me.Label5)
        Me.Basic_GroupBox.Controls.Add(Me.Label219)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_DesingerChinese_ComboBox)
        Me.Basic_GroupBox.Controls.Add(Me.Label218)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_CheckerChinese_ComboBox)
        Me.Basic_GroupBox.Controls.Add(Me.Label217)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_JobNoNew_Label)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_ApproverEnglish_ComboBox)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_Local_ComboBox)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_CheckerEnglish_ComboBox)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_ApproverChinese_ComboBox)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_DesingerEnglish_ComboBox)
        Me.Basic_GroupBox.Controls.Add(Me.Label53)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_DrawDate_DateTimePicker)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_JobNoMOD_TextBox)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_JobNoNew_TextBox)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_JobNoOld_TextBox)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_JobName_TextBox)
        Me.Basic_GroupBox.Controls.Add(Me.Basic_JobNoMOD_Label)
        Me.Basic_GroupBox.Enabled = False
        Me.Basic_GroupBox.Location = New System.Drawing.Point(10, 30)
        Me.Basic_GroupBox.Name = "Basic_GroupBox"
        Me.Basic_GroupBox.Size = New System.Drawing.Size(632, 537)
        Me.Basic_GroupBox.TabIndex = 63
        Me.Basic_GroupBox.TabStop = False
        '
        'Basic_Local_Label
        '
        Me.Basic_Local_Label.AutoSize = True
        Me.Basic_Local_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_Local_Label.Location = New System.Drawing.Point(19, 30)
        Me.Basic_Local_Label.Name = "Basic_Local_Label"
        Me.Basic_Local_Label.Size = New System.Drawing.Size(41, 16)
        Me.Basic_Local_Label.TabIndex = 2
        Me.Basic_Local_Label.Text = "Local."
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(587, 505)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(39, 23)
        Me.NumericUpDown1.TabIndex = 50
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label2.Location = New System.Drawing.Point(19, 187)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 16)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "JobName."
        '
        'Basic_JobNoOld_Label
        '
        Me.Basic_JobNoOld_Label.AutoSize = True
        Me.Basic_JobNoOld_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_JobNoOld_Label.Location = New System.Drawing.Point(19, 107)
        Me.Basic_JobNoOld_Label.Name = "Basic_JobNoOld_Label"
        Me.Basic_JobNoOld_Label.Size = New System.Drawing.Size(70, 16)
        Me.Basic_JobNoOld_Label.TabIndex = 0
        Me.Basic_JobNoOld_Label.Text = "JobNo(舊)."
        '
        'Label222
        '
        Me.Label222.AutoSize = True
        Me.Label222.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label222.Location = New System.Drawing.Point(275, 307)
        Me.Label222.Name = "Label222"
        Me.Label222.Size = New System.Drawing.Size(35, 16)
        Me.Label222.TabIndex = 60
        Me.Label222.Text = "英文."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label3.Location = New System.Drawing.Point(19, 227)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Designer."
        '
        'Label221
        '
        Me.Label221.AutoSize = True
        Me.Label221.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label221.Location = New System.Drawing.Point(275, 267)
        Me.Label221.Name = "Label221"
        Me.Label221.Size = New System.Drawing.Size(35, 16)
        Me.Label221.TabIndex = 59
        Me.Label221.Text = "英文."
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label4.Location = New System.Drawing.Point(19, 267)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Checker."
        '
        'Label220
        '
        Me.Label220.AutoSize = True
        Me.Label220.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label220.Location = New System.Drawing.Point(116, 307)
        Me.Label220.Name = "Label220"
        Me.Label220.Size = New System.Drawing.Size(35, 16)
        Me.Label220.TabIndex = 58
        Me.Label220.Text = "中文."
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label5.Location = New System.Drawing.Point(19, 307)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 16)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Approver."
        '
        'Label219
        '
        Me.Label219.AutoSize = True
        Me.Label219.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label219.Location = New System.Drawing.Point(116, 267)
        Me.Label219.Name = "Label219"
        Me.Label219.Size = New System.Drawing.Size(35, 16)
        Me.Label219.TabIndex = 57
        Me.Label219.Text = "中文."
        '
        'Basic_DesingerChinese_ComboBox
        '
        Me.Basic_DesingerChinese_ComboBox.FormattingEnabled = True
        Me.Basic_DesingerChinese_ComboBox.Location = New System.Drawing.Point(154, 223)
        Me.Basic_DesingerChinese_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_DesingerChinese_ComboBox.Name = "Basic_DesingerChinese_ComboBox"
        Me.Basic_DesingerChinese_ComboBox.Size = New System.Drawing.Size(105, 24)
        Me.Basic_DesingerChinese_ComboBox.TabIndex = 4
        '
        'Label218
        '
        Me.Label218.AutoSize = True
        Me.Label218.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label218.Location = New System.Drawing.Point(275, 227)
        Me.Label218.Name = "Label218"
        Me.Label218.Size = New System.Drawing.Size(35, 16)
        Me.Label218.TabIndex = 56
        Me.Label218.Text = "英文."
        '
        'Basic_CheckerChinese_ComboBox
        '
        Me.Basic_CheckerChinese_ComboBox.FormattingEnabled = True
        Me.Basic_CheckerChinese_ComboBox.Location = New System.Drawing.Point(154, 263)
        Me.Basic_CheckerChinese_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_CheckerChinese_ComboBox.Name = "Basic_CheckerChinese_ComboBox"
        Me.Basic_CheckerChinese_ComboBox.Size = New System.Drawing.Size(105, 24)
        Me.Basic_CheckerChinese_ComboBox.TabIndex = 4
        '
        'Label217
        '
        Me.Label217.AutoSize = True
        Me.Label217.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label217.Location = New System.Drawing.Point(116, 227)
        Me.Label217.Name = "Label217"
        Me.Label217.Size = New System.Drawing.Size(35, 16)
        Me.Label217.TabIndex = 55
        Me.Label217.Text = "中文."
        '
        'Basic_JobNoNew_Label
        '
        Me.Basic_JobNoNew_Label.AutoSize = True
        Me.Basic_JobNoNew_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_JobNoNew_Label.Location = New System.Drawing.Point(19, 67)
        Me.Basic_JobNoNew_Label.Name = "Basic_JobNoNew_Label"
        Me.Basic_JobNoNew_Label.Size = New System.Drawing.Size(70, 16)
        Me.Basic_JobNoNew_Label.TabIndex = 0
        Me.Basic_JobNoNew_Label.Text = "JobNo(新)."
        '
        'Basic_ApproverEnglish_ComboBox
        '
        Me.Basic_ApproverEnglish_ComboBox.FormattingEnabled = True
        Me.Basic_ApproverEnglish_ComboBox.Location = New System.Drawing.Point(313, 303)
        Me.Basic_ApproverEnglish_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_ApproverEnglish_ComboBox.Name = "Basic_ApproverEnglish_ComboBox"
        Me.Basic_ApproverEnglish_ComboBox.Size = New System.Drawing.Size(105, 24)
        Me.Basic_ApproverEnglish_ComboBox.TabIndex = 54
        '
        'Basic_Local_ComboBox
        '
        Me.Basic_Local_ComboBox.FormattingEnabled = True
        Me.Basic_Local_ComboBox.Location = New System.Drawing.Point(119, 23)
        Me.Basic_Local_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_Local_ComboBox.Name = "Basic_Local_ComboBox"
        Me.Basic_Local_ComboBox.Size = New System.Drawing.Size(140, 24)
        Me.Basic_Local_ComboBox.TabIndex = 5
        '
        'Basic_CheckerEnglish_ComboBox
        '
        Me.Basic_CheckerEnglish_ComboBox.FormattingEnabled = True
        Me.Basic_CheckerEnglish_ComboBox.Location = New System.Drawing.Point(313, 263)
        Me.Basic_CheckerEnglish_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_CheckerEnglish_ComboBox.Name = "Basic_CheckerEnglish_ComboBox"
        Me.Basic_CheckerEnglish_ComboBox.Size = New System.Drawing.Size(105, 24)
        Me.Basic_CheckerEnglish_ComboBox.TabIndex = 53
        '
        'Basic_ApproverChinese_ComboBox
        '
        Me.Basic_ApproverChinese_ComboBox.FormattingEnabled = True
        Me.Basic_ApproverChinese_ComboBox.Location = New System.Drawing.Point(154, 303)
        Me.Basic_ApproverChinese_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_ApproverChinese_ComboBox.Name = "Basic_ApproverChinese_ComboBox"
        Me.Basic_ApproverChinese_ComboBox.Size = New System.Drawing.Size(105, 24)
        Me.Basic_ApproverChinese_ComboBox.TabIndex = 6
        '
        'Basic_DesingerEnglish_ComboBox
        '
        Me.Basic_DesingerEnglish_ComboBox.FormattingEnabled = True
        Me.Basic_DesingerEnglish_ComboBox.Location = New System.Drawing.Point(313, 223)
        Me.Basic_DesingerEnglish_ComboBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_DesingerEnglish_ComboBox.Name = "Basic_DesingerEnglish_ComboBox"
        Me.Basic_DesingerEnglish_ComboBox.Size = New System.Drawing.Size(105, 24)
        Me.Basic_DesingerEnglish_ComboBox.TabIndex = 52
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label53.Location = New System.Drawing.Point(18, 352)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(94, 16)
        Me.Label53.TabIndex = 31
        Me.Label53.Text = "Date(設計日期)."
        '
        'Basic_DrawDate_DateTimePicker
        '
        Me.Basic_DrawDate_DateTimePicker.Location = New System.Drawing.Point(119, 349)
        Me.Basic_DrawDate_DateTimePicker.Name = "Basic_DrawDate_DateTimePicker"
        Me.Basic_DrawDate_DateTimePicker.Size = New System.Drawing.Size(153, 23)
        Me.Basic_DrawDate_DateTimePicker.TabIndex = 32
        '
        'Basic_JobNoMOD_TextBox
        '
        Me.Basic_JobNoMOD_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_JobNoMOD_TextBox.Location = New System.Drawing.Point(119, 142)
        Me.Basic_JobNoMOD_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_JobNoMOD_TextBox.MaxLength = 50
        Me.Basic_JobNoMOD_TextBox.Name = "Basic_JobNoMOD_TextBox"
        Me.Basic_JobNoMOD_TextBox.Size = New System.Drawing.Size(140, 23)
        Me.Basic_JobNoMOD_TextBox.TabIndex = 48
        '
        'Basic_JobNoNew_TextBox
        '
        Me.Basic_JobNoNew_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_JobNoNew_TextBox.Location = New System.Drawing.Point(119, 64)
        Me.Basic_JobNoNew_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_JobNoNew_TextBox.MaxLength = 50
        Me.Basic_JobNoNew_TextBox.Name = "Basic_JobNoNew_TextBox"
        Me.Basic_JobNoNew_TextBox.Size = New System.Drawing.Size(140, 23)
        Me.Basic_JobNoNew_TextBox.TabIndex = 1
        '
        'Basic_JobNoOld_TextBox
        '
        Me.Basic_JobNoOld_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_JobNoOld_TextBox.Location = New System.Drawing.Point(119, 105)
        Me.Basic_JobNoOld_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_JobNoOld_TextBox.MaxLength = 50
        Me.Basic_JobNoOld_TextBox.Name = "Basic_JobNoOld_TextBox"
        Me.Basic_JobNoOld_TextBox.Size = New System.Drawing.Size(140, 23)
        Me.Basic_JobNoOld_TextBox.TabIndex = 1
        '
        'Basic_JobName_TextBox
        '
        Me.Basic_JobName_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_JobName_TextBox.Location = New System.Drawing.Point(119, 181)
        Me.Basic_JobName_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Basic_JobName_TextBox.MaxLength = 50
        Me.Basic_JobName_TextBox.Name = "Basic_JobName_TextBox"
        Me.Basic_JobName_TextBox.Size = New System.Drawing.Size(140, 23)
        Me.Basic_JobName_TextBox.TabIndex = 1
        '
        'Basic_JobNoMOD_Label
        '
        Me.Basic_JobNoMOD_Label.AutoSize = True
        Me.Basic_JobNoMOD_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Basic_JobNoMOD_Label.Location = New System.Drawing.Point(19, 147)
        Me.Basic_JobNoMOD_Label.Name = "Basic_JobNoMOD_Label"
        Me.Basic_JobNoMOD_Label.Size = New System.Drawing.Size(89, 16)
        Me.Basic_JobNoMOD_Label.TabIndex = 47
        Me.Basic_JobNoMOD_Label.Text = "JobNo(MOD)."
        '
        'Use_Basic_CheckBox
        '
        Me.Use_Basic_CheckBox.AutoSize = True
        Me.Use_Basic_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_Basic_CheckBox.Name = "Use_Basic_CheckBox"
        Me.Use_Basic_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Use_Basic_CheckBox.TabIndex = 51
        Me.Use_Basic_CheckBox.UseVisualStyleBackColor = True
        '
        'ReminderMarquee_Label
        '
        Me.ReminderMarquee_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.ReminderMarquee_Label.Location = New System.Drawing.Point(7, 11)
        Me.ReminderMarquee_Label.Name = "ReminderMarquee_Label"
        Me.ReminderMarquee_Label.Size = New System.Drawing.Size(635, 16)
        Me.ReminderMarquee_Label.TabIndex = 49
        '
        'Load_TabPage
        '
        Me.Load_TabPage.Controls.Add(Me.Load_Other_btn_GroupBox)
        Me.Load_TabPage.Controls.Add(Me.Load_SpecDWG_btn_GroupBox)
        Me.Load_TabPage.Controls.Add(Me.Load_TabControl)
        Me.Load_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.Load_TabPage.Name = "Load_TabPage"
        Me.Load_TabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.Load_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.Load_TabPage.TabIndex = 10
        Me.Load_TabPage.Text = "Load"
        Me.Load_TabPage.UseVisualStyleBackColor = True
        '
        'Load_Other_btn_GroupBox
        '
        Me.Load_Other_btn_GroupBox.Controls.Add(Me.CheckList_OutputButton)
        Me.Load_Other_btn_GroupBox.Controls.Add(Me.DWG_OutputButton)
        Me.Load_Other_btn_GroupBox.Controls.Add(Me.Spec_OutputButton)
        Me.Load_Other_btn_GroupBox.Location = New System.Drawing.Point(16, 452)
        Me.Load_Other_btn_GroupBox.Name = "Load_Other_btn_GroupBox"
        Me.Load_Other_btn_GroupBox.Size = New System.Drawing.Size(632, 81)
        Me.Load_Other_btn_GroupBox.TabIndex = 66
        Me.Load_Other_btn_GroupBox.TabStop = False
        Me.Load_Other_btn_GroupBox.Text = "其他單獨匯出"
        '
        'CheckList_OutputButton
        '
        Me.CheckList_OutputButton.Enabled = False
        Me.CheckList_OutputButton.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.CheckList_OutputButton.Location = New System.Drawing.Point(6, 22)
        Me.CheckList_OutputButton.Name = "CheckList_OutputButton"
        Me.CheckList_OutputButton.Size = New System.Drawing.Size(140, 45)
        Me.CheckList_OutputButton.TabIndex = 58
        Me.CheckList_OutputButton.Text = "Check list"
        Me.CheckList_OutputButton.UseVisualStyleBackColor = True
        '
        'DWG_OutputButton
        '
        Me.DWG_OutputButton.Enabled = False
        Me.DWG_OutputButton.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.DWG_OutputButton.Location = New System.Drawing.Point(153, 22)
        Me.DWG_OutputButton.Name = "DWG_OutputButton"
        Me.DWG_OutputButton.Size = New System.Drawing.Size(140, 45)
        Me.DWG_OutputButton.TabIndex = 61
        Me.DWG_OutputButton.Text = "送狀"
        Me.DWG_OutputButton.UseVisualStyleBackColor = True
        '
        'Spec_OutputButton
        '
        Me.Spec_OutputButton.Enabled = False
        Me.Spec_OutputButton.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Spec_OutputButton.Location = New System.Drawing.Point(299, 22)
        Me.Spec_OutputButton.Name = "Spec_OutputButton"
        Me.Spec_OutputButton.Size = New System.Drawing.Size(140, 45)
        Me.Spec_OutputButton.TabIndex = 62
        Me.Spec_OutputButton.Text = "仕樣書"
        Me.Spec_OutputButton.UseVisualStyleBackColor = True
        '
        'Load_SpecDWG_btn_GroupBox
        '
        Me.Load_SpecDWG_btn_GroupBox.Controls.Add(Me.All_OutputButton)
        Me.Load_SpecDWG_btn_GroupBox.Location = New System.Drawing.Point(16, 370)
        Me.Load_SpecDWG_btn_GroupBox.Name = "Load_SpecDWG_btn_GroupBox"
        Me.Load_SpecDWG_btn_GroupBox.Size = New System.Drawing.Size(632, 76)
        Me.Load_SpecDWG_btn_GroupBox.TabIndex = 65
        Me.Load_SpecDWG_btn_GroupBox.TabStop = False
        Me.Load_SpecDWG_btn_GroupBox.Text = "送狀 , 仕樣書 統一匯出"
        '
        'All_OutputButton
        '
        Me.All_OutputButton.Enabled = False
        Me.All_OutputButton.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.All_OutputButton.Location = New System.Drawing.Point(6, 22)
        Me.All_OutputButton.Name = "All_OutputButton"
        Me.All_OutputButton.Size = New System.Drawing.Size(140, 42)
        Me.All_OutputButton.TabIndex = 44
        Me.All_OutputButton.Text = "送狀+仕樣書"
        Me.All_OutputButton.UseVisualStyleBackColor = True
        '
        'Load_TabControl
        '
        Me.Load_TabControl.Controls.Add(Me.AutoLoad_TabPage)
        Me.Load_TabControl.Controls.Add(Me.Spec_TabPage)
        Me.Load_TabControl.Controls.Add(Me.CheckList_TabPage)
        Me.Load_TabControl.Controls.Add(Me.LoadSQL_TabPage)
        Me.Load_TabControl.Location = New System.Drawing.Point(6, 6)
        Me.Load_TabControl.Name = "Load_TabControl"
        Me.Load_TabControl.SelectedIndex = 0
        Me.Load_TabControl.Size = New System.Drawing.Size(652, 350)
        Me.Load_TabControl.TabIndex = 52
        '
        'AutoLoad_TabPage
        '
        Me.AutoLoad_TabPage.Controls.Add(Me.Load_AutoLoad_GroupBox)
        Me.AutoLoad_TabPage.Controls.Add(Me.JobMaker_LOAD_AutoLoad_CheckBox)
        Me.AutoLoad_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.AutoLoad_TabPage.Name = "AutoLoad_TabPage"
        Me.AutoLoad_TabPage.Size = New System.Drawing.Size(644, 321)
        Me.AutoLoad_TabPage.TabIndex = 4
        Me.AutoLoad_TabPage.Text = "自動讀取"
        Me.AutoLoad_TabPage.UseVisualStyleBackColor = True
        '
        'Load_AutoLoad_GroupBox
        '
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.JMFileConfirm_AutoLoad_Button)
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.Label54)
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.ComboBox1)
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.TextBox1)
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.Label57)
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.Label66)
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.JMFileCho_AutoLoad_Button)
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.PictureBox2)
        Me.Load_AutoLoad_GroupBox.Controls.Add(Me.JMFileCho_AutoLoad_TextBox)
        Me.Load_AutoLoad_GroupBox.Enabled = False
        Me.Load_AutoLoad_GroupBox.Location = New System.Drawing.Point(6, 35)
        Me.Load_AutoLoad_GroupBox.Name = "Load_AutoLoad_GroupBox"
        Me.Load_AutoLoad_GroupBox.Size = New System.Drawing.Size(630, 280)
        Me.Load_AutoLoad_GroupBox.TabIndex = 53
        Me.Load_AutoLoad_GroupBox.TabStop = False
        Me.Load_AutoLoad_GroupBox.Visible = False
        '
        'JMFileConfirm_AutoLoad_Button
        '
        Me.JMFileConfirm_AutoLoad_Button.Enabled = False
        Me.JMFileConfirm_AutoLoad_Button.Location = New System.Drawing.Point(597, 66)
        Me.JMFileConfirm_AutoLoad_Button.Name = "JMFileConfirm_AutoLoad_Button"
        Me.JMFileConfirm_AutoLoad_Button.Size = New System.Drawing.Size(23, 23)
        Me.JMFileConfirm_AutoLoad_Button.TabIndex = 68
        Me.JMFileConfirm_AutoLoad_Button.Text = "v"
        Me.JMFileConfirm_AutoLoad_Button.UseVisualStyleBackColor = True
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.ForeColor = System.Drawing.Color.Silver
        Me.Label54.Location = New System.Drawing.Point(125, 25)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(26, 16)
        Me.Label54.TabIndex = 67
        Me.Label54.Text = ">>"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(153, 21)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(250, 24)
        Me.ComboBox1.TabIndex = 66
        '
        'TextBox1
        '
        Me.TextBox1.AllowDrop = True
        Me.TextBox1.Location = New System.Drawing.Point(39, 22)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(80, 23)
        Me.TextBox1.TabIndex = 65
        '
        'Label57
        '
        Me.Label57.AutoSize = True
        Me.Label57.ForeColor = System.Drawing.Color.Silver
        Me.Label57.Location = New System.Drawing.Point(39, 120)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(35, 16)
        Me.Label57.TabIndex = 60
        Me.Label57.Text = "~~~"
        '
        'Label66
        '
        Me.Label66.AutoSize = True
        Me.Label66.ForeColor = System.Drawing.Color.Silver
        Me.Label66.Location = New System.Drawing.Point(39, 100)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(307, 16)
        Me.Label66.TabIndex = 59
        Me.Label66.Text = "當前預設路徑 ( 設定方式 : 設定>基本設定>預設路徑 ) ："
        '
        'JMFileCho_AutoLoad_Button
        '
        Me.JMFileCho_AutoLoad_Button.Location = New System.Drawing.Point(568, 66)
        Me.JMFileCho_AutoLoad_Button.Name = "JMFileCho_AutoLoad_Button"
        Me.JMFileCho_AutoLoad_Button.Size = New System.Drawing.Size(23, 23)
        Me.JMFileCho_AutoLoad_Button.TabIndex = 58
        Me.JMFileCho_AutoLoad_Button.Text = "..."
        Me.JMFileCho_AutoLoad_Button.UseVisualStyleBackColor = True
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.WindowsApp1.My.Resources.Resources.yaSan01
        Me.PictureBox2.Location = New System.Drawing.Point(9, 21)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 56
        Me.PictureBox2.TabStop = False
        '
        'JMFileCho_AutoLoad_TextBox
        '
        Me.JMFileCho_AutoLoad_TextBox.AllowDrop = True
        Me.JMFileCho_AutoLoad_TextBox.Location = New System.Drawing.Point(39, 66)
        Me.JMFileCho_AutoLoad_TextBox.Name = "JMFileCho_AutoLoad_TextBox"
        Me.JMFileCho_AutoLoad_TextBox.Size = New System.Drawing.Size(523, 23)
        Me.JMFileCho_AutoLoad_TextBox.TabIndex = 55
        '
        'JobMaker_LOAD_AutoLoad_CheckBox
        '
        Me.JobMaker_LOAD_AutoLoad_CheckBox.AutoSize = True
        Me.JobMaker_LOAD_AutoLoad_CheckBox.Location = New System.Drawing.Point(12, 15)
        Me.JobMaker_LOAD_AutoLoad_CheckBox.Name = "JobMaker_LOAD_AutoLoad_CheckBox"
        Me.JobMaker_LOAD_AutoLoad_CheckBox.Size = New System.Drawing.Size(104, 20)
        Me.JobMaker_LOAD_AutoLoad_CheckBox.TabIndex = 52
        Me.JobMaker_LOAD_AutoLoad_CheckBox.Text = "自動讀取Excel"
        Me.JobMaker_LOAD_AutoLoad_CheckBox.UseVisualStyleBackColor = True
        Me.JobMaker_LOAD_AutoLoad_CheckBox.Visible = False
        '
        'Spec_TabPage
        '
        Me.Spec_TabPage.Controls.Add(Me.Load_Spec_GroupBox)
        Me.Spec_TabPage.Controls.Add(Me.JobMaker_LOAD_Spec_CheckBox)
        Me.Spec_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.Spec_TabPage.Name = "Spec_TabPage"
        Me.Spec_TabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.Spec_TabPage.Size = New System.Drawing.Size(644, 321)
        Me.Spec_TabPage.TabIndex = 0
        Me.Spec_TabPage.Text = "仕樣書"
        Me.Spec_TabPage.UseVisualStyleBackColor = True
        '
        'Load_Spec_GroupBox
        '
        Me.Load_Spec_GroupBox.BackColor = System.Drawing.Color.Transparent
        Me.Load_Spec_GroupBox.Controls.Add(Me.Label1)
        Me.Load_Spec_GroupBox.Controls.Add(Me.JM_JobSelect_Spec_ComboBox)
        Me.Load_Spec_GroupBox.Controls.Add(Me.JM_JobSelect_Spec_TextBox)
        Me.Load_Spec_GroupBox.Controls.Add(Me.Button2)
        Me.Load_Spec_GroupBox.Controls.Add(Me.Button1)
        Me.Load_Spec_GroupBox.Controls.Add(Me.JM_DefaultPath_Spec_Label)
        Me.Load_Spec_GroupBox.Controls.Add(Me.Label149)
        Me.Load_Spec_GroupBox.Controls.Add(Me.JMFileCho_Spec_Button)
        Me.Load_Spec_GroupBox.Controls.Add(Me.PictureBox1)
        Me.Load_Spec_GroupBox.Controls.Add(Me.JMFileCho_Spec_TextBox)
        Me.Load_Spec_GroupBox.Enabled = False
        Me.Load_Spec_GroupBox.Location = New System.Drawing.Point(6, 35)
        Me.Load_Spec_GroupBox.Name = "Load_Spec_GroupBox"
        Me.Load_Spec_GroupBox.Size = New System.Drawing.Size(630, 280)
        Me.Load_Spec_GroupBox.TabIndex = 50
        Me.Load_Spec_GroupBox.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Silver
        Me.Label1.Location = New System.Drawing.Point(125, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(26, 16)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = ">>"
        '
        'JM_JobSelect_Spec_ComboBox
        '
        Me.JM_JobSelect_Spec_ComboBox.FormattingEnabled = True
        Me.JM_JobSelect_Spec_ComboBox.Location = New System.Drawing.Point(153, 21)
        Me.JM_JobSelect_Spec_ComboBox.Name = "JM_JobSelect_Spec_ComboBox"
        Me.JM_JobSelect_Spec_ComboBox.Size = New System.Drawing.Size(250, 24)
        Me.JM_JobSelect_Spec_ComboBox.TabIndex = 66
        '
        'JM_JobSelect_Spec_TextBox
        '
        Me.JM_JobSelect_Spec_TextBox.AllowDrop = True
        Me.JM_JobSelect_Spec_TextBox.Location = New System.Drawing.Point(39, 22)
        Me.JM_JobSelect_Spec_TextBox.Name = "JM_JobSelect_Spec_TextBox"
        Me.JM_JobSelect_Spec_TextBox.Size = New System.Drawing.Size(80, 23)
        Me.JM_JobSelect_Spec_TextBox.TabIndex = 65
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(35, 251)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(23, 23)
        Me.Button2.TabIndex = 64
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(5, 250)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(23, 23)
        Me.Button1.TabIndex = 63
        Me.Button1.Text = "..."
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'JM_DefaultPath_Spec_Label
        '
        Me.JM_DefaultPath_Spec_Label.AutoSize = True
        Me.JM_DefaultPath_Spec_Label.ForeColor = System.Drawing.Color.Silver
        Me.JM_DefaultPath_Spec_Label.Location = New System.Drawing.Point(39, 120)
        Me.JM_DefaultPath_Spec_Label.Name = "JM_DefaultPath_Spec_Label"
        Me.JM_DefaultPath_Spec_Label.Size = New System.Drawing.Size(35, 16)
        Me.JM_DefaultPath_Spec_Label.TabIndex = 59
        Me.JM_DefaultPath_Spec_Label.Text = "~~~"
        '
        'Label149
        '
        Me.Label149.AutoSize = True
        Me.Label149.ForeColor = System.Drawing.Color.Silver
        Me.Label149.Location = New System.Drawing.Point(39, 100)
        Me.Label149.Name = "Label149"
        Me.Label149.Size = New System.Drawing.Size(307, 16)
        Me.Label149.TabIndex = 58
        Me.Label149.Text = "當前預設路徑 ( 設定方式 : 設定>基本設定>預設路徑 ) ：" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'JMFileCho_Spec_Button
        '
        Me.JMFileCho_Spec_Button.Location = New System.Drawing.Point(591, 66)
        Me.JMFileCho_Spec_Button.Name = "JMFileCho_Spec_Button"
        Me.JMFileCho_Spec_Button.Size = New System.Drawing.Size(23, 23)
        Me.JMFileCho_Spec_Button.TabIndex = 57
        Me.JMFileCho_Spec_Button.Text = "..."
        Me.JMFileCho_Spec_Button.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.WindowsApp1.My.Resources.Resources.yaSan01
        Me.PictureBox1.Location = New System.Drawing.Point(9, 21)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 51
        Me.PictureBox1.TabStop = False
        '
        'JMFileCho_Spec_TextBox
        '
        Me.JMFileCho_Spec_TextBox.AllowDrop = True
        Me.JMFileCho_Spec_TextBox.Location = New System.Drawing.Point(39, 66)
        Me.JMFileCho_Spec_TextBox.Name = "JMFileCho_Spec_TextBox"
        Me.JMFileCho_Spec_TextBox.Size = New System.Drawing.Size(550, 23)
        Me.JMFileCho_Spec_TextBox.TabIndex = 50
        '
        'JobMaker_LOAD_Spec_CheckBox
        '
        Me.JobMaker_LOAD_Spec_CheckBox.AutoSize = True
        Me.JobMaker_LOAD_Spec_CheckBox.Location = New System.Drawing.Point(12, 15)
        Me.JobMaker_LOAD_Spec_CheckBox.Name = "JobMaker_LOAD_Spec_CheckBox"
        Me.JobMaker_LOAD_Spec_CheckBox.Size = New System.Drawing.Size(63, 20)
        Me.JobMaker_LOAD_Spec_CheckBox.TabIndex = 49
        Me.JobMaker_LOAD_Spec_CheckBox.Text = "仕樣書"
        Me.JobMaker_LOAD_Spec_CheckBox.UseVisualStyleBackColor = True
        '
        'CheckList_TabPage
        '
        Me.CheckList_TabPage.Controls.Add(Me.Load_ChkList_GroupBox)
        Me.CheckList_TabPage.Controls.Add(Me.JobMaker_LOAD_ChkList_CheckBox)
        Me.CheckList_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.CheckList_TabPage.Name = "CheckList_TabPage"
        Me.CheckList_TabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.CheckList_TabPage.Size = New System.Drawing.Size(644, 321)
        Me.CheckList_TabPage.TabIndex = 1
        Me.CheckList_TabPage.Text = "CheckList"
        Me.CheckList_TabPage.UseVisualStyleBackColor = True
        '
        'Load_ChkList_GroupBox
        '
        Me.Load_ChkList_GroupBox.Controls.Add(Me.Label6)
        Me.Load_ChkList_GroupBox.Controls.Add(Me.JM_JobSelect_CheckList_ComboBox)
        Me.Load_ChkList_GroupBox.Controls.Add(Me.JM_JobSelect_CheckList_TextBox)
        Me.Load_ChkList_GroupBox.Controls.Add(Me.JM_DefaultPath_CheckList_Label)
        Me.Load_ChkList_GroupBox.Controls.Add(Me.Label173)
        Me.Load_ChkList_GroupBox.Controls.Add(Me.JMFileCho_ChkList_Button)
        Me.Load_ChkList_GroupBox.Controls.Add(Me.PictureBox3)
        Me.Load_ChkList_GroupBox.Controls.Add(Me.JMFileCho_ChkList_TextBox)
        Me.Load_ChkList_GroupBox.Enabled = False
        Me.Load_ChkList_GroupBox.Location = New System.Drawing.Point(6, 35)
        Me.Load_ChkList_GroupBox.Name = "Load_ChkList_GroupBox"
        Me.Load_ChkList_GroupBox.Size = New System.Drawing.Size(630, 280)
        Me.Load_ChkList_GroupBox.TabIndex = 51
        Me.Load_ChkList_GroupBox.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Silver
        Me.Label6.Location = New System.Drawing.Point(125, 25)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(26, 16)
        Me.Label6.TabIndex = 67
        Me.Label6.Text = ">>"
        '
        'JM_JobSelect_CheckList_ComboBox
        '
        Me.JM_JobSelect_CheckList_ComboBox.FormattingEnabled = True
        Me.JM_JobSelect_CheckList_ComboBox.Location = New System.Drawing.Point(153, 21)
        Me.JM_JobSelect_CheckList_ComboBox.Name = "JM_JobSelect_CheckList_ComboBox"
        Me.JM_JobSelect_CheckList_ComboBox.Size = New System.Drawing.Size(250, 24)
        Me.JM_JobSelect_CheckList_ComboBox.TabIndex = 66
        '
        'JM_JobSelect_CheckList_TextBox
        '
        Me.JM_JobSelect_CheckList_TextBox.AllowDrop = True
        Me.JM_JobSelect_CheckList_TextBox.Location = New System.Drawing.Point(39, 22)
        Me.JM_JobSelect_CheckList_TextBox.Name = "JM_JobSelect_CheckList_TextBox"
        Me.JM_JobSelect_CheckList_TextBox.Size = New System.Drawing.Size(80, 23)
        Me.JM_JobSelect_CheckList_TextBox.TabIndex = 65
        '
        'JM_DefaultPath_CheckList_Label
        '
        Me.JM_DefaultPath_CheckList_Label.AutoSize = True
        Me.JM_DefaultPath_CheckList_Label.ForeColor = System.Drawing.Color.Silver
        Me.JM_DefaultPath_CheckList_Label.Location = New System.Drawing.Point(39, 120)
        Me.JM_DefaultPath_CheckList_Label.Name = "JM_DefaultPath_CheckList_Label"
        Me.JM_DefaultPath_CheckList_Label.Size = New System.Drawing.Size(35, 16)
        Me.JM_DefaultPath_CheckList_Label.TabIndex = 60
        Me.JM_DefaultPath_CheckList_Label.Text = "~~~"
        '
        'Label173
        '
        Me.Label173.AutoSize = True
        Me.Label173.ForeColor = System.Drawing.Color.Silver
        Me.Label173.Location = New System.Drawing.Point(39, 100)
        Me.Label173.Name = "Label173"
        Me.Label173.Size = New System.Drawing.Size(307, 16)
        Me.Label173.TabIndex = 59
        Me.Label173.Text = "當前預設路徑 ( 設定方式 : 設定>基本設定>預設路徑 ) ："
        '
        'JMFileCho_ChkList_Button
        '
        Me.JMFileCho_ChkList_Button.Location = New System.Drawing.Point(593, 66)
        Me.JMFileCho_ChkList_Button.Name = "JMFileCho_ChkList_Button"
        Me.JMFileCho_ChkList_Button.Size = New System.Drawing.Size(23, 23)
        Me.JMFileCho_ChkList_Button.TabIndex = 58
        Me.JMFileCho_ChkList_Button.Text = "..."
        Me.JMFileCho_ChkList_Button.UseVisualStyleBackColor = True
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(9, 21)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 56
        Me.PictureBox3.TabStop = False
        '
        'JMFileCho_ChkList_TextBox
        '
        Me.JMFileCho_ChkList_TextBox.AllowDrop = True
        Me.JMFileCho_ChkList_TextBox.Location = New System.Drawing.Point(39, 66)
        Me.JMFileCho_ChkList_TextBox.Name = "JMFileCho_ChkList_TextBox"
        Me.JMFileCho_ChkList_TextBox.Size = New System.Drawing.Size(550, 23)
        Me.JMFileCho_ChkList_TextBox.TabIndex = 55
        '
        'JobMaker_LOAD_ChkList_CheckBox
        '
        Me.JobMaker_LOAD_ChkList_CheckBox.AutoSize = True
        Me.JobMaker_LOAD_ChkList_CheckBox.Location = New System.Drawing.Point(12, 15)
        Me.JobMaker_LOAD_ChkList_CheckBox.Name = "JobMaker_LOAD_ChkList_CheckBox"
        Me.JobMaker_LOAD_ChkList_CheckBox.Size = New System.Drawing.Size(82, 20)
        Me.JobMaker_LOAD_ChkList_CheckBox.TabIndex = 49
        Me.JobMaker_LOAD_ChkList_CheckBox.Text = "Check List"
        Me.JobMaker_LOAD_ChkList_CheckBox.UseVisualStyleBackColor = True
        '
        'LoadSQL_TabPage
        '
        Me.LoadSQL_TabPage.Controls.Add(Me.Load_SQLite_GroupBox)
        Me.LoadSQL_TabPage.Controls.Add(Me.JobMaker_LOAD_SQLite_CheckBox)
        Me.LoadSQL_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.LoadSQL_TabPage.Name = "LoadSQL_TabPage"
        Me.LoadSQL_TabPage.Size = New System.Drawing.Size(644, 321)
        Me.LoadSQL_TabPage.TabIndex = 3
        Me.LoadSQL_TabPage.Text = "載入SQLite"
        Me.LoadSQL_TabPage.UseVisualStyleBackColor = True
        '
        'Load_SQLite_GroupBox
        '
        Me.Load_SQLite_GroupBox.BackColor = System.Drawing.Color.Transparent
        Me.Load_SQLite_GroupBox.Controls.Add(Me.Label109)
        Me.Load_SQLite_GroupBox.Controls.Add(Me.JM_JobSelect_SQLite_ComboBox)
        Me.Load_SQLite_GroupBox.Controls.Add(Me.JM_JobSelect_SQLite_TextBox)
        Me.Load_SQLite_GroupBox.Controls.Add(Me.JMFileConfirm_SQLite_Button)
        Me.Load_SQLite_GroupBox.Controls.Add(Me.JM_DefaultPath_SQLite_Label)
        Me.Load_SQLite_GroupBox.Controls.Add(Me.Label188)
        Me.Load_SQLite_GroupBox.Controls.Add(Me.JMFileCho_SQLite_Button)
        Me.Load_SQLite_GroupBox.Controls.Add(Me.PictureBox4)
        Me.Load_SQLite_GroupBox.Controls.Add(Me.JMFileCho_SQLite_TextBox)
        Me.Load_SQLite_GroupBox.Enabled = False
        Me.Load_SQLite_GroupBox.Location = New System.Drawing.Point(6, 35)
        Me.Load_SQLite_GroupBox.Name = "Load_SQLite_GroupBox"
        Me.Load_SQLite_GroupBox.Size = New System.Drawing.Size(630, 280)
        Me.Load_SQLite_GroupBox.TabIndex = 51
        Me.Load_SQLite_GroupBox.TabStop = False
        '
        'Label109
        '
        Me.Label109.AutoSize = True
        Me.Label109.ForeColor = System.Drawing.Color.Silver
        Me.Label109.Location = New System.Drawing.Point(125, 25)
        Me.Label109.Name = "Label109"
        Me.Label109.Size = New System.Drawing.Size(26, 16)
        Me.Label109.TabIndex = 64
        Me.Label109.Text = ">>"
        '
        'JM_JobSelect_SQLite_ComboBox
        '
        Me.JM_JobSelect_SQLite_ComboBox.FormattingEnabled = True
        Me.JM_JobSelect_SQLite_ComboBox.Location = New System.Drawing.Point(153, 21)
        Me.JM_JobSelect_SQLite_ComboBox.Name = "JM_JobSelect_SQLite_ComboBox"
        Me.JM_JobSelect_SQLite_ComboBox.Size = New System.Drawing.Size(150, 24)
        Me.JM_JobSelect_SQLite_ComboBox.TabIndex = 63
        '
        'JM_JobSelect_SQLite_TextBox
        '
        Me.JM_JobSelect_SQLite_TextBox.AllowDrop = True
        Me.JM_JobSelect_SQLite_TextBox.Location = New System.Drawing.Point(39, 22)
        Me.JM_JobSelect_SQLite_TextBox.Name = "JM_JobSelect_SQLite_TextBox"
        Me.JM_JobSelect_SQLite_TextBox.Size = New System.Drawing.Size(80, 23)
        Me.JM_JobSelect_SQLite_TextBox.TabIndex = 62
        '
        'JMFileConfirm_SQLite_Button
        '
        Me.JMFileConfirm_SQLite_Button.Enabled = False
        Me.JMFileConfirm_SQLite_Button.Location = New System.Drawing.Point(597, 66)
        Me.JMFileConfirm_SQLite_Button.Name = "JMFileConfirm_SQLite_Button"
        Me.JMFileConfirm_SQLite_Button.Size = New System.Drawing.Size(23, 23)
        Me.JMFileConfirm_SQLite_Button.TabIndex = 61
        Me.JMFileConfirm_SQLite_Button.Text = "v"
        Me.JMFileConfirm_SQLite_Button.UseVisualStyleBackColor = True
        '
        'JM_DefaultPath_SQLite_Label
        '
        Me.JM_DefaultPath_SQLite_Label.AutoSize = True
        Me.JM_DefaultPath_SQLite_Label.ForeColor = System.Drawing.Color.Silver
        Me.JM_DefaultPath_SQLite_Label.Location = New System.Drawing.Point(39, 120)
        Me.JM_DefaultPath_SQLite_Label.Name = "JM_DefaultPath_SQLite_Label"
        Me.JM_DefaultPath_SQLite_Label.Size = New System.Drawing.Size(35, 16)
        Me.JM_DefaultPath_SQLite_Label.TabIndex = 60
        Me.JM_DefaultPath_SQLite_Label.Text = "~~~"
        '
        'Label188
        '
        Me.Label188.AutoSize = True
        Me.Label188.ForeColor = System.Drawing.Color.Silver
        Me.Label188.Location = New System.Drawing.Point(39, 100)
        Me.Label188.Name = "Label188"
        Me.Label188.Size = New System.Drawing.Size(163, 16)
        Me.Label188.TabIndex = 59
        Me.Label188.Text = "當前預設路徑(無法個人設定):"
        '
        'JMFileCho_SQLite_Button
        '
        Me.JMFileCho_SQLite_Button.Location = New System.Drawing.Point(568, 66)
        Me.JMFileCho_SQLite_Button.Name = "JMFileCho_SQLite_Button"
        Me.JMFileCho_SQLite_Button.Size = New System.Drawing.Size(23, 23)
        Me.JMFileCho_SQLite_Button.TabIndex = 57
        Me.JMFileCho_SQLite_Button.Text = "..."
        Me.JMFileCho_SQLite_Button.UseVisualStyleBackColor = True
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = Global.WindowsApp1.My.Resources.Resources.yaSan01
        Me.PictureBox4.Location = New System.Drawing.Point(9, 21)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 51
        Me.PictureBox4.TabStop = False
        '
        'JMFileCho_SQLite_TextBox
        '
        Me.JMFileCho_SQLite_TextBox.AllowDrop = True
        Me.JMFileCho_SQLite_TextBox.Location = New System.Drawing.Point(39, 66)
        Me.JMFileCho_SQLite_TextBox.Name = "JMFileCho_SQLite_TextBox"
        Me.JMFileCho_SQLite_TextBox.Size = New System.Drawing.Size(523, 23)
        Me.JMFileCho_SQLite_TextBox.TabIndex = 50
        '
        'JobMaker_LOAD_SQLite_CheckBox
        '
        Me.JobMaker_LOAD_SQLite_CheckBox.AutoSize = True
        Me.JobMaker_LOAD_SQLite_CheckBox.Location = New System.Drawing.Point(12, 15)
        Me.JobMaker_LOAD_SQLite_CheckBox.Name = "JobMaker_LOAD_SQLite_CheckBox"
        Me.JobMaker_LOAD_SQLite_CheckBox.Size = New System.Drawing.Size(96, 20)
        Me.JobMaker_LOAD_SQLite_CheckBox.TabIndex = 49
        Me.JobMaker_LOAD_SQLite_CheckBox.Text = "載入工番Job"
        Me.JobMaker_LOAD_SQLite_CheckBox.UseVisualStyleBackColor = True
        '
        'JobMaker_TabControl
        '
        Me.JobMaker_TabControl.Controls.Add(Me.Load_TabPage)
        Me.JobMaker_TabControl.Controls.Add(Me.Basic_TabPage)
        Me.JobMaker_TabControl.Controls.Add(Me.CheckList)
        Me.JobMaker_TabControl.Controls.Add(Me.ProgramChange_TabPage)
        Me.JobMaker_TabControl.Controls.Add(Me.Spec)
        Me.JobMaker_TabControl.Controls.Add(Me.Important_TabPage)
        Me.JobMaker_TabControl.Controls.Add(Me.MMIC_TabPage)
        Me.JobMaker_TabControl.Controls.Add(Me.EepData_TabPage)
        Me.JobMaker_TabControl.Controls.Add(Me.G_TabPage)
        Me.JobMaker_TabControl.Controls.Add(Me.FinalCheck_TabPage)
        Me.JobMaker_TabControl.Controls.Add(Me.DWG_TabPage)
        Me.JobMaker_TabControl.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.JobMaker_TabControl.Location = New System.Drawing.Point(11, 16)
        Me.JobMaker_TabControl.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.JobMaker_TabControl.Name = "JobMaker_TabControl"
        Me.JobMaker_TabControl.SelectedIndex = 0
        Me.JobMaker_TabControl.Size = New System.Drawing.Size(672, 613)
        Me.JobMaker_TabControl.TabIndex = 5
        '
        'EepData_TabPage
        '
        Me.EepData_TabPage.Controls.Add(Me.Use_EepData_CheckBox)
        Me.EepData_TabPage.Controls.Add(Me.EepData_TabControl)
        Me.EepData_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.EepData_TabPage.Name = "EepData_TabPage"
        Me.EepData_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.EepData_TabPage.TabIndex = 12
        Me.EepData_TabPage.Text = "EepData"
        Me.EepData_TabPage.UseVisualStyleBackColor = True
        '
        'Use_EepData_CheckBox
        '
        Me.Use_EepData_CheckBox.AutoSize = True
        Me.Use_EepData_CheckBox.Location = New System.Drawing.Point(0, 0)
        Me.Use_EepData_CheckBox.Name = "Use_EepData_CheckBox"
        Me.Use_EepData_CheckBox.Size = New System.Drawing.Size(15, 14)
        Me.Use_EepData_CheckBox.TabIndex = 45
        Me.Use_EepData_CheckBox.UseVisualStyleBackColor = True
        Me.Use_EepData_CheckBox.Visible = False
        '
        'EepData_TabControl
        '
        Me.EepData_TabControl.Controls.Add(Me.EepData_TabPage1)
        Me.EepData_TabControl.Controls.Add(Me.EepData_TabPage2)
        Me.EepData_TabControl.Controls.Add(Me.EepData_TabPage3)
        Me.EepData_TabControl.Controls.Add(Me.EepData_TabPage4)
        Me.EepData_TabControl.Controls.Add(Me.EepData_TabPage5)
        Me.EepData_TabControl.Controls.Add(Me.EepData_TabPage6)
        Me.EepData_TabControl.Enabled = False
        Me.EepData_TabControl.Location = New System.Drawing.Point(3, 26)
        Me.EepData_TabControl.Name = "EepData_TabControl"
        Me.EepData_TabControl.SelectedIndex = 0
        Me.EepData_TabControl.Size = New System.Drawing.Size(649, 553)
        Me.EepData_TabControl.TabIndex = 0
        Me.EepData_TabControl.Visible = False
        '
        'EepData_TabPage1
        '
        Me.EepData_TabPage1.Controls.Add(Me.EepData_Page1_GroupBox)
        Me.EepData_TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.EepData_TabPage1.Name = "EepData_TabPage1"
        Me.EepData_TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.EepData_TabPage1.Size = New System.Drawing.Size(641, 524)
        Me.EepData_TabPage1.TabIndex = 0
        Me.EepData_TabPage1.Text = "Page 1"
        Me.EepData_TabPage1.UseVisualStyleBackColor = True
        '
        'EepData_Page1_GroupBox
        '
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_MachineRoom_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_MachineRoom_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_Speed_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_Speed_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_Capactity_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_Capactity_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_TopFL_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_TopFL_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_BtmFL_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_BtmFL_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_StopFL_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_StopFL_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_OpeType_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_OpeType_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_GspType_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_GspType_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_Purpose_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_Purpose_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_GroupNo_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_GroupNo_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_CarNo_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_CarNo_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_DrCloser_TextBox)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_DrCloser_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_DrType_Label)
        Me.EepData_Page1_GroupBox.Controls.Add(Me.EepData_DrType_TextBox)
        Me.EepData_Page1_GroupBox.Location = New System.Drawing.Point(6, 3)
        Me.EepData_Page1_GroupBox.Name = "EepData_Page1_GroupBox"
        Me.EepData_Page1_GroupBox.Size = New System.Drawing.Size(629, 515)
        Me.EepData_Page1_GroupBox.TabIndex = 142
        Me.EepData_Page1_GroupBox.TabStop = False
        '
        'EepData_MachineRoom_Label
        '
        Me.EepData_MachineRoom_Label.AutoSize = True
        Me.EepData_MachineRoom_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MachineRoom_Label.Location = New System.Drawing.Point(6, 19)
        Me.EepData_MachineRoom_Label.Name = "EepData_MachineRoom_Label"
        Me.EepData_MachineRoom_Label.Size = New System.Drawing.Size(108, 16)
        Me.EepData_MachineRoom_Label.TabIndex = 1
        Me.EepData_MachineRoom_Label.Text = "MACHINE ROOM"
        '
        'EepData_MachineRoom_TextBox
        '
        Me.EepData_MachineRoom_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MachineRoom_TextBox.Location = New System.Drawing.Point(224, 15)
        Me.EepData_MachineRoom_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_MachineRoom_TextBox.MaxLength = 999
        Me.EepData_MachineRoom_TextBox.Multiline = True
        Me.EepData_MachineRoom_TextBox.Name = "EepData_MachineRoom_TextBox"
        Me.EepData_MachineRoom_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_MachineRoom_TextBox.TabIndex = 14
        Me.EepData_MachineRoom_TextBox.Text = "機種(FP-17)"
        '
        'EepData_Speed_Label
        '
        Me.EepData_Speed_Label.AutoSize = True
        Me.EepData_Speed_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Speed_Label.Location = New System.Drawing.Point(6, 57)
        Me.EepData_Speed_Label.Name = "EepData_Speed_Label"
        Me.EepData_Speed_Label.Size = New System.Drawing.Size(90, 16)
        Me.EepData_Speed_Label.TabIndex = 2
        Me.EepData_Speed_Label.Text = "SPEED(m/min)"
        '
        'EepData_Speed_TextBox
        '
        Me.EepData_Speed_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Speed_TextBox.Location = New System.Drawing.Point(224, 53)
        Me.EepData_Speed_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Speed_TextBox.MaxLength = 999
        Me.EepData_Speed_TextBox.Multiline = True
        Me.EepData_Speed_TextBox.Name = "EepData_Speed_TextBox"
        Me.EepData_Speed_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Speed_TextBox.TabIndex = 15
        Me.EepData_Speed_TextBox.Text = "定格速度(m/min)"
        '
        'EepData_Capactity_Label
        '
        Me.EepData_Capactity_Label.AutoSize = True
        Me.EepData_Capactity_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Capactity_Label.Location = New System.Drawing.Point(6, 95)
        Me.EepData_Capactity_Label.Name = "EepData_Capactity_Label"
        Me.EepData_Capactity_Label.Size = New System.Drawing.Size(93, 16)
        Me.EepData_Capactity_Label.TabIndex = 3
        Me.EepData_Capactity_Label.Text = "CAPACTITY(kg)"
        '
        'EepData_Capactity_TextBox
        '
        Me.EepData_Capactity_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Capactity_TextBox.Location = New System.Drawing.Point(224, 91)
        Me.EepData_Capactity_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Capactity_TextBox.MaxLength = 999
        Me.EepData_Capactity_TextBox.Multiline = True
        Me.EepData_Capactity_TextBox.Name = "EepData_Capactity_TextBox"
        Me.EepData_Capactity_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Capactity_TextBox.TabIndex = 16
        Me.EepData_Capactity_TextBox.Text = "定格積載(kg)"
        '
        'EepData_TopFL_Label
        '
        Me.EepData_TopFL_Label.AutoSize = True
        Me.EepData_TopFL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_TopFL_Label.Location = New System.Drawing.Point(6, 133)
        Me.EepData_TopFL_Label.Name = "EepData_TopFL_Label"
        Me.EepData_TopFL_Label.Size = New System.Drawing.Size(139, 16)
        Me.EepData_TopFL_Label.TabIndex = 4
        Me.EepData_TopFL_Label.Text = "TOP TERMINAL FLOOR"
        '
        'EepData_TopFL_TextBox
        '
        Me.EepData_TopFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_TopFL_TextBox.Location = New System.Drawing.Point(224, 129)
        Me.EepData_TopFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_TopFL_TextBox.MaxLength = 999
        Me.EepData_TopFL_TextBox.Multiline = True
        Me.EepData_TopFL_TextBox.Name = "EepData_TopFL_TextBox"
        Me.EepData_TopFL_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_TopFL_TextBox.TabIndex = 17
        Me.EepData_TopFL_TextBox.Text = "最上階(實際樓層名)"
        '
        'EepData_BtmFL_Label
        '
        Me.EepData_BtmFL_Label.AutoSize = True
        Me.EepData_BtmFL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_BtmFL_Label.Location = New System.Drawing.Point(6, 171)
        Me.EepData_BtmFL_Label.Name = "EepData_BtmFL_Label"
        Me.EepData_BtmFL_Label.Size = New System.Drawing.Size(168, 16)
        Me.EepData_BtmFL_Label.TabIndex = 5
        Me.EepData_BtmFL_Label.Text = "BOTTOM TERMINAL FLOOR"
        '
        'EepData_BtmFL_TextBox
        '
        Me.EepData_BtmFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_BtmFL_TextBox.Location = New System.Drawing.Point(224, 167)
        Me.EepData_BtmFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_BtmFL_TextBox.MaxLength = 999
        Me.EepData_BtmFL_TextBox.Multiline = True
        Me.EepData_BtmFL_TextBox.Name = "EepData_BtmFL_TextBox"
        Me.EepData_BtmFL_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_BtmFL_TextBox.TabIndex = 18
        Me.EepData_BtmFL_TextBox.Text = "最下階(實際樓層名)"
        '
        'EepData_StopFL_Label
        '
        Me.EepData_StopFL_Label.AutoSize = True
        Me.EepData_StopFL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_StopFL_Label.Location = New System.Drawing.Point(6, 209)
        Me.EepData_StopFL_Label.Name = "EepData_StopFL_Label"
        Me.EepData_StopFL_Label.Size = New System.Drawing.Size(99, 16)
        Me.EepData_StopFL_Label.TabIndex = 6
        Me.EepData_StopFL_Label.Text = "NUM OF STOPS"
        '
        'EepData_StopFL_TextBox
        '
        Me.EepData_StopFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_StopFL_TextBox.Location = New System.Drawing.Point(224, 205)
        Me.EepData_StopFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_StopFL_TextBox.MaxLength = 999
        Me.EepData_StopFL_TextBox.Multiline = True
        Me.EepData_StopFL_TextBox.Name = "EepData_StopFL_TextBox"
        Me.EepData_StopFL_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_StopFL_TextBox.TabIndex = 19
        Me.EepData_StopFL_TextBox.Text = "停止數"
        '
        'EepData_OpeType_Label
        '
        Me.EepData_OpeType_Label.AutoSize = True
        Me.EepData_OpeType_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_OpeType_Label.Location = New System.Drawing.Point(6, 247)
        Me.EepData_OpeType_Label.Name = "EepData_OpeType_Label"
        Me.EepData_OpeType_Label.Size = New System.Drawing.Size(109, 16)
        Me.EepData_OpeType_Label.TabIndex = 7
        Me.EepData_OpeType_Label.Text = "OPERATION TYPE"
        '
        'EepData_OpeType_TextBox
        '
        Me.EepData_OpeType_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_OpeType_TextBox.Location = New System.Drawing.Point(224, 243)
        Me.EepData_OpeType_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_OpeType_TextBox.MaxLength = 999
        Me.EepData_OpeType_TextBox.Multiline = True
        Me.EepData_OpeType_TextBox.Name = "EepData_OpeType_TextBox"
        Me.EepData_OpeType_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_OpeType_TextBox.TabIndex = 20
        Me.EepData_OpeType_TextBox.Text = "操作方式"
        '
        'EepData_GspType_Label
        '
        Me.EepData_GspType_Label.AutoSize = True
        Me.EepData_GspType_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_GspType_Label.Location = New System.Drawing.Point(6, 285)
        Me.EepData_GspType_Label.Name = "EepData_GspType_Label"
        Me.EepData_GspType_Label.Size = New System.Drawing.Size(209, 16)
        Me.EepData_GspType_Label.TabIndex = 8
        Me.EepData_GspType_Label.Text = "GROUP MIC TYPE(GROUP SYSTEM)"
        '
        'EepData_GspType_TextBox
        '
        Me.EepData_GspType_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_GspType_TextBox.Location = New System.Drawing.Point(224, 281)
        Me.EepData_GspType_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_GspType_TextBox.MaxLength = 999
        Me.EepData_GspType_TextBox.Multiline = True
        Me.EepData_GspType_TextBox.Name = "EepData_GspType_TextBox"
        Me.EepData_GspType_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_GspType_TextBox.TabIndex = 21
        '
        'EepData_Purpose_Label
        '
        Me.EepData_Purpose_Label.AutoSize = True
        Me.EepData_Purpose_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Purpose_Label.Location = New System.Drawing.Point(6, 323)
        Me.EepData_Purpose_Label.Name = "EepData_Purpose_Label"
        Me.EepData_Purpose_Label.Size = New System.Drawing.Size(63, 16)
        Me.EepData_Purpose_Label.TabIndex = 9
        Me.EepData_Purpose_Label.Text = "PURPOSE"
        '
        'EepData_Purpose_TextBox
        '
        Me.EepData_Purpose_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Purpose_TextBox.Location = New System.Drawing.Point(224, 319)
        Me.EepData_Purpose_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Purpose_TextBox.MaxLength = 999
        Me.EepData_Purpose_TextBox.Multiline = True
        Me.EepData_Purpose_TextBox.Name = "EepData_Purpose_TextBox"
        Me.EepData_Purpose_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Purpose_TextBox.TabIndex = 22
        Me.EepData_Purpose_TextBox.Text = "用途"
        '
        'EepData_GroupNo_Label
        '
        Me.EepData_GroupNo_Label.AutoSize = True
        Me.EepData_GroupNo_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_GroupNo_Label.Location = New System.Drawing.Point(6, 361)
        Me.EepData_GroupNo_Label.Name = "EepData_GroupNo_Label"
        Me.EepData_GroupNo_Label.Size = New System.Drawing.Size(139, 16)
        Me.EepData_GroupNo_Label.TabIndex = 10
        Me.EepData_GroupNo_Label.Text = "NO. OF CAR IN GROUP"
        '
        'EepData_GroupNo_TextBox
        '
        Me.EepData_GroupNo_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_GroupNo_TextBox.Location = New System.Drawing.Point(224, 357)
        Me.EepData_GroupNo_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_GroupNo_TextBox.MaxLength = 999
        Me.EepData_GroupNo_TextBox.Multiline = True
        Me.EepData_GroupNo_TextBox.Name = "EepData_GroupNo_TextBox"
        Me.EepData_GroupNo_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_GroupNo_TextBox.TabIndex = 23
        '
        'EepData_CarNo_Label
        '
        Me.EepData_CarNo_Label.AutoSize = True
        Me.EepData_CarNo_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_CarNo_Label.Location = New System.Drawing.Point(6, 399)
        Me.EepData_CarNo_Label.Name = "EepData_CarNo_Label"
        Me.EepData_CarNo_Label.Size = New System.Drawing.Size(58, 16)
        Me.EepData_CarNo_Label.TabIndex = 11
        Me.EepData_CarNo_Label.Text = "CAR NO."
        '
        'EepData_CarNo_TextBox
        '
        Me.EepData_CarNo_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_CarNo_TextBox.Location = New System.Drawing.Point(224, 395)
        Me.EepData_CarNo_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_CarNo_TextBox.MaxLength = 999
        Me.EepData_CarNo_TextBox.Multiline = True
        Me.EepData_CarNo_TextBox.Name = "EepData_CarNo_TextBox"
        Me.EepData_CarNo_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_CarNo_TextBox.TabIndex = 24
        '
        'EepData_DrCloser_TextBox
        '
        Me.EepData_DrCloser_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrCloser_TextBox.Location = New System.Drawing.Point(224, 433)
        Me.EepData_DrCloser_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_DrCloser_TextBox.MaxLength = 999
        Me.EepData_DrCloser_TextBox.Multiline = True
        Me.EepData_DrCloser_TextBox.Name = "EepData_DrCloser_TextBox"
        Me.EepData_DrCloser_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_DrCloser_TextBox.TabIndex = 25
        '
        'EepData_DrCloser_Label
        '
        Me.EepData_DrCloser_Label.AutoSize = True
        Me.EepData_DrCloser_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrCloser_Label.Location = New System.Drawing.Point(6, 437)
        Me.EepData_DrCloser_Label.Name = "EepData_DrCloser_Label"
        Me.EepData_DrCloser_Label.Size = New System.Drawing.Size(94, 16)
        Me.EepData_DrCloser_Label.TabIndex = 12
        Me.EepData_DrCloser_Label.Text = "DOOR CLOSER"
        '
        'EepData_DrType_Label
        '
        Me.EepData_DrType_Label.AutoSize = True
        Me.EepData_DrType_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrType_Label.Location = New System.Drawing.Point(6, 475)
        Me.EepData_DrType_Label.Name = "EepData_DrType_Label"
        Me.EepData_DrType_Label.Size = New System.Drawing.Size(125, 16)
        Me.EepData_DrType_Label.TabIndex = 13
        Me.EepData_DrType_Label.Text = "DOOR TYPE(FRONT)"
        '
        'EepData_DrType_TextBox
        '
        Me.EepData_DrType_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrType_TextBox.Location = New System.Drawing.Point(224, 471)
        Me.EepData_DrType_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_DrType_TextBox.MaxLength = 999
        Me.EepData_DrType_TextBox.Multiline = True
        Me.EepData_DrType_TextBox.Name = "EepData_DrType_TextBox"
        Me.EepData_DrType_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_DrType_TextBox.TabIndex = 26
        '
        'EepData_TabPage2
        '
        Me.EepData_TabPage2.Controls.Add(Me.EepData_Page2_GroupBox)
        Me.EepData_TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.EepData_TabPage2.Name = "EepData_TabPage2"
        Me.EepData_TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.EepData_TabPage2.Size = New System.Drawing.Size(641, 524)
        Me.EepData_TabPage2.TabIndex = 1
        Me.EepData_TabPage2.Text = "Page 2"
        Me.EepData_TabPage2.UseVisualStyleBackColor = True
        '
        'EepData_Page2_GroupBox
        '
        Me.EepData_Page2_GroupBox.Controls.Add(Me.Label80)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.Label78)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.Label77)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_DrFrontWidth_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_DrFrontWidth_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_Landic_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_Landic_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_EnergyRe_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_EnergyRe_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_Indep_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_Indep_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_SpecMainFL_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_SpecMainFL_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_MainFL_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_MainFL_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_MainFL_FR_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_MainFL_FR_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_Seismic_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_Seismic_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_Nudging_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_Nudging_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_AutoByPass_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_AutoByPass_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_AutoFan_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_AutoFan_TextBox)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_DrHold_Label)
        Me.EepData_Page2_GroupBox.Controls.Add(Me.EepData_DrHold_TextBox)
        Me.EepData_Page2_GroupBox.Location = New System.Drawing.Point(6, 3)
        Me.EepData_Page2_GroupBox.Name = "EepData_Page2_GroupBox"
        Me.EepData_Page2_GroupBox.Size = New System.Drawing.Size(629, 515)
        Me.EepData_Page2_GroupBox.TabIndex = 168
        Me.EepData_Page2_GroupBox.TabStop = False
        '
        'Label80
        '
        Me.Label80.AutoSize = True
        Me.Label80.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label80.Location = New System.Drawing.Point(174, 246)
        Me.Label80.Name = "Label80"
        Me.Label80.Size = New System.Drawing.Size(45, 16)
        Me.Label80.TabIndex = 168
        Me.Label80.Text = "(正/背)"
        '
        'Label78
        '
        Me.Label78.AutoSize = True
        Me.Label78.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label78.Location = New System.Drawing.Point(174, 208)
        Me.Label78.Name = "Label78"
        Me.Label78.Size = New System.Drawing.Size(40, 16)
        Me.Label78.TabIndex = 168
        Me.Label78.Text = "(樓層)"
        '
        'Label77
        '
        Me.Label77.AutoSize = True
        Me.Label77.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label77.Location = New System.Drawing.Point(174, 170)
        Me.Label77.Name = "Label77"
        Me.Label77.Size = New System.Drawing.Size(45, 16)
        Me.Label77.TabIndex = 167
        Me.Label77.Text = "(有/無)"
        '
        'EepData_DrFrontWidth_Label
        '
        Me.EepData_DrFrontWidth_Label.AutoSize = True
        Me.EepData_DrFrontWidth_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrFrontWidth_Label.Location = New System.Drawing.Point(12, 18)
        Me.EepData_DrFrontWidth_Label.Name = "EepData_DrFrontWidth_Label"
        Me.EepData_DrFrontWidth_Label.Size = New System.Drawing.Size(189, 16)
        Me.EepData_DrFrontWidth_Label.TabIndex = 142
        Me.EepData_DrFrontWidth_Label.Text = "OPENING WIDTH(mm) (FRONT)"
        '
        'EepData_DrFrontWidth_TextBox
        '
        Me.EepData_DrFrontWidth_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrFrontWidth_TextBox.Location = New System.Drawing.Point(224, 15)
        Me.EepData_DrFrontWidth_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_DrFrontWidth_TextBox.MaxLength = 999
        Me.EepData_DrFrontWidth_TextBox.Multiline = True
        Me.EepData_DrFrontWidth_TextBox.Name = "EepData_DrFrontWidth_TextBox"
        Me.EepData_DrFrontWidth_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_DrFrontWidth_TextBox.TabIndex = 117
        '
        'EepData_Landic_TextBox
        '
        Me.EepData_Landic_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Landic_TextBox.Location = New System.Drawing.Point(224, 53)
        Me.EepData_Landic_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Landic_TextBox.MaxLength = 999
        Me.EepData_Landic_TextBox.Multiline = True
        Me.EepData_Landic_TextBox.Name = "EepData_Landic_TextBox"
        Me.EepData_Landic_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Landic_TextBox.TabIndex = 119
        '
        'EepData_Landic_Label
        '
        Me.EepData_Landic_Label.AutoSize = True
        Me.EepData_Landic_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Landic_Label.Location = New System.Drawing.Point(12, 56)
        Me.EepData_Landic_Label.Name = "EepData_Landic_Label"
        Me.EepData_Landic_Label.Size = New System.Drawing.Size(52, 16)
        Me.EepData_Landic_Label.TabIndex = 144
        Me.EepData_Landic_Label.Text = "LANDIC"
        '
        'EepData_EnergyRe_Label
        '
        Me.EepData_EnergyRe_Label.AutoSize = True
        Me.EepData_EnergyRe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EnergyRe_Label.Location = New System.Drawing.Point(12, 94)
        Me.EepData_EnergyRe_Label.Name = "EepData_EnergyRe_Label"
        Me.EepData_EnergyRe_Label.Size = New System.Drawing.Size(203, 16)
        Me.EepData_EnergyRe_Label.TabIndex = 146
        Me.EepData_EnergyRe_Label.Text = "ENERGY REGENERATION SYSTEM"
        '
        'EepData_EnergyRe_TextBox
        '
        Me.EepData_EnergyRe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EnergyRe_TextBox.Location = New System.Drawing.Point(224, 91)
        Me.EepData_EnergyRe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_EnergyRe_TextBox.MaxLength = 999
        Me.EepData_EnergyRe_TextBox.Multiline = True
        Me.EepData_EnergyRe_TextBox.Name = "EepData_EnergyRe_TextBox"
        Me.EepData_EnergyRe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_EnergyRe_TextBox.TabIndex = 121
        '
        'EepData_Indep_Label
        '
        Me.EepData_Indep_Label.AutoSize = True
        Me.EepData_Indep_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Indep_Label.Location = New System.Drawing.Point(12, 132)
        Me.EepData_Indep_Label.Name = "EepData_Indep_Label"
        Me.EepData_Indep_Label.Size = New System.Drawing.Size(167, 16)
        Me.EepData_Indep_Label.TabIndex = 148
        Me.EepData_Indep_Label.Text = "INDEPENDENT OPERATION"
        '
        'EepData_Indep_TextBox
        '
        Me.EepData_Indep_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Indep_TextBox.Location = New System.Drawing.Point(224, 129)
        Me.EepData_Indep_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Indep_TextBox.MaxLength = 999
        Me.EepData_Indep_TextBox.Multiline = True
        Me.EepData_Indep_TextBox.Name = "EepData_Indep_TextBox"
        Me.EepData_Indep_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Indep_TextBox.TabIndex = 123
        '
        'EepData_SpecMainFL_Label
        '
        Me.EepData_SpecMainFL_Label.AutoSize = True
        Me.EepData_SpecMainFL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_SpecMainFL_Label.Location = New System.Drawing.Point(12, 170)
        Me.EepData_SpecMainFL_Label.Name = "EepData_SpecMainFL_Label"
        Me.EepData_SpecMainFL_Label.Size = New System.Drawing.Size(156, 16)
        Me.EepData_SpecMainFL_Label.TabIndex = 150
        Me.EepData_SpecMainFL_Label.Text = "RETURN TO MAIN FLOOR"
        '
        'EepData_SpecMainFL_TextBox
        '
        Me.EepData_SpecMainFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_SpecMainFL_TextBox.Location = New System.Drawing.Point(224, 167)
        Me.EepData_SpecMainFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_SpecMainFL_TextBox.MaxLength = 999
        Me.EepData_SpecMainFL_TextBox.Multiline = True
        Me.EepData_SpecMainFL_TextBox.Name = "EepData_SpecMainFL_TextBox"
        Me.EepData_SpecMainFL_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_SpecMainFL_TextBox.TabIndex = 125
        '
        'EepData_MainFL_Label
        '
        Me.EepData_MainFL_Label.AutoSize = True
        Me.EepData_MainFL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MainFL_Label.Location = New System.Drawing.Point(12, 208)
        Me.EepData_MainFL_Label.Name = "EepData_MainFL_Label"
        Me.EepData_MainFL_Label.Size = New System.Drawing.Size(141, 16)
        Me.EepData_MainFL_Label.TabIndex = 152
        Me.EepData_MainFL_Label.Text = "MAIN(RETURN) FLOOR"
        '
        'EepData_MainFL_TextBox
        '
        Me.EepData_MainFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MainFL_TextBox.Location = New System.Drawing.Point(224, 205)
        Me.EepData_MainFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_MainFL_TextBox.MaxLength = 999
        Me.EepData_MainFL_TextBox.Multiline = True
        Me.EepData_MainFL_TextBox.Name = "EepData_MainFL_TextBox"
        Me.EepData_MainFL_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_MainFL_TextBox.TabIndex = 127
        '
        'EepData_MainFL_FR_Label
        '
        Me.EepData_MainFL_FR_Label.AutoSize = True
        Me.EepData_MainFL_FR_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MainFL_FR_Label.Location = New System.Drawing.Point(12, 246)
        Me.EepData_MainFL_FR_Label.Name = "EepData_MainFL_FR_Label"
        Me.EepData_MainFL_FR_Label.Size = New System.Drawing.Size(84, 16)
        Me.EepData_MainFL_FR_Label.TabIndex = 154
        Me.EepData_MainFL_FR_Label.Text = "MAIN FLOOR"
        '
        'EepData_MainFL_FR_TextBox
        '
        Me.EepData_MainFL_FR_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MainFL_FR_TextBox.Location = New System.Drawing.Point(224, 243)
        Me.EepData_MainFL_FR_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_MainFL_FR_TextBox.MaxLength = 999
        Me.EepData_MainFL_FR_TextBox.Multiline = True
        Me.EepData_MainFL_FR_TextBox.Name = "EepData_MainFL_FR_TextBox"
        Me.EepData_MainFL_FR_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_MainFL_FR_TextBox.TabIndex = 129
        '
        'EepData_Seismic_Label
        '
        Me.EepData_Seismic_Label.AutoSize = True
        Me.EepData_Seismic_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Seismic_Label.Location = New System.Drawing.Point(12, 284)
        Me.EepData_Seismic_Label.Name = "EepData_Seismic_Label"
        Me.EepData_Seismic_Label.Size = New System.Drawing.Size(128, 16)
        Me.EepData_Seismic_Label.TabIndex = 156
        Me.EepData_Seismic_Label.Text = "SEISMIC OPERATION"
        '
        'EepData_Seismic_TextBox
        '
        Me.EepData_Seismic_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Seismic_TextBox.Location = New System.Drawing.Point(224, 281)
        Me.EepData_Seismic_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Seismic_TextBox.MaxLength = 999
        Me.EepData_Seismic_TextBox.Multiline = True
        Me.EepData_Seismic_TextBox.Name = "EepData_Seismic_TextBox"
        Me.EepData_Seismic_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Seismic_TextBox.TabIndex = 131
        '
        'EepData_Nudging_Label
        '
        Me.EepData_Nudging_Label.AutoSize = True
        Me.EepData_Nudging_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Nudging_Label.Location = New System.Drawing.Point(12, 325)
        Me.EepData_Nudging_Label.Name = "EepData_Nudging_Label"
        Me.EepData_Nudging_Label.Size = New System.Drawing.Size(67, 16)
        Me.EepData_Nudging_Label.TabIndex = 160
        Me.EepData_Nudging_Label.Text = "NUDGING"
        '
        'EepData_Nudging_TextBox
        '
        Me.EepData_Nudging_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Nudging_TextBox.Location = New System.Drawing.Point(224, 322)
        Me.EepData_Nudging_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Nudging_TextBox.MaxLength = 999
        Me.EepData_Nudging_TextBox.Multiline = True
        Me.EepData_Nudging_TextBox.Name = "EepData_Nudging_TextBox"
        Me.EepData_Nudging_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Nudging_TextBox.TabIndex = 135
        '
        'EepData_AutoByPass_Label
        '
        Me.EepData_AutoByPass_Label.AutoSize = True
        Me.EepData_AutoByPass_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_AutoByPass_Label.Location = New System.Drawing.Point(12, 363)
        Me.EepData_AutoByPass_Label.Name = "EepData_AutoByPass_Label"
        Me.EepData_AutoByPass_Label.Size = New System.Drawing.Size(88, 16)
        Me.EepData_AutoByPass_Label.TabIndex = 162
        Me.EepData_AutoByPass_Label.Text = "AUTO BYPASS"
        '
        'EepData_AutoByPass_TextBox
        '
        Me.EepData_AutoByPass_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_AutoByPass_TextBox.Location = New System.Drawing.Point(224, 360)
        Me.EepData_AutoByPass_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_AutoByPass_TextBox.MaxLength = 999
        Me.EepData_AutoByPass_TextBox.Multiline = True
        Me.EepData_AutoByPass_TextBox.Name = "EepData_AutoByPass_TextBox"
        Me.EepData_AutoByPass_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_AutoByPass_TextBox.TabIndex = 137
        '
        'EepData_AutoFan_Label
        '
        Me.EepData_AutoFan_Label.AutoSize = True
        Me.EepData_AutoFan_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_AutoFan_Label.Location = New System.Drawing.Point(12, 401)
        Me.EepData_AutoFan_Label.Name = "EepData_AutoFan_Label"
        Me.EepData_AutoFan_Label.Size = New System.Drawing.Size(132, 16)
        Me.EepData_AutoFan_Label.TabIndex = 164
        Me.EepData_AutoFan_Label.Text = "AUTOMATIC FAN OFF"
        '
        'EepData_AutoFan_TextBox
        '
        Me.EepData_AutoFan_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_AutoFan_TextBox.Location = New System.Drawing.Point(224, 398)
        Me.EepData_AutoFan_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_AutoFan_TextBox.MaxLength = 999
        Me.EepData_AutoFan_TextBox.Multiline = True
        Me.EepData_AutoFan_TextBox.Name = "EepData_AutoFan_TextBox"
        Me.EepData_AutoFan_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_AutoFan_TextBox.TabIndex = 139
        '
        'EepData_DrHold_Label
        '
        Me.EepData_DrHold_Label.AutoSize = True
        Me.EepData_DrHold_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrHold_Label.Location = New System.Drawing.Point(12, 439)
        Me.EepData_DrHold_Label.Name = "EepData_DrHold_Label"
        Me.EepData_DrHold_Label.Size = New System.Drawing.Size(186, 16)
        Me.EepData_DrHold_Label.TabIndex = 166
        Me.EepData_DrHold_Label.Text = "DOOR HOLD SWITCH/BUTTON"
        '
        'EepData_DrHold_TextBox
        '
        Me.EepData_DrHold_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrHold_TextBox.Location = New System.Drawing.Point(224, 436)
        Me.EepData_DrHold_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_DrHold_TextBox.MaxLength = 999
        Me.EepData_DrHold_TextBox.Multiline = True
        Me.EepData_DrHold_TextBox.Name = "EepData_DrHold_TextBox"
        Me.EepData_DrHold_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_DrHold_TextBox.TabIndex = 141
        '
        'EepData_TabPage3
        '
        Me.EepData_TabPage3.Controls.Add(Me.EepData_Page3_GroupBox)
        Me.EepData_TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.EepData_TabPage3.Name = "EepData_TabPage3"
        Me.EepData_TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.EepData_TabPage3.Size = New System.Drawing.Size(641, 524)
        Me.EepData_TabPage3.TabIndex = 2
        Me.EepData_TabPage3.Text = "Page 3"
        Me.EepData_TabPage3.UseVisualStyleBackColor = True
        '
        'EepData_Page3_GroupBox
        '
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_DrCloseBtn_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_DrCloseBtn_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_CarChime_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_CarChime_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_HallChime_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_HallChime_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_ParkingOpe_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_ParkingOpe_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_ParkingSW_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_ParkingSW_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_ParkingFL_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_ParkingFL_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_ParkingFL_ForR_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_ParkingFL_ForR_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_EscapeOpe_ForR_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_EscapeOpe_ForR_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_PhotoEye_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_PhotoEye_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_SafetyShoe_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_SafetyShoe_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_EscapeOpe_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_EscapeOpe_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_EscapeFL_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_EscapeFL_TextBox)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_Overbalance_Label)
        Me.EepData_Page3_GroupBox.Controls.Add(Me.EepData_Overbalance_TextBox)
        Me.EepData_Page3_GroupBox.Location = New System.Drawing.Point(6, 3)
        Me.EepData_Page3_GroupBox.Name = "EepData_Page3_GroupBox"
        Me.EepData_Page3_GroupBox.Size = New System.Drawing.Size(629, 515)
        Me.EepData_Page3_GroupBox.TabIndex = 169
        Me.EepData_Page3_GroupBox.TabStop = False
        '
        'EepData_DrCloseBtn_Label
        '
        Me.EepData_DrCloseBtn_Label.AutoSize = True
        Me.EepData_DrCloseBtn_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrCloseBtn_Label.Location = New System.Drawing.Point(12, 18)
        Me.EepData_DrCloseBtn_Label.Name = "EepData_DrCloseBtn_Label"
        Me.EepData_DrCloseBtn_Label.Size = New System.Drawing.Size(139, 16)
        Me.EepData_DrCloseBtn_Label.TabIndex = 142
        Me.EepData_DrCloseBtn_Label.Text = "DOOR CLOSE BUTTON"
        '
        'EepData_DrCloseBtn_TextBox
        '
        Me.EepData_DrCloseBtn_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_DrCloseBtn_TextBox.Location = New System.Drawing.Point(224, 15)
        Me.EepData_DrCloseBtn_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_DrCloseBtn_TextBox.MaxLength = 999
        Me.EepData_DrCloseBtn_TextBox.Multiline = True
        Me.EepData_DrCloseBtn_TextBox.Name = "EepData_DrCloseBtn_TextBox"
        Me.EepData_DrCloseBtn_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_DrCloseBtn_TextBox.TabIndex = 117
        '
        'EepData_CarChime_Label
        '
        Me.EepData_CarChime_Label.AutoSize = True
        Me.EepData_CarChime_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_CarChime_Label.Location = New System.Drawing.Point(12, 56)
        Me.EepData_CarChime_Label.Name = "EepData_CarChime_Label"
        Me.EepData_CarChime_Label.Size = New System.Drawing.Size(74, 16)
        Me.EepData_CarChime_Label.TabIndex = 144
        Me.EepData_CarChime_Label.Text = "CAR CHIME"
        '
        'EepData_CarChime_TextBox
        '
        Me.EepData_CarChime_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_CarChime_TextBox.Location = New System.Drawing.Point(224, 53)
        Me.EepData_CarChime_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_CarChime_TextBox.MaxLength = 999
        Me.EepData_CarChime_TextBox.Multiline = True
        Me.EepData_CarChime_TextBox.Name = "EepData_CarChime_TextBox"
        Me.EepData_CarChime_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_CarChime_TextBox.TabIndex = 119
        '
        'EepData_HallChime_Label
        '
        Me.EepData_HallChime_Label.AutoSize = True
        Me.EepData_HallChime_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HallChime_Label.Location = New System.Drawing.Point(12, 94)
        Me.EepData_HallChime_Label.Name = "EepData_HallChime_Label"
        Me.EepData_HallChime_Label.Size = New System.Drawing.Size(79, 16)
        Me.EepData_HallChime_Label.TabIndex = 146
        Me.EepData_HallChime_Label.Text = "HALL CHIME"
        '
        'EepData_HallChime_TextBox
        '
        Me.EepData_HallChime_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HallChime_TextBox.Location = New System.Drawing.Point(224, 91)
        Me.EepData_HallChime_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_HallChime_TextBox.MaxLength = 999
        Me.EepData_HallChime_TextBox.Multiline = True
        Me.EepData_HallChime_TextBox.Name = "EepData_HallChime_TextBox"
        Me.EepData_HallChime_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_HallChime_TextBox.TabIndex = 121
        '
        'EepData_ParkingOpe_Label
        '
        Me.EepData_ParkingOpe_Label.AutoSize = True
        Me.EepData_ParkingOpe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_ParkingOpe_Label.Location = New System.Drawing.Point(12, 132)
        Me.EepData_ParkingOpe_Label.Name = "EepData_ParkingOpe_Label"
        Me.EepData_ParkingOpe_Label.Size = New System.Drawing.Size(133, 16)
        Me.EepData_ParkingOpe_Label.TabIndex = 148
        Me.EepData_ParkingOpe_Label.Text = "PARKING OPERATION"
        '
        'EepData_ParkingOpe_TextBox
        '
        Me.EepData_ParkingOpe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_ParkingOpe_TextBox.Location = New System.Drawing.Point(224, 129)
        Me.EepData_ParkingOpe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_ParkingOpe_TextBox.MaxLength = 999
        Me.EepData_ParkingOpe_TextBox.Multiline = True
        Me.EepData_ParkingOpe_TextBox.Name = "EepData_ParkingOpe_TextBox"
        Me.EepData_ParkingOpe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_ParkingOpe_TextBox.TabIndex = 123
        '
        'EepData_ParkingSW_Label
        '
        Me.EepData_ParkingSW_Label.AutoSize = True
        Me.EepData_ParkingSW_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_ParkingSW_Label.Location = New System.Drawing.Point(12, 170)
        Me.EepData_ParkingSW_Label.Name = "EepData_ParkingSW_Label"
        Me.EepData_ParkingSW_Label.Size = New System.Drawing.Size(82, 16)
        Me.EepData_ParkingSW_Label.TabIndex = 150
        Me.EepData_ParkingSW_Label.Text = "PARKIGN SW"
        '
        'EepData_ParkingSW_TextBox
        '
        Me.EepData_ParkingSW_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_ParkingSW_TextBox.Location = New System.Drawing.Point(224, 167)
        Me.EepData_ParkingSW_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_ParkingSW_TextBox.MaxLength = 999
        Me.EepData_ParkingSW_TextBox.Multiline = True
        Me.EepData_ParkingSW_TextBox.Name = "EepData_ParkingSW_TextBox"
        Me.EepData_ParkingSW_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_ParkingSW_TextBox.TabIndex = 125
        '
        'EepData_ParkingFL_Label
        '
        Me.EepData_ParkingFL_Label.AutoSize = True
        Me.EepData_ParkingFL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_ParkingFL_Label.Location = New System.Drawing.Point(12, 208)
        Me.EepData_ParkingFL_Label.Name = "EepData_ParkingFL_Label"
        Me.EepData_ParkingFL_Label.Size = New System.Drawing.Size(103, 16)
        Me.EepData_ParkingFL_Label.TabIndex = 152
        Me.EepData_ParkingFL_Label.Text = "PARKING FLOOR"
        '
        'EepData_ParkingFL_TextBox
        '
        Me.EepData_ParkingFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_ParkingFL_TextBox.Location = New System.Drawing.Point(224, 205)
        Me.EepData_ParkingFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_ParkingFL_TextBox.MaxLength = 999
        Me.EepData_ParkingFL_TextBox.Multiline = True
        Me.EepData_ParkingFL_TextBox.Name = "EepData_ParkingFL_TextBox"
        Me.EepData_ParkingFL_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_ParkingFL_TextBox.TabIndex = 127
        '
        'EepData_ParkingFL_ForR_Label
        '
        Me.EepData_ParkingFL_ForR_Label.AutoSize = True
        Me.EepData_ParkingFL_ForR_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_ParkingFL_ForR_Label.Location = New System.Drawing.Point(12, 246)
        Me.EepData_ParkingFL_ForR_Label.Name = "EepData_ParkingFL_ForR_Label"
        Me.EepData_ParkingFL_ForR_Label.Size = New System.Drawing.Size(181, 32)
        Me.EepData_ParkingFL_ForR_Label.TabIndex = 154
        Me.EepData_ParkingFL_ForR_Label.Text = "FRONT OR REAR OF PARKING " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "FLOOR"
        '
        'EepData_ParkingFL_ForR_TextBox
        '
        Me.EepData_ParkingFL_ForR_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_ParkingFL_ForR_TextBox.Location = New System.Drawing.Point(224, 243)
        Me.EepData_ParkingFL_ForR_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_ParkingFL_ForR_TextBox.MaxLength = 999
        Me.EepData_ParkingFL_ForR_TextBox.Multiline = True
        Me.EepData_ParkingFL_ForR_TextBox.Name = "EepData_ParkingFL_ForR_TextBox"
        Me.EepData_ParkingFL_ForR_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_ParkingFL_ForR_TextBox.TabIndex = 129
        '
        'EepData_EscapeOpe_ForR_Label
        '
        Me.EepData_EscapeOpe_ForR_Label.AutoSize = True
        Me.EepData_EscapeOpe_ForR_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EscapeOpe_ForR_Label.Location = New System.Drawing.Point(12, 436)
        Me.EepData_EscapeOpe_ForR_Label.Name = "EepData_EscapeOpe_ForR_Label"
        Me.EepData_EscapeOpe_ForR_Label.Size = New System.Drawing.Size(126, 32)
        Me.EepData_EscapeOpe_ForR_Label.TabIndex = 164
        Me.EepData_EscapeOpe_ForR_Label.Text = "FRONT OR REAR OF " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "ESCAPE FLOOR"
        '
        'EepData_EscapeOpe_ForR_TextBox
        '
        Me.EepData_EscapeOpe_ForR_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EscapeOpe_ForR_TextBox.Location = New System.Drawing.Point(224, 433)
        Me.EepData_EscapeOpe_ForR_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_EscapeOpe_ForR_TextBox.MaxLength = 999
        Me.EepData_EscapeOpe_ForR_TextBox.Multiline = True
        Me.EepData_EscapeOpe_ForR_TextBox.Name = "EepData_EscapeOpe_ForR_TextBox"
        Me.EepData_EscapeOpe_ForR_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_EscapeOpe_ForR_TextBox.TabIndex = 139
        '
        'EepData_PhotoEye_Label
        '
        Me.EepData_PhotoEye_Label.AutoSize = True
        Me.EepData_PhotoEye_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_PhotoEye_Label.Location = New System.Drawing.Point(12, 284)
        Me.EepData_PhotoEye_Label.Name = "EepData_PhotoEye_Label"
        Me.EepData_PhotoEye_Label.Size = New System.Drawing.Size(75, 16)
        Me.EepData_PhotoEye_Label.TabIndex = 156
        Me.EepData_PhotoEye_Label.Text = "PHOTO EYE"
        '
        'EepData_PhotoEye_TextBox
        '
        Me.EepData_PhotoEye_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_PhotoEye_TextBox.Location = New System.Drawing.Point(224, 281)
        Me.EepData_PhotoEye_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_PhotoEye_TextBox.MaxLength = 999
        Me.EepData_PhotoEye_TextBox.Multiline = True
        Me.EepData_PhotoEye_TextBox.Name = "EepData_PhotoEye_TextBox"
        Me.EepData_PhotoEye_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_PhotoEye_TextBox.TabIndex = 131
        '
        'EepData_SafetyShoe_Label
        '
        Me.EepData_SafetyShoe_Label.AutoSize = True
        Me.EepData_SafetyShoe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_SafetyShoe_Label.Location = New System.Drawing.Point(12, 322)
        Me.EepData_SafetyShoe_Label.Name = "EepData_SafetyShoe_Label"
        Me.EepData_SafetyShoe_Label.Size = New System.Drawing.Size(86, 16)
        Me.EepData_SafetyShoe_Label.TabIndex = 158
        Me.EepData_SafetyShoe_Label.Text = "SAFETY SHOE"
        '
        'EepData_SafetyShoe_TextBox
        '
        Me.EepData_SafetyShoe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_SafetyShoe_TextBox.Location = New System.Drawing.Point(224, 319)
        Me.EepData_SafetyShoe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_SafetyShoe_TextBox.MaxLength = 999
        Me.EepData_SafetyShoe_TextBox.Multiline = True
        Me.EepData_SafetyShoe_TextBox.Name = "EepData_SafetyShoe_TextBox"
        Me.EepData_SafetyShoe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_SafetyShoe_TextBox.TabIndex = 133
        '
        'EepData_EscapeOpe_Label
        '
        Me.EepData_EscapeOpe_Label.AutoSize = True
        Me.EepData_EscapeOpe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EscapeOpe_Label.Location = New System.Drawing.Point(12, 360)
        Me.EepData_EscapeOpe_Label.Name = "EepData_EscapeOpe_Label"
        Me.EepData_EscapeOpe_Label.Size = New System.Drawing.Size(188, 32)
        Me.EepData_EscapeOpe_Label.TabIndex = 160
        Me.EepData_EscapeOpe_Label.Text = "RETURN TO ESCAPE(HOMING) " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "FLOOR"
        '
        'EepData_EscapeOpe_TextBox
        '
        Me.EepData_EscapeOpe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EscapeOpe_TextBox.Location = New System.Drawing.Point(224, 357)
        Me.EepData_EscapeOpe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_EscapeOpe_TextBox.MaxLength = 999
        Me.EepData_EscapeOpe_TextBox.Multiline = True
        Me.EepData_EscapeOpe_TextBox.Name = "EepData_EscapeOpe_TextBox"
        Me.EepData_EscapeOpe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_EscapeOpe_TextBox.TabIndex = 135
        '
        'EepData_EscapeFL_Label
        '
        Me.EepData_EscapeFL_Label.AutoSize = True
        Me.EepData_EscapeFL_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EscapeFL_Label.Location = New System.Drawing.Point(12, 398)
        Me.EepData_EscapeFL_Label.Name = "EepData_EscapeFL_Label"
        Me.EepData_EscapeFL_Label.Size = New System.Drawing.Size(95, 16)
        Me.EepData_EscapeFL_Label.TabIndex = 162
        Me.EepData_EscapeFL_Label.Text = "ESCAPE FLOOR"
        '
        'EepData_EscapeFL_TextBox
        '
        Me.EepData_EscapeFL_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EscapeFL_TextBox.Location = New System.Drawing.Point(224, 395)
        Me.EepData_EscapeFL_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_EscapeFL_TextBox.MaxLength = 999
        Me.EepData_EscapeFL_TextBox.Multiline = True
        Me.EepData_EscapeFL_TextBox.Name = "EepData_EscapeFL_TextBox"
        Me.EepData_EscapeFL_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_EscapeFL_TextBox.TabIndex = 137
        '
        'EepData_Overbalance_Label
        '
        Me.EepData_Overbalance_Label.AutoSize = True
        Me.EepData_Overbalance_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Overbalance_Label.Location = New System.Drawing.Point(12, 474)
        Me.EepData_Overbalance_Label.Name = "EepData_Overbalance_Label"
        Me.EepData_Overbalance_Label.Size = New System.Drawing.Size(95, 16)
        Me.EepData_Overbalance_Label.TabIndex = 166
        Me.EepData_Overbalance_Label.Text = "OVERBALANCE"
        '
        'EepData_Overbalance_TextBox
        '
        Me.EepData_Overbalance_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Overbalance_TextBox.Location = New System.Drawing.Point(224, 471)
        Me.EepData_Overbalance_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Overbalance_TextBox.MaxLength = 999
        Me.EepData_Overbalance_TextBox.Multiline = True
        Me.EepData_Overbalance_TextBox.Name = "EepData_Overbalance_TextBox"
        Me.EepData_Overbalance_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Overbalance_TextBox.TabIndex = 141
        '
        'EepData_TabPage4
        '
        Me.EepData_TabPage4.Controls.Add(Me.EepData_Page4_GroupBox)
        Me.EepData_TabPage4.Location = New System.Drawing.Point(4, 25)
        Me.EepData_TabPage4.Name = "EepData_TabPage4"
        Me.EepData_TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.EepData_TabPage4.Size = New System.Drawing.Size(641, 524)
        Me.EepData_TabPage4.TabIndex = 3
        Me.EepData_TabPage4.Text = "Page 4"
        Me.EepData_TabPage4.UseVisualStyleBackColor = True
        '
        'EepData_Page4_GroupBox
        '
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_SheaveDia_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_SheaveDia_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MachineType_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MachineType_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_Gear_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_Gear_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_Inverter_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_Inverter_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MotorPole_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MotorPole_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MotorVoltage_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MotorVoltage_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MotorCapacity_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MotorCapacity_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MotorDirection_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_MotorDirection_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_Encoder_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_Encoder_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_FireOpe_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_FireOpe_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_FMSOpe_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_FMSOpe_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_FMSSW_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_FMSSW_TextBox)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_EmerOpe_Label)
        Me.EepData_Page4_GroupBox.Controls.Add(Me.EepData_EmerOpe_TextBox)
        Me.EepData_Page4_GroupBox.Location = New System.Drawing.Point(6, 3)
        Me.EepData_Page4_GroupBox.Name = "EepData_Page4_GroupBox"
        Me.EepData_Page4_GroupBox.Size = New System.Drawing.Size(629, 515)
        Me.EepData_Page4_GroupBox.TabIndex = 169
        Me.EepData_Page4_GroupBox.TabStop = False
        '
        'EepData_SheaveDia_Label
        '
        Me.EepData_SheaveDia_Label.AutoSize = True
        Me.EepData_SheaveDia_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_SheaveDia_Label.Location = New System.Drawing.Point(12, 18)
        Me.EepData_SheaveDia_Label.Name = "EepData_SheaveDia_Label"
        Me.EepData_SheaveDia_Label.Size = New System.Drawing.Size(77, 16)
        Me.EepData_SheaveDia_Label.TabIndex = 142
        Me.EepData_SheaveDia_Label.Text = "SHEAVE DIA"
        '
        'EepData_SheaveDia_TextBox
        '
        Me.EepData_SheaveDia_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_SheaveDia_TextBox.Location = New System.Drawing.Point(224, 15)
        Me.EepData_SheaveDia_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_SheaveDia_TextBox.MaxLength = 999
        Me.EepData_SheaveDia_TextBox.Multiline = True
        Me.EepData_SheaveDia_TextBox.Name = "EepData_SheaveDia_TextBox"
        Me.EepData_SheaveDia_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_SheaveDia_TextBox.TabIndex = 117
        '
        'EepData_MachineType_Label
        '
        Me.EepData_MachineType_Label.AutoSize = True
        Me.EepData_MachineType_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MachineType_Label.Location = New System.Drawing.Point(12, 56)
        Me.EepData_MachineType_Label.Name = "EepData_MachineType_Label"
        Me.EepData_MachineType_Label.Size = New System.Drawing.Size(99, 16)
        Me.EepData_MachineType_Label.TabIndex = 144
        Me.EepData_MachineType_Label.Text = "MACHINE TYPE "
        '
        'EepData_MachineType_TextBox
        '
        Me.EepData_MachineType_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MachineType_TextBox.Location = New System.Drawing.Point(224, 53)
        Me.EepData_MachineType_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_MachineType_TextBox.MaxLength = 999
        Me.EepData_MachineType_TextBox.Multiline = True
        Me.EepData_MachineType_TextBox.Name = "EepData_MachineType_TextBox"
        Me.EepData_MachineType_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_MachineType_TextBox.TabIndex = 119
        '
        'EepData_Gear_Label
        '
        Me.EepData_Gear_Label.AutoSize = True
        Me.EepData_Gear_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Gear_Label.Location = New System.Drawing.Point(12, 94)
        Me.EepData_Gear_Label.Name = "EepData_Gear_Label"
        Me.EepData_Gear_Label.Size = New System.Drawing.Size(79, 16)
        Me.EepData_Gear_Label.TabIndex = 146
        Me.EepData_Gear_Label.Text = "GEAR RATIO"
        '
        'EepData_Gear_TextBox
        '
        Me.EepData_Gear_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Gear_TextBox.Location = New System.Drawing.Point(224, 91)
        Me.EepData_Gear_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Gear_TextBox.MaxLength = 999
        Me.EepData_Gear_TextBox.Multiline = True
        Me.EepData_Gear_TextBox.Name = "EepData_Gear_TextBox"
        Me.EepData_Gear_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Gear_TextBox.TabIndex = 121
        '
        'EepData_Inverter_Label
        '
        Me.EepData_Inverter_Label.AutoSize = True
        Me.EepData_Inverter_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Inverter_Label.Location = New System.Drawing.Point(12, 132)
        Me.EepData_Inverter_Label.Name = "EepData_Inverter_Label"
        Me.EepData_Inverter_Label.Size = New System.Drawing.Size(122, 16)
        Me.EepData_Inverter_Label.TabIndex = 148
        Me.EepData_Inverter_Label.Text = "INVERTER PART NO"
        '
        'EepData_Inverter_TextBox
        '
        Me.EepData_Inverter_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Inverter_TextBox.Location = New System.Drawing.Point(224, 129)
        Me.EepData_Inverter_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Inverter_TextBox.MaxLength = 999
        Me.EepData_Inverter_TextBox.Multiline = True
        Me.EepData_Inverter_TextBox.Name = "EepData_Inverter_TextBox"
        Me.EepData_Inverter_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Inverter_TextBox.TabIndex = 123
        '
        'EepData_MotorPole_Label
        '
        Me.EepData_MotorPole_Label.AutoSize = True
        Me.EepData_MotorPole_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MotorPole_Label.Location = New System.Drawing.Point(12, 170)
        Me.EepData_MotorPole_Label.Name = "EepData_MotorPole_Label"
        Me.EepData_MotorPole_Label.Size = New System.Drawing.Size(91, 16)
        Me.EepData_MotorPole_Label.TabIndex = 150
        Me.EepData_MotorPole_Label.Text = "MOTOR POLE "
        '
        'EepData_MotorPole_TextBox
        '
        Me.EepData_MotorPole_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MotorPole_TextBox.Location = New System.Drawing.Point(224, 167)
        Me.EepData_MotorPole_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_MotorPole_TextBox.MaxLength = 999
        Me.EepData_MotorPole_TextBox.Multiline = True
        Me.EepData_MotorPole_TextBox.Name = "EepData_MotorPole_TextBox"
        Me.EepData_MotorPole_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_MotorPole_TextBox.TabIndex = 125
        '
        'EepData_MotorVoltage_Label
        '
        Me.EepData_MotorVoltage_Label.AutoSize = True
        Me.EepData_MotorVoltage_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MotorVoltage_Label.Location = New System.Drawing.Point(12, 208)
        Me.EepData_MotorVoltage_Label.Name = "EepData_MotorVoltage_Label"
        Me.EepData_MotorVoltage_Label.Size = New System.Drawing.Size(113, 16)
        Me.EepData_MotorVoltage_Label.TabIndex = 152
        Me.EepData_MotorVoltage_Label.Text = "MOTOR VOLTAGE"
        '
        'EepData_MotorVoltage_TextBox
        '
        Me.EepData_MotorVoltage_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MotorVoltage_TextBox.Location = New System.Drawing.Point(224, 205)
        Me.EepData_MotorVoltage_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_MotorVoltage_TextBox.MaxLength = 999
        Me.EepData_MotorVoltage_TextBox.Multiline = True
        Me.EepData_MotorVoltage_TextBox.Name = "EepData_MotorVoltage_TextBox"
        Me.EepData_MotorVoltage_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_MotorVoltage_TextBox.TabIndex = 127
        '
        'EepData_MotorCapacity_Label
        '
        Me.EepData_MotorCapacity_Label.AutoSize = True
        Me.EepData_MotorCapacity_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MotorCapacity_Label.Location = New System.Drawing.Point(12, 246)
        Me.EepData_MotorCapacity_Label.Name = "EepData_MotorCapacity_Label"
        Me.EepData_MotorCapacity_Label.Size = New System.Drawing.Size(114, 16)
        Me.EepData_MotorCapacity_Label.TabIndex = 154
        Me.EepData_MotorCapacity_Label.Text = "MOTOR CAPACITY"
        '
        'EepData_MotorCapacity_TextBox
        '
        Me.EepData_MotorCapacity_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MotorCapacity_TextBox.Location = New System.Drawing.Point(224, 243)
        Me.EepData_MotorCapacity_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_MotorCapacity_TextBox.MaxLength = 999
        Me.EepData_MotorCapacity_TextBox.Multiline = True
        Me.EepData_MotorCapacity_TextBox.Name = "EepData_MotorCapacity_TextBox"
        Me.EepData_MotorCapacity_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_MotorCapacity_TextBox.TabIndex = 129
        '
        'EepData_MotorDirection_Label
        '
        Me.EepData_MotorDirection_Label.AutoSize = True
        Me.EepData_MotorDirection_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MotorDirection_Label.Location = New System.Drawing.Point(12, 284)
        Me.EepData_MotorDirection_Label.Name = "EepData_MotorDirection_Label"
        Me.EepData_MotorDirection_Label.Size = New System.Drawing.Size(126, 16)
        Me.EepData_MotorDirection_Label.TabIndex = 156
        Me.EepData_MotorDirection_Label.Text = " MOTOR DIRECTION"
        '
        'EepData_MotorDirection_TextBox
        '
        Me.EepData_MotorDirection_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_MotorDirection_TextBox.Location = New System.Drawing.Point(224, 281)
        Me.EepData_MotorDirection_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_MotorDirection_TextBox.MaxLength = 999
        Me.EepData_MotorDirection_TextBox.Multiline = True
        Me.EepData_MotorDirection_TextBox.Name = "EepData_MotorDirection_TextBox"
        Me.EepData_MotorDirection_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_MotorDirection_TextBox.TabIndex = 131
        '
        'EepData_Encoder_Label
        '
        Me.EepData_Encoder_Label.AutoSize = True
        Me.EepData_Encoder_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Encoder_Label.Location = New System.Drawing.Point(12, 322)
        Me.EepData_Encoder_Label.Name = "EepData_Encoder_Label"
        Me.EepData_Encoder_Label.Size = New System.Drawing.Size(99, 16)
        Me.EepData_Encoder_Label.TabIndex = 158
        Me.EepData_Encoder_Label.Text = "ENCODER PULS"
        '
        'EepData_Encoder_TextBox
        '
        Me.EepData_Encoder_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Encoder_TextBox.Location = New System.Drawing.Point(224, 319)
        Me.EepData_Encoder_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Encoder_TextBox.MaxLength = 999
        Me.EepData_Encoder_TextBox.Multiline = True
        Me.EepData_Encoder_TextBox.Name = "EepData_Encoder_TextBox"
        Me.EepData_Encoder_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Encoder_TextBox.TabIndex = 133
        '
        'EepData_FireOpe_Label
        '
        Me.EepData_FireOpe_Label.AutoSize = True
        Me.EepData_FireOpe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FireOpe_Label.Location = New System.Drawing.Point(12, 360)
        Me.EepData_FireOpe_Label.Name = "EepData_FireOpe_Label"
        Me.EepData_FireOpe_Label.Size = New System.Drawing.Size(105, 16)
        Me.EepData_FireOpe_Label.TabIndex = 160
        Me.EepData_FireOpe_Label.Text = "FIRE OPERATION"
        '
        'EepData_FireOpe_TextBox
        '
        Me.EepData_FireOpe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FireOpe_TextBox.Location = New System.Drawing.Point(224, 357)
        Me.EepData_FireOpe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_FireOpe_TextBox.MaxLength = 999
        Me.EepData_FireOpe_TextBox.Multiline = True
        Me.EepData_FireOpe_TextBox.Name = "EepData_FireOpe_TextBox"
        Me.EepData_FireOpe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_FireOpe_TextBox.TabIndex = 135
        '
        'EepData_FMSOpe_Label
        '
        Me.EepData_FMSOpe_Label.AutoSize = True
        Me.EepData_FMSOpe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FMSOpe_Label.Location = New System.Drawing.Point(12, 398)
        Me.EepData_FMSOpe_Label.Name = "EepData_FMSOpe_Label"
        Me.EepData_FMSOpe_Label.Size = New System.Drawing.Size(170, 16)
        Me.EepData_FMSOpe_Label.TabIndex = 162
        Me.EepData_FMSOpe_Label.Text = "FIREMAN'S LIFT OPERATION"
        '
        'EepData_FMSOpe_TextBox
        '
        Me.EepData_FMSOpe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FMSOpe_TextBox.Location = New System.Drawing.Point(224, 395)
        Me.EepData_FMSOpe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_FMSOpe_TextBox.MaxLength = 999
        Me.EepData_FMSOpe_TextBox.Multiline = True
        Me.EepData_FMSOpe_TextBox.Name = "EepData_FMSOpe_TextBox"
        Me.EepData_FMSOpe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_FMSOpe_TextBox.TabIndex = 137
        '
        'EepData_FMSSW_Label
        '
        Me.EepData_FMSSW_Label.AutoSize = True
        Me.EepData_FMSSW_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FMSSW_Label.Location = New System.Drawing.Point(12, 436)
        Me.EepData_FMSSW_Label.Name = "EepData_FMSSW_Label"
        Me.EepData_FMSSW_Label.Size = New System.Drawing.Size(119, 16)
        Me.EepData_FMSSW_Label.TabIndex = 164
        Me.EepData_FMSSW_Label.Text = "FIREMAN'S LIFT SW"
        '
        'EepData_FMSSW_TextBox
        '
        Me.EepData_FMSSW_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FMSSW_TextBox.Location = New System.Drawing.Point(224, 433)
        Me.EepData_FMSSW_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_FMSSW_TextBox.MaxLength = 999
        Me.EepData_FMSSW_TextBox.Multiline = True
        Me.EepData_FMSSW_TextBox.Name = "EepData_FMSSW_TextBox"
        Me.EepData_FMSSW_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_FMSSW_TextBox.TabIndex = 139
        '
        'EepData_EmerOpe_Label
        '
        Me.EepData_EmerOpe_Label.AutoSize = True
        Me.EepData_EmerOpe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EmerOpe_Label.Location = New System.Drawing.Point(12, 474)
        Me.EepData_EmerOpe_Label.Name = "EepData_EmerOpe_Label"
        Me.EepData_EmerOpe_Label.Size = New System.Drawing.Size(162, 16)
        Me.EepData_EmerOpe_Label.TabIndex = 166
        Me.EepData_EmerOpe_Label.Text = "EMER POWER OPERATION"
        '
        'EepData_EmerOpe_TextBox
        '
        Me.EepData_EmerOpe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_EmerOpe_TextBox.Location = New System.Drawing.Point(224, 471)
        Me.EepData_EmerOpe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_EmerOpe_TextBox.MaxLength = 999
        Me.EepData_EmerOpe_TextBox.Multiline = True
        Me.EepData_EmerOpe_TextBox.Name = "EepData_EmerOpe_TextBox"
        Me.EepData_EmerOpe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_EmerOpe_TextBox.TabIndex = 141
        '
        'EepData_TabPage5
        '
        Me.EepData_TabPage5.Controls.Add(Me.EepData_Page5_GroupBox)
        Me.EepData_TabPage5.Location = New System.Drawing.Point(4, 25)
        Me.EepData_TabPage5.Name = "EepData_TabPage5"
        Me.EepData_TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.EepData_TabPage5.Size = New System.Drawing.Size(641, 524)
        Me.EepData_TabPage5.TabIndex = 4
        Me.EepData_TabPage5.Text = "Page 5"
        Me.EepData_TabPage5.UseVisualStyleBackColor = True
        '
        'EepData_Page5_GroupBox
        '
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_FloodOpe_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_FloodOpe_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_Vonic_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_Vonic_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_HIN1_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_HIN1_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_HIN2_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_HIN2_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_HIN3_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_HIN3_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_HIN4_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_HIN4_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_SCOB_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_SCOB_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_WCOB_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_WCOB_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_WSCOB_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_WSCOB_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_WCOB_Spec_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_WCOB_Spec_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_FRDr_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_FRDr_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_AttOpe_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_AttOpe_TextBox)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_Rope_Label)
        Me.EepData_Page5_GroupBox.Controls.Add(Me.EepData_Rope_TextBox)
        Me.EepData_Page5_GroupBox.Location = New System.Drawing.Point(6, 3)
        Me.EepData_Page5_GroupBox.Name = "EepData_Page5_GroupBox"
        Me.EepData_Page5_GroupBox.Size = New System.Drawing.Size(629, 515)
        Me.EepData_Page5_GroupBox.TabIndex = 169
        Me.EepData_Page5_GroupBox.TabStop = False
        '
        'EepData_FloodOpe_Label
        '
        Me.EepData_FloodOpe_Label.AutoSize = True
        Me.EepData_FloodOpe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FloodOpe_Label.Location = New System.Drawing.Point(12, 18)
        Me.EepData_FloodOpe_Label.Name = "EepData_FloodOpe_Label"
        Me.EepData_FloodOpe_Label.Size = New System.Drawing.Size(122, 16)
        Me.EepData_FloodOpe_Label.TabIndex = 142
        Me.EepData_FloodOpe_Label.Text = "FLOOD OPERATION"
        '
        'EepData_FloodOpe_TextBox
        '
        Me.EepData_FloodOpe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FloodOpe_TextBox.Location = New System.Drawing.Point(224, 15)
        Me.EepData_FloodOpe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_FloodOpe_TextBox.MaxLength = 999
        Me.EepData_FloodOpe_TextBox.Multiline = True
        Me.EepData_FloodOpe_TextBox.Name = "EepData_FloodOpe_TextBox"
        Me.EepData_FloodOpe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_FloodOpe_TextBox.TabIndex = 117
        '
        'EepData_Vonic_Label
        '
        Me.EepData_Vonic_Label.AutoSize = True
        Me.EepData_Vonic_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Vonic_Label.Location = New System.Drawing.Point(12, 56)
        Me.EepData_Vonic_Label.Name = "EepData_Vonic_Label"
        Me.EepData_Vonic_Label.Size = New System.Drawing.Size(194, 32)
        Me.EepData_Vonic_Label.TabIndex = 144
        Me.EepData_Vonic_Label.Text = "AUTOMATIC ANNOUNCEMENT " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "SYSTEM"
        '
        'EepData_Vonic_TextBox
        '
        Me.EepData_Vonic_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Vonic_TextBox.Location = New System.Drawing.Point(224, 53)
        Me.EepData_Vonic_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Vonic_TextBox.MaxLength = 999
        Me.EepData_Vonic_TextBox.Multiline = True
        Me.EepData_Vonic_TextBox.Name = "EepData_Vonic_TextBox"
        Me.EepData_Vonic_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Vonic_TextBox.TabIndex = 119
        '
        'EepData_HIN1_Label
        '
        Me.EepData_HIN1_Label.AutoSize = True
        Me.EepData_HIN1_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HIN1_Label.Location = New System.Drawing.Point(12, 94)
        Me.EepData_HIN1_Label.Name = "EepData_HIN1_Label"
        Me.EepData_HIN1_Label.Size = New System.Drawing.Size(159, 16)
        Me.EepData_HIN1_Label.TabIndex = 146
        Me.EepData_HIN1_Label.Text = "HALL INDICATOR SIGNAL1"
        '
        'EepData_HIN1_TextBox
        '
        Me.EepData_HIN1_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HIN1_TextBox.Location = New System.Drawing.Point(224, 91)
        Me.EepData_HIN1_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_HIN1_TextBox.MaxLength = 999
        Me.EepData_HIN1_TextBox.Multiline = True
        Me.EepData_HIN1_TextBox.Name = "EepData_HIN1_TextBox"
        Me.EepData_HIN1_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_HIN1_TextBox.TabIndex = 121
        '
        'EepData_HIN2_Label
        '
        Me.EepData_HIN2_Label.AutoSize = True
        Me.EepData_HIN2_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HIN2_Label.Location = New System.Drawing.Point(12, 132)
        Me.EepData_HIN2_Label.Name = "EepData_HIN2_Label"
        Me.EepData_HIN2_Label.Size = New System.Drawing.Size(159, 16)
        Me.EepData_HIN2_Label.TabIndex = 148
        Me.EepData_HIN2_Label.Text = "HALL INDICATOR SIGNAL2"
        '
        'EepData_HIN2_TextBox
        '
        Me.EepData_HIN2_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HIN2_TextBox.Location = New System.Drawing.Point(224, 129)
        Me.EepData_HIN2_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_HIN2_TextBox.MaxLength = 999
        Me.EepData_HIN2_TextBox.Multiline = True
        Me.EepData_HIN2_TextBox.Name = "EepData_HIN2_TextBox"
        Me.EepData_HIN2_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_HIN2_TextBox.TabIndex = 123
        '
        'EepData_HIN3_Label
        '
        Me.EepData_HIN3_Label.AutoSize = True
        Me.EepData_HIN3_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HIN3_Label.Location = New System.Drawing.Point(12, 170)
        Me.EepData_HIN3_Label.Name = "EepData_HIN3_Label"
        Me.EepData_HIN3_Label.Size = New System.Drawing.Size(159, 16)
        Me.EepData_HIN3_Label.TabIndex = 150
        Me.EepData_HIN3_Label.Text = "HALL INDICATOR SIGNAL3"
        '
        'EepData_HIN3_TextBox
        '
        Me.EepData_HIN3_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HIN3_TextBox.Location = New System.Drawing.Point(224, 167)
        Me.EepData_HIN3_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_HIN3_TextBox.MaxLength = 999
        Me.EepData_HIN3_TextBox.Multiline = True
        Me.EepData_HIN3_TextBox.Name = "EepData_HIN3_TextBox"
        Me.EepData_HIN3_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_HIN3_TextBox.TabIndex = 125
        '
        'EepData_HIN4_Label
        '
        Me.EepData_HIN4_Label.AutoSize = True
        Me.EepData_HIN4_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HIN4_Label.Location = New System.Drawing.Point(12, 208)
        Me.EepData_HIN4_Label.Name = "EepData_HIN4_Label"
        Me.EepData_HIN4_Label.Size = New System.Drawing.Size(159, 16)
        Me.EepData_HIN4_Label.TabIndex = 152
        Me.EepData_HIN4_Label.Text = "HALL INDICATOR SIGNAL4"
        '
        'EepData_HIN4_TextBox
        '
        Me.EepData_HIN4_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_HIN4_TextBox.Location = New System.Drawing.Point(224, 205)
        Me.EepData_HIN4_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_HIN4_TextBox.MaxLength = 999
        Me.EepData_HIN4_TextBox.Multiline = True
        Me.EepData_HIN4_TextBox.Name = "EepData_HIN4_TextBox"
        Me.EepData_HIN4_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_HIN4_TextBox.TabIndex = 127
        '
        'EepData_SCOB_Label
        '
        Me.EepData_SCOB_Label.AutoSize = True
        Me.EepData_SCOB_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_SCOB_Label.Location = New System.Drawing.Point(12, 246)
        Me.EepData_SCOB_Label.Name = "EepData_SCOB_Label"
        Me.EepData_SCOB_Label.Size = New System.Drawing.Size(59, 16)
        Me.EepData_SCOB_Label.TabIndex = 154
        Me.EepData_SCOB_Label.Text = "SUB COB"
        '
        'EepData_SCOB_TextBox
        '
        Me.EepData_SCOB_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_SCOB_TextBox.Location = New System.Drawing.Point(224, 243)
        Me.EepData_SCOB_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_SCOB_TextBox.MaxLength = 999
        Me.EepData_SCOB_TextBox.Multiline = True
        Me.EepData_SCOB_TextBox.Name = "EepData_SCOB_TextBox"
        Me.EepData_SCOB_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_SCOB_TextBox.TabIndex = 129
        '
        'EepData_WCOB_Label
        '
        Me.EepData_WCOB_Label.AutoSize = True
        Me.EepData_WCOB_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_WCOB_Label.Location = New System.Drawing.Point(12, 284)
        Me.EepData_WCOB_Label.Name = "EepData_WCOB_Label"
        Me.EepData_WCOB_Label.Size = New System.Drawing.Size(152, 16)
        Me.EepData_WCOB_Label.TabIndex = 156
        Me.EepData_WCOB_Label.Text = "WHEEL CHAIR MAIN COB"
        '
        'EepData_WCOB_TextBox
        '
        Me.EepData_WCOB_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_WCOB_TextBox.Location = New System.Drawing.Point(224, 281)
        Me.EepData_WCOB_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_WCOB_TextBox.MaxLength = 999
        Me.EepData_WCOB_TextBox.Multiline = True
        Me.EepData_WCOB_TextBox.Name = "EepData_WCOB_TextBox"
        Me.EepData_WCOB_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_WCOB_TextBox.TabIndex = 131
        '
        'EepData_WSCOB_Label
        '
        Me.EepData_WSCOB_Label.AutoSize = True
        Me.EepData_WSCOB_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_WSCOB_Label.Location = New System.Drawing.Point(12, 322)
        Me.EepData_WSCOB_Label.Name = "EepData_WSCOB_Label"
        Me.EepData_WSCOB_Label.Size = New System.Drawing.Size(142, 16)
        Me.EepData_WSCOB_Label.TabIndex = 158
        Me.EepData_WSCOB_Label.Text = "WHEEL CHAIR SUB COB"
        '
        'EepData_WSCOB_TextBox
        '
        Me.EepData_WSCOB_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_WSCOB_TextBox.Location = New System.Drawing.Point(224, 319)
        Me.EepData_WSCOB_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_WSCOB_TextBox.MaxLength = 999
        Me.EepData_WSCOB_TextBox.Multiline = True
        Me.EepData_WSCOB_TextBox.Name = "EepData_WSCOB_TextBox"
        Me.EepData_WSCOB_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_WSCOB_TextBox.TabIndex = 133
        '
        'EepData_WCOB_Spec_Label
        '
        Me.EepData_WCOB_Spec_Label.AutoSize = True
        Me.EepData_WCOB_Spec_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_WCOB_Spec_Label.Location = New System.Drawing.Point(12, 360)
        Me.EepData_WCOB_Spec_Label.Name = "EepData_WCOB_Spec_Label"
        Me.EepData_WCOB_Spec_Label.Size = New System.Drawing.Size(120, 16)
        Me.EepData_WCOB_Spec_Label.TabIndex = 160
        Me.EepData_WCOB_Spec_Label.Text = "WHEEL CHAIR SPEC"
        '
        'EepData_WCOB_Spec_TextBox
        '
        Me.EepData_WCOB_Spec_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_WCOB_Spec_TextBox.Location = New System.Drawing.Point(224, 357)
        Me.EepData_WCOB_Spec_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_WCOB_Spec_TextBox.MaxLength = 999
        Me.EepData_WCOB_Spec_TextBox.Multiline = True
        Me.EepData_WCOB_Spec_TextBox.Name = "EepData_WCOB_Spec_TextBox"
        Me.EepData_WCOB_Spec_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_WCOB_Spec_TextBox.TabIndex = 135
        '
        'EepData_FRDr_Label
        '
        Me.EepData_FRDr_Label.AutoSize = True
        Me.EepData_FRDr_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FRDr_Label.Location = New System.Drawing.Point(12, 398)
        Me.EepData_FRDr_Label.Name = "EepData_FRDr_Label"
        Me.EepData_FRDr_Label.Size = New System.Drawing.Size(123, 16)
        Me.EepData_FRDr_Label.TabIndex = 162
        Me.EepData_FRDr_Label.Text = "FRONT REAR DOOR"
        '
        'EepData_FRDr_TextBox
        '
        Me.EepData_FRDr_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_FRDr_TextBox.Location = New System.Drawing.Point(224, 395)
        Me.EepData_FRDr_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_FRDr_TextBox.MaxLength = 999
        Me.EepData_FRDr_TextBox.Multiline = True
        Me.EepData_FRDr_TextBox.Name = "EepData_FRDr_TextBox"
        Me.EepData_FRDr_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_FRDr_TextBox.TabIndex = 137
        '
        'EepData_AttOpe_Label
        '
        Me.EepData_AttOpe_Label.AutoSize = True
        Me.EepData_AttOpe_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_AttOpe_Label.Location = New System.Drawing.Point(12, 436)
        Me.EepData_AttOpe_Label.Name = "EepData_AttOpe_Label"
        Me.EepData_AttOpe_Label.Size = New System.Drawing.Size(154, 16)
        Me.EepData_AttOpe_Label.TabIndex = 164
        Me.EepData_AttOpe_Label.Text = "ATTENDANT OPERATION"
        '
        'EepData_AttOpe_TextBox
        '
        Me.EepData_AttOpe_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_AttOpe_TextBox.Location = New System.Drawing.Point(224, 433)
        Me.EepData_AttOpe_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_AttOpe_TextBox.MaxLength = 999
        Me.EepData_AttOpe_TextBox.Multiline = True
        Me.EepData_AttOpe_TextBox.Name = "EepData_AttOpe_TextBox"
        Me.EepData_AttOpe_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_AttOpe_TextBox.TabIndex = 139
        '
        'EepData_Rope_Label
        '
        Me.EepData_Rope_Label.AutoSize = True
        Me.EepData_Rope_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Rope_Label.Location = New System.Drawing.Point(12, 474)
        Me.EepData_Rope_Label.Name = "EepData_Rope_Label"
        Me.EepData_Rope_Label.Size = New System.Drawing.Size(55, 16)
        Me.EepData_Rope_Label.TabIndex = 166
        Me.EepData_Rope_Label.Text = "ROPING"
        '
        'EepData_Rope_TextBox
        '
        Me.EepData_Rope_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Rope_TextBox.Location = New System.Drawing.Point(224, 471)
        Me.EepData_Rope_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Rope_TextBox.MaxLength = 999
        Me.EepData_Rope_TextBox.Multiline = True
        Me.EepData_Rope_TextBox.Name = "EepData_Rope_TextBox"
        Me.EepData_Rope_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Rope_TextBox.TabIndex = 141
        '
        'EepData_TabPage6
        '
        Me.EepData_TabPage6.Controls.Add(Me.EepData_Page6_GroupBox)
        Me.EepData_TabPage6.Location = New System.Drawing.Point(4, 25)
        Me.EepData_TabPage6.Name = "EepData_TabPage6"
        Me.EepData_TabPage6.Padding = New System.Windows.Forms.Padding(3)
        Me.EepData_TabPage6.Size = New System.Drawing.Size(641, 524)
        Me.EepData_TabPage6.TabIndex = 5
        Me.EepData_TabPage6.Text = "Page 6"
        Me.EepData_TabPage6.UseVisualStyleBackColor = True
        '
        'EepData_Page6_GroupBox
        '
        Me.EepData_Page6_GroupBox.Controls.Add(Me.EepData_Travel_Label)
        Me.EepData_Page6_GroupBox.Controls.Add(Me.EepData_Travel_TextBox)
        Me.EepData_Page6_GroupBox.Controls.Add(Me.EepData_Hight_Label)
        Me.EepData_Page6_GroupBox.Controls.Add(Me.EepData_Hight_TextBox)
        Me.EepData_Page6_GroupBox.Location = New System.Drawing.Point(6, 5)
        Me.EepData_Page6_GroupBox.Name = "EepData_Page6_GroupBox"
        Me.EepData_Page6_GroupBox.Size = New System.Drawing.Size(629, 515)
        Me.EepData_Page6_GroupBox.TabIndex = 170
        Me.EepData_Page6_GroupBox.TabStop = False
        '
        'EepData_Travel_Label
        '
        Me.EepData_Travel_Label.AutoSize = True
        Me.EepData_Travel_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Travel_Label.Location = New System.Drawing.Point(12, 18)
        Me.EepData_Travel_Label.Name = "EepData_Travel_Label"
        Me.EepData_Travel_Label.Size = New System.Drawing.Size(82, 16)
        Me.EepData_Travel_Label.TabIndex = 142
        Me.EepData_Travel_Label.Text = "TRAVEL(mm)"
        '
        'EepData_Travel_TextBox
        '
        Me.EepData_Travel_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Travel_TextBox.Location = New System.Drawing.Point(224, 15)
        Me.EepData_Travel_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Travel_TextBox.MaxLength = 999
        Me.EepData_Travel_TextBox.Multiline = True
        Me.EepData_Travel_TextBox.Name = "EepData_Travel_TextBox"
        Me.EepData_Travel_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Travel_TextBox.TabIndex = 117
        '
        'EepData_Hight_Label
        '
        Me.EepData_Hight_Label.AutoSize = True
        Me.EepData_Hight_Label.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Hight_Label.Location = New System.Drawing.Point(12, 56)
        Me.EepData_Hight_Label.Name = "EepData_Hight_Label"
        Me.EepData_Hight_Label.Size = New System.Drawing.Size(119, 16)
        Me.EepData_Hight_Label.TabIndex = 144
        Me.EepData_Hight_Label.Text = " TOTAL HIGHT(mm)"
        '
        'EepData_Hight_TextBox
        '
        Me.EepData_Hight_TextBox.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.EepData_Hight_TextBox.Location = New System.Drawing.Point(224, 53)
        Me.EepData_Hight_TextBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.EepData_Hight_TextBox.MaxLength = 999
        Me.EepData_Hight_TextBox.Multiline = True
        Me.EepData_Hight_TextBox.Name = "EepData_Hight_TextBox"
        Me.EepData_Hight_TextBox.Size = New System.Drawing.Size(220, 20)
        Me.EepData_Hight_TextBox.TabIndex = 119
        '
        'FinalCheck_TabPage
        '
        Me.FinalCheck_TabPage.Controls.Add(Me.FinalCheck_Button)
        Me.FinalCheck_TabPage.Location = New System.Drawing.Point(4, 25)
        Me.FinalCheck_TabPage.Name = "FinalCheck_TabPage"
        Me.FinalCheck_TabPage.Size = New System.Drawing.Size(664, 584)
        Me.FinalCheck_TabPage.TabIndex = 11
        Me.FinalCheck_TabPage.Text = "最後檢查"
        Me.FinalCheck_TabPage.UseVisualStyleBackColor = True
        '
        'FinalCheck_Button
        '
        Me.FinalCheck_Button.Font = New System.Drawing.Font("微軟正黑體", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.FinalCheck_Button.Location = New System.Drawing.Point(40, 43)
        Me.FinalCheck_Button.Name = "FinalCheck_Button"
        Me.FinalCheck_Button.Size = New System.Drawing.Size(157, 107)
        Me.FinalCheck_Button.TabIndex = 41
        Me.FinalCheck_Button.Text = "輸出前請點我" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "進行檢查"
        Me.FinalCheck_Button.UseVisualStyleBackColor = True
        '
        'ResultFailOutput_TextBox
        '
        Me.ResultFailOutput_TextBox.Location = New System.Drawing.Point(708, 273)
        Me.ResultFailOutput_TextBox.Multiline = True
        Me.ResultFailOutput_TextBox.Name = "ResultFailOutput_TextBox"
        Me.ResultFailOutput_TextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.ResultFailOutput_TextBox.Size = New System.Drawing.Size(412, 227)
        Me.ResultFailOutput_TextBox.TabIndex = 9
        '
        'JobMaker_Close_Button
        '
        Me.JobMaker_Close_Button.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.JobMaker_Close_Button.ForeColor = System.Drawing.Color.White
        Me.JobMaker_Close_Button.Location = New System.Drawing.Point(660, 8)
        Me.JobMaker_Close_Button.Name = "JobMaker_Close_Button"
        Me.JobMaker_Close_Button.Size = New System.Drawing.Size(23, 23)
        Me.JobMaker_Close_Button.TabIndex = 67
        Me.JobMaker_Close_Button.Text = "X"
        Me.JobMaker_Close_Button.UseVisualStyleBackColor = False
        '
        'EntityCommand1
        '
        Me.EntityCommand1.CommandTimeout = 0
        Me.EntityCommand1.CommandTree = Nothing
        Me.EntityCommand1.Connection = Nothing
        Me.EntityCommand1.EnablePlanCaching = True
        Me.EntityCommand1.Transaction = Nothing
        '
        'JobMaker_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1132, 632)
        Me.ControlBox = False
        Me.Controls.Add(Me.JobMaker_Close_Button)
        Me.Controls.Add(Me.ResultFailOutput_TextBox)
        Me.Controls.Add(Me.ResultCheck_Button)
        Me.Controls.Add(Me.ResultOutput_TextBox)
        Me.Controls.Add(Me.JobMaker_TabControl)
        Me.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "JobMaker_Form"
        Me.Text = "JobMaker_價補妹可▼ω▼"
        Me.G_TabPage.ResumeLayout(False)
        Me.G_TabPage.PerformLayout
        Me.GWeb_GroupBox.ResumeLayout(False)
        Me.GWeb_GroupBox.PerformLayout
        Me.MMIC_TabPage.ResumeLayout(False)
        Me.MMIC_TabPage.PerformLayout
        Me.MMIC_Panel.ResumeLayout(False)
        Me.Panel17.ResumeLayout(False)
        Me.Panel17.PerformLayout
        Me.Panel15.ResumeLayout(False)
        Me.Panel15.PerformLayout
        Me.MMIC_VD10_GroupBox.ResumeLayout(False)
        Me.MMIC_VD10_GroupBox.PerformLayout
        CType(Me.MMIC_VD10_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.MMIC_SV_E_GroupBox.ResumeLayout(False)
        Me.MMIC_SV_E_GroupBox.PerformLayout
        CType(Me.MMIC_SV_E_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.MMIC_SV_GroupBox.ResumeLayout(False)
        Me.MMIC_SV_GroupBox.PerformLayout
        CType(Me.MMIC_SV_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.MMIC_MR_E_GroupBox.ResumeLayout(False)
        Me.MMIC_MR_E_GroupBox.PerformLayout
        CType(Me.MMIC_MR_E_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.MMIC_MR_GroupBox.ResumeLayout(False)
        Me.MMIC_MR_GroupBox.PerformLayout
        CType(Me.MMIC_MR_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.MMIC_GroupBox.ResumeLayout(False)
        Me.MMIC_GroupBox.PerformLayout
        Me.Important_TabPage.ResumeLayout(False)
        Me.Important_TabPage.PerformLayout
        Me.ImpSetting_GroupBox.ResumeLayout(False)
        Me.ImpSetting_GroupBox.PerformLayout
        Me.Spec.ResumeLayout(False)
        Me.Spec_TabControl.ResumeLayout(False)
        Me.Spec_BasicAll_TabPage.ResumeLayout(False)
        Me.Spec_BasicAll_TabPage.PerformLayout
        Me.Spec_BasicAll_TabControl.ResumeLayout(False)
        Me.TabPage7.ResumeLayout(False)
        Me.SpecBasic_GroupBox.ResumeLayout(False)
        Me.SpecBasic_GroupBox.PerformLayout
        CType(Me.Spec_LiftNum_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.SpecBasic_LiftItem_Panel.ResumeLayout(False)
        Me.SpecBasic_LiftItem_Panel.PerformLayout
        Me.TabPage8.ResumeLayout(False)
        Me.SpecBasic_GroupBox2.ResumeLayout(False)
        Me.SpecBasic_GroupBox2.PerformLayout
        CType(Me.Spec_MachineType_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.SpecBasic_p2_base_Panel.ResumeLayout(False)
        CType(Me.Spec_Purpose_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.Spec_TW_TabPage.ResumeLayout(False)
        Me.Spec_TW_TabPage.PerformLayout
        Me.Spec_TW_TabControl.ResumeLayout(False)
        Me.TabPage9.ResumeLayout(False)
        Me.Spec_TW_FlowLayoutPanel1.ResumeLayout(False)
        Me.Spec_DRAuto_Panel.ResumeLayout(False)
        Me.Spec_DRAuto_Panel.PerformLayout
        Me.Spec_CancellCall_Panel.ResumeLayout(False)
        Me.Spec_CancellCall_Panel.PerformLayout
        Me.Spec_AutoFan_Panel.ResumeLayout(False)
        Me.Spec_AutoFan_Panel.PerformLayout
        Me.Spec_AutoPass_Panel.ResumeLayout(False)
        Me.Spec_AutoPass_Panel.PerformLayout
        Me.Spec_Indep_Panel.ResumeLayout(False)
        Me.Spec_Indep_Panel.PerformLayout
        Me.Spec_HinCpi_Panel.ResumeLayout(False)
        Me.Spec_HinCpi_Panel.PerformLayout
        Me.Spec_Fire_Panel.ResumeLayout(False)
        Me.Spec_Fire_Panel.PerformLayout
        Me.Spec_Fireman_Panel.ResumeLayout(False)
        Me.Spec_Fireman_Panel.PerformLayout
        Me.TabPage10.ResumeLayout(False)
        Me.Spec_TW_FlowLayoutPanel2.ResumeLayout(False)
        Me.Spec_Parking_Panel.ResumeLayout(False)
        Me.Spec_Parking_Panel.PerformLayout
        Me.Spec_Seismic_Panel.ResumeLayout(False)
        Me.Spec_Seismic_Panel.PerformLayout
        Me.Spec_CPI_Panel.ResumeLayout(False)
        Me.Spec_CPI_Panel.PerformLayout
        Me.Spec_HallGong_Panel.ResumeLayout(False)
        Me.Spec_HallGong_Panel.PerformLayout
        Me.Spec_HPIMsg_Panel.ResumeLayout(False)
        Me.Spec_HPIMsg_Panel.PerformLayout
        Me.TabPage12.ResumeLayout(False)
        Me.Spec_TW_FlowLayoutPanel3.ResumeLayout(False)
        Me.Spec_CarGong_Panel.ResumeLayout(False)
        Me.Spec_CarGong_Panel.PerformLayout
        Me.Spec_CRD_Panel.ResumeLayout(False)
        Me.Spec_CRD_Panel.PerformLayout
        Me.TabPage13.ResumeLayout(False)
        Me.Spec_TW_FlowLayoutPanel4.ResumeLayout(False)
        Me.Spec_VonicBz_Panel.ResumeLayout(False)
        Me.Spec_VonicBz_Panel.PerformLayout
        Me.Spec_DrHold_Panel.ResumeLayout(False)
        Me.Spec_DrHold_Panel.PerformLayout
        Me.Spec_Landic_Panel.ResumeLayout(False)
        Me.Spec_Landic_Panel.PerformLayout
        Me.Spec_MFLReturn_Panel.ResumeLayout(False)
        Me.Spec_MFLReturn_Panel.PerformLayout
        Me.Spec_Vonic_Panel.ResumeLayout(False)
        Me.Spec_Vonic_Panel.PerformLayout
        Me.Spec_Emer_Panel.ResumeLayout(False)
        Me.Spec_Emer_Panel.PerformLayout
        CType(Me.Spec_EmerNum_NumericUpDown, System.ComponentModel.ISupportInitialize).EndInit
        Me.Spec_emerGroup_TabControl.ResumeLayout(False)
        Me.TabPage14.ResumeLayout(False)
        Me.Spec_TW_FlowLayoutPanel5.ResumeLayout(False)
        Me.Spec_Elvic_Panel.ResumeLayout(False)
        Me.Spec_Elvic_Panel.PerformLayout
        Me.Spec_WCOB_Panel.ResumeLayout(False)
        Me.Spec_WCOB_Panel.PerformLayout
        Me.TabPage15.ResumeLayout(False)
        Me.Spec_TW_FlowLayoutPanel6.ResumeLayout(False)
        Me.Spec_HLL_Panel.ResumeLayout(False)
        Me.Spec_HLL_Panel.PerformLayout
        Me.Spec_ATT_Panel.ResumeLayout(False)
        Me.Spec_ATT_Panel.PerformLayout
        Me.Spec_Flood_Panel.ResumeLayout(False)
        Me.Spec_Flood_Panel.PerformLayout
        Me.Spec_LS1M_Panel.ResumeLayout(False)
        Me.Spec_LS1M_Panel.PerformLayout
        Me.Spec_PRU_Panel.ResumeLayout(False)
        Me.Spec_PRU_Panel.PerformLayout
        Me.Spec_LoadCell_Panel.ResumeLayout(False)
        Me.Spec_LoadCell_Panel.PerformLayout
        Me.Spec_FrontRearDr_Panel.ResumeLayout(False)
        Me.Spec_FrontRearDr_Panel.PerformLayout
        Me.Spec_OpeSw_Panel.ResumeLayout(False)
        Me.Spec_OpeSw_Panel.PerformLayout
        Me.TabPage11.ResumeLayout(False)
        Me.Spec_TW_unUse_FlowLayoutPanel.ResumeLayout(False)
        Me.Panel42.ResumeLayout(False)
        Me.Panel42.PerformLayout
        Me.Panel43.ResumeLayout(False)
        Me.Panel43.PerformLayout
        Me.Panel54.ResumeLayout(False)
        Me.Panel54.PerformLayout
        Me.Panel66.ResumeLayout(False)
        Me.Panel66.PerformLayout
        Me.Spec_WTB_Panel.ResumeLayout(False)
        Me.Spec_WTB_Panel.PerformLayout
        Me.Spec_IF79x_Panel.ResumeLayout(False)
        Me.Spec_IF79x_Panel.PerformLayout
        Me.Spec_EachStop_Panel.ResumeLayout(False)
        Me.Spec_EachStop_Panel.PerformLayout
        Me.Panel115.ResumeLayout(False)
        Me.Panel115.PerformLayout
        Me.Spec_Operation_Panel.ResumeLayout(False)
        Me.Spec_Operation_Panel.PerformLayout
        Me.DWG_TabPage.ResumeLayout(False)
        Me.DWG_TabPage.PerformLayout
        Me.DWG_GroupBox.ResumeLayout(False)
        Me.DWG_GroupBox.PerformLayout
        Me.ProgramChange_TabPage.ResumeLayout(False)
        Me.ProgramChange_TabPage.PerformLayout
        Me.TabControl3.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        Me.ProgramChange_FlowLayoutPanel.ResumeLayout(False)
        Me.use_ProgramChg_Panel1.ResumeLayout(False)
        Me.use_ProgramChg_Panel1.PerformLayout
        Me.use_ProgramChg_Panel2.ResumeLayout(False)
        Me.use_ProgramChg_Panel2.PerformLayout
        Me.use_ProgramChg_Panel3.ResumeLayout(False)
        Me.use_ProgramChg_Panel3.PerformLayout
        Me.use_ProgramChg_Panel5.ResumeLayout(False)
        Me.use_ProgramChg_Panel5.PerformLayout
        Me.TabPage6.ResumeLayout(False)
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.use_ProgramChg_Panel4.ResumeLayout(False)
        Me.use_ProgramChg_Panel4.PerformLayout
        Me.Panel11.ResumeLayout(False)
        Me.Panel11.PerformLayout
        Me.Panel7.ResumeLayout(False)
        Me.Panel7.PerformLayout
        Me.Panel12.ResumeLayout(False)
        Me.Panel12.PerformLayout
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout
        Me.Panel13.ResumeLayout(False)
        Me.Panel13.PerformLayout
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout
        Me.Panel14.ResumeLayout(False)
        Me.Panel14.PerformLayout
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout
        Me.Panel9.ResumeLayout(False)
        Me.Panel9.PerformLayout
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout
        Me.Panel10.ResumeLayout(False)
        Me.Panel10.PerformLayout
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout
        Me.CheckList.ResumeLayout(False)
        Me.CheckList.PerformLayout
        Me.CheckList_GroupBox.ResumeLayout(False)
        Me.CheckList_GroupBox.PerformLayout
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.CheckList_FlowLayoutPanel.ResumeLayout(False)
        Me.ChkList_1_Panel.ResumeLayout(False)
        Me.ChkList_1_Panel.PerformLayout
        Me.ChkList_2_Panel.ResumeLayout(False)
        Me.ChkList_2_Panel.PerformLayout
        Me.ChkList_3_Panel.ResumeLayout(False)
        Me.ChkList_3_Panel.PerformLayout
        Me.TabPage3.ResumeLayout(False)
        Me.CheckList2_FlowLayoutPanel.ResumeLayout(False)
        Me.ChkList_6_Panel.ResumeLayout(False)
        Me.ChkList_6_Panel.PerformLayout
        Me.Panel24.ResumeLayout(False)
        Me.Panel24.PerformLayout
        Me.ChkList_4_Panel.ResumeLayout(False)
        Me.ChkList_4_Panel.PerformLayout
        Me.ChkList_5_Panel.ResumeLayout(False)
        Me.ChkList_5_Panel.PerformLayout
        Me.TabPage4.ResumeLayout(False)
        Me.CheckList3_FlowLayoutPanel.ResumeLayout(False)
        Me.ChkList_7_Panel.ResumeLayout(False)
        Me.ChkList_7_Panel.PerformLayout
        Me.ChkList_8_Panel.ResumeLayout(False)
        Me.ChkList_8_Panel.PerformLayout
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout
        Me.ChkList_9_Panel.ResumeLayout(False)
        Me.ChkList_9_Panel.PerformLayout
        Me.Basic_TabPage.ResumeLayout(False)
        Me.Basic_TabPage.PerformLayout
        Me.Basic_GroupBox.ResumeLayout(False)
        Me.Basic_GroupBox.PerformLayout
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit
        Me.Load_TabPage.ResumeLayout(False)
        Me.Load_Other_btn_GroupBox.ResumeLayout(False)
        Me.Load_SpecDWG_btn_GroupBox.ResumeLayout(False)
        Me.Load_TabControl.ResumeLayout(False)
        Me.AutoLoad_TabPage.ResumeLayout(False)
        Me.AutoLoad_TabPage.PerformLayout
        Me.Load_AutoLoad_GroupBox.ResumeLayout(False)
        Me.Load_AutoLoad_GroupBox.PerformLayout
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit
        Me.Spec_TabPage.ResumeLayout(False)
        Me.Spec_TabPage.PerformLayout
        Me.Load_Spec_GroupBox.ResumeLayout(False)
        Me.Load_Spec_GroupBox.PerformLayout
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit
        Me.CheckList_TabPage.ResumeLayout(False)
        Me.CheckList_TabPage.PerformLayout
        Me.Load_ChkList_GroupBox.ResumeLayout(False)
        Me.Load_ChkList_GroupBox.PerformLayout
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit
        Me.LoadSQL_TabPage.ResumeLayout(False)
        Me.LoadSQL_TabPage.PerformLayout
        Me.Load_SQLite_GroupBox.ResumeLayout(False)
        Me.Load_SQLite_GroupBox.PerformLayout
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit
        Me.JobMaker_TabControl.ResumeLayout(False)
        Me.EepData_TabPage.ResumeLayout(False)
        Me.EepData_TabPage.PerformLayout
        Me.EepData_TabControl.ResumeLayout(False)
        Me.EepData_TabPage1.ResumeLayout(False)
        Me.EepData_Page1_GroupBox.ResumeLayout(False)
        Me.EepData_Page1_GroupBox.PerformLayout
        Me.EepData_TabPage2.ResumeLayout(False)
        Me.EepData_Page2_GroupBox.ResumeLayout(False)
        Me.EepData_Page2_GroupBox.PerformLayout
        Me.EepData_TabPage3.ResumeLayout(False)
        Me.EepData_Page3_GroupBox.ResumeLayout(False)
        Me.EepData_Page3_GroupBox.PerformLayout
        Me.EepData_TabPage4.ResumeLayout(False)
        Me.EepData_Page4_GroupBox.ResumeLayout(False)
        Me.EepData_Page4_GroupBox.PerformLayout
        Me.EepData_TabPage5.ResumeLayout(False)
        Me.EepData_Page5_GroupBox.ResumeLayout(False)
        Me.EepData_Page5_GroupBox.PerformLayout
        Me.EepData_TabPage6.ResumeLayout(False)
        Me.EepData_Page6_GroupBox.ResumeLayout(False)
        Me.EepData_Page6_GroupBox.PerformLayout
        Me.FinalCheck_TabPage.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout

    End Sub
    Friend WithEvents ResultCheck_Button As Button
    Friend WithEvents ResultOutput_TextBox As TextBox
    Friend WithEvents JobMaker_Timer As Timer
    Friend WithEvents G_TabPage As TabPage
    Friend WithEvents Use_G_CheckBox As CheckBox
    Friend WithEvents Label86 As Label
    Friend WithEvents GWeb_Button As Button
    Friend WithEvents MMIC_TabPage As TabPage
    Friend WithEvents Use_mmic_CheckBox As CheckBox
    Friend WithEvents MMIC_MachineType_ComboBox As ComboBox
    Friend WithEvents Label111 As Label
    Friend WithEvents MMIC_FLEX_N_ComboBox As ComboBox
    Friend WithEvents Label112 As Label
    Friend WithEvents Important_TabPage As TabPage
    Friend WithEvents Use_Imp_CheckBox As CheckBox
    Friend WithEvents Imp_DoorType_TextBox As TextBox
    Friend WithEvents Label127 As Label
    Friend WithEvents Imp_OverBalance_ComboBox As ComboBox
    Friend WithEvents Imp_WHB_ComboBox As ComboBox
    Friend WithEvents Label93 As Label
    Friend WithEvents Label61 As Label
    Friend WithEvents Imp_MachineRoom_ComboBox As ComboBox
    Friend WithEvents Label94 As Label
    Friend WithEvents Label96 As Label
    Friend WithEvents Imp_FAN_ComboBox As ComboBox
    Friend WithEvents Label97 As Label
    Friend WithEvents Spec As TabPage
    Friend WithEvents Spec_TabControl As TabControl
    Friend WithEvents Spec_BasicAll_TabPage As TabPage
    Friend WithEvents Use_SpecBasic_CheckBox As CheckBox
    Friend WithEvents SpecBasic_LiftItem_Panel As Panel
    Friend WithEvents Spec_FLName_TextBox As TextBox
    Friend WithEvents Spec_Speed_TextBox As TextBox
    Friend WithEvents Spec_StopFL_TextBox As TextBox
    Friend WithEvents Spec_LiftName_TextBox As TextBox
    Friend WithEvents Spec_BtmFL_TextBox As TextBox
    Friend WithEvents Spec_LiftMem_TextBox As TextBox
    Friend WithEvents Spec_TopFL_TextBox As TextBox
    Friend WithEvents SpecBasic_LiftItem_Dynamic_Panel As Panel
    Friend WithEvents Label8 As Label
    Friend WithEvents Spec_TW_TabPage As TabPage
    Friend WithEvents Use_SpecTWFP17_CheckBox As CheckBox
    Friend WithEvents Use_SpecTWIDU_CheckBox As CheckBox
    Friend WithEvents Spec_TW_FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents Spec_DRAuto_Panel As Panel
    Friend WithEvents Spec_DRAuto_Label As Label
    Friend WithEvents Spec_MechSafety_Label As Label
    Friend WithEvents Spec_MechSafety_ComboBox As ComboBox
    Friend WithEvents Spec_PhotoEye_Label As Label
    Friend WithEvents Spec_PhotoEye_ComboBox As ComboBox
    Friend WithEvents Spec_DRAuto_ComboBox As ComboBox
    Friend WithEvents Spec_CancellCall_Panel As Panel
    Friend WithEvents Spec_CancellCall_Label As Label
    Friend WithEvents Spec_CancellCall_ComboBox As ComboBox
    Friend WithEvents Spec_SCOB_Label As Label
    Friend WithEvents Spec_SCOB_ComboBox As ComboBox
    Friend WithEvents Panel42 As Panel
    Friend WithEvents Label155 As Label
    Friend WithEvents Spec_CancellBehind_ComboBox As ComboBox
    Friend WithEvents Panel43 As Panel
    Friend WithEvents Label156 As Label
    Friend WithEvents Spec_LampChk_ComboBox As ComboBox
    Friend WithEvents Spec_AutoFan_Panel As Panel
    Friend WithEvents Spec_AutoFan_Label As Label
    Friend WithEvents Spec_AutoFan_ComboBox As ComboBox
    Friend WithEvents Spec_ION_Label As Label
    Friend WithEvents Spec_ION_ComboBox As ComboBox
    Friend WithEvents Panel54 As Panel
    Friend WithEvents Label163 As Label
    Friend WithEvents Spec_CCCancell_ComboBox As ComboBox
    Friend WithEvents Spec_AutoPass_Panel As Panel
    Friend WithEvents Spec_AutoPass_Label As Label
    Friend WithEvents Spec_AutoPass_ComboBox As ComboBox
    Friend WithEvents Spec_Indep_Panel As Panel
    Friend WithEvents Spec_Indep_Label As Label
    Friend WithEvents Spec_Indep_ComboBox As ComboBox
    Friend WithEvents Panel66 As Panel
    Friend WithEvents Label169 As Label
    Friend WithEvents Spec_UCMP_ComboBox As ComboBox
    Friend WithEvents Spec_HinCpi_Panel As Panel
    Friend WithEvents Spec_HinCpi_Label As Label
    Friend WithEvents Spec_HinCpi_ComboBox As ComboBox
    Friend WithEvents Spec_Fire_Panel As Panel
    Friend WithEvents Spec_FireSignal_ComboBox As ComboBox
    Friend WithEvents Spec_FireSignal_Label As Label
    Friend WithEvents Spec_Fire_Label As Label
    Friend WithEvents Spec_Fire_ComboBox As ComboBox
    Friend WithEvents Spec_Fireman_Panel As Panel
    Friend WithEvents Label55 As Label
    Friend WithEvents Spec_Fireman_Only_TextBox As TextBox
    Friend WithEvents Spec_EscapeFL_TextBox As TextBox
    Friend WithEvents Spec_EscapeFL_Label As Label
    Friend WithEvents Spec_Fireman_Label As Label
    Friend WithEvents Spec_Fireman_ComboBox As ComboBox
    Friend WithEvents Spec_Parking_Panel As Panel
    Friend WithEvents Spec_ParkingFL_DR_ComboBox As ComboBox
    Friend WithEvents Spec_ParkingFL_DR_Label As Label
    Friend WithEvents Spec_ParkingFL_HALL_ComboBox As ComboBox
    Friend WithEvents Spec_ParkingFL_HALL_Label As Label
    Friend WithEvents Spec_ParkingFL_COB_ComboBox As ComboBox
    Friend WithEvents Spec_ParkingFL_COB_Label As Label
    Friend WithEvents Spec_ParkingFL_WTB_ComboBox As ComboBox
    Friend WithEvents Spec_ParkingFL_WTB_Label As Label
    Friend WithEvents Spec_ParkingFL_ELVIC_Label As Label
    Friend WithEvents Spec_ParkingFL_ELVIC_ComboBox As ComboBox
    Friend WithEvents Spec_Parking_FL_TextBox As TextBox
    Friend WithEvents Spec_Parking_FL_Label As Label
    Friend WithEvents Spec_Parking_Label As Label
    Friend WithEvents Spec_Parking_ComboBox As ComboBox
    Friend WithEvents Spec_Seismic_Panel As Panel
    Friend WithEvents Spec_SeismicSensor_Label As Label
    Friend WithEvents Spec_SeismicSW_Label As Label
    Friend WithEvents Spec_SeismicSW_ComboBox As ComboBox
    Friend WithEvents Spec_Seismic_Label As Label
    Friend WithEvents Spec_Seismic_ComboBox As ComboBox
    Friend WithEvents Spec_CPI_Panel As Panel
    Friend WithEvents Spec_CpiOLT_ComboBox As ComboBox
    Friend WithEvents Spec_CpiOLT_Label As Label
    Friend WithEvents Spec_CpiFM_ComboBox As ComboBox
    Friend WithEvents Spec_CpiFM_Label As Label
    Friend WithEvents Spec_CpiEmer_ComboBox As ComboBox
    Friend WithEvents Spec_CpiEmer_Label As Label
    Friend WithEvents Spec_CpiFire_ComboBox As ComboBox
    Friend WithEvents Spec_CpiFire_Label As Label
    Friend WithEvents Spec_CpiSeismic_ComboBox As ComboBox
    Friend WithEvents Spec_CpiSeismic_Label As Label
    Friend WithEvents Spec_CPI_Label As Label
    Friend WithEvents Spec_CPI_ComboBox As ComboBox
    Friend WithEvents Spec_CarGong_Panel As Panel
    Friend WithEvents Spec_CarGong_Label As Label
    Friend WithEvents Spec_CarGong_ComboBox As ComboBox
    Friend WithEvents Spec_HallGong_Panel As Panel
    Friend WithEvents Spec_HallGong_Label As Label
    Friend WithEvents Spec_HallGong_ComboBox As ComboBox
    Friend WithEvents Spec_HPIMsg_Panel As Panel
    Friend WithEvents Spec_HpiFM_ComboBox As ComboBox
    Friend WithEvents Spec_HpiFM_Label As Label
    Friend WithEvents Spec_HpiIndep_ComboBox As ComboBox
    Friend WithEvents Spec_HpiIndep_Label As Label
    Friend WithEvents Spec_HpiMain_ComboBox As ComboBox
    Friend WithEvents Spec_HpiMain_Label As Label
    Friend WithEvents Spec_HpiOLT_ComboBox As ComboBox
    Friend WithEvents Spec_HpiOLT_Label As Label
    Friend WithEvents Spec_HPIMsg_Label As Label
    Friend WithEvents Spec_HPIMsg_ComboBox As ComboBox
    Friend WithEvents Spec_DrHold_Panel As Panel
    Friend WithEvents Spec_DrHold_Label As Label
    Friend WithEvents Spec_DrHold_ComboBox As ComboBox
    Friend WithEvents Spec_CRD_Panel As Panel
    Friend WithEvents Spec_CRDID5_Label As Label
    Friend WithEvents Spec_CRD_Label As Label
    Friend WithEvents Spec_CRD_ComboBox As ComboBox
    Friend WithEvents Spec_CRDSpec_Label As Label
    Friend WithEvents Spec_CRDSpec_ComboBox As ComboBox
    Friend WithEvents Spec_CRDCancell_Label As Label
    Friend WithEvents Spec_CRDCancell_ComboBox As ComboBox
    Friend WithEvents Spec_CRDNuisance_Label As Label
    Friend WithEvents Spec_CRDNuisance_ComboBox As ComboBox
    Friend WithEvents Spec_CRDReg_Label As Label
    Friend WithEvents Spec_CRDReg_ComboBox As ComboBox
    Friend WithEvents Spec_CRDID4_Label As Label
    Friend WithEvents Spec_CRDID4_ComboBox As ComboBox
    Friend WithEvents Spec_CRDID5_ComboBox As ComboBox
    Friend WithEvents Spec_Landic_Panel As Panel
    Friend WithEvents Spec_Landic_Label As Label
    Friend WithEvents Spec_Landic_ComboBox As ComboBox
    Friend WithEvents Spec_MFLReturn_Panel As Panel
    Friend WithEvents Spec_MFLReturn_Label As Label
    Friend WithEvents Spec_MFLReturn_FL_TextBox As TextBox
    Friend WithEvents Spec_MFLReturn_ComboBox As ComboBox
    Friend WithEvents Spec_MFLReturn_FL_Label As Label
    Friend WithEvents Spec_Vonic_Panel As Panel
    Friend WithEvents Spec_Vonic_standard_Label As Label
    Friend WithEvents Spec_Vonic_standard_ComboBox As ComboBox
    Friend WithEvents Spec_Vonic_Label As Label
    Friend WithEvents Spec_Vonic_ComboBox As ComboBox
    Friend WithEvents Spec_Elvic_Panel As Panel
    Friend WithEvents Spec_Elvic_Label As Label
    Friend WithEvents Spec_Elvic_ComboBox As ComboBox
    Friend WithEvents Spec_Elvic_Parking_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_VIP_CheckBox As CheckBox
    Friend WithEvents Label202 As Label
    Friend WithEvents Spec_Elvic_Indep_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_FloorLockOut_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_Express_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_ReturnFL_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_Traffic_Peak_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_MainFL_CheckBox As CheckBox
    Friend WithEvents Label203 As Label
    Friend WithEvents Spec_Elvic_FloorLockOut_GR_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_Zoning_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_CarCall_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_Traffic_Peak_ComboBox As ComboBox
    Friend WithEvents Spec_Elvic_Fire_CheckBox As CheckBox
    Friend WithEvents Spec_Elvic_Wavic_CheckBox As CheckBox
    Friend WithEvents Label204 As Label
    Friend WithEvents Spec_Elvic_CRD_CheckBox As CheckBox
    Friend WithEvents Spec_HLL_Panel As Panel
    Friend WithEvents Spec_HLL_Label As Label
    Friend WithEvents Spec_HLL_ComboBox As ComboBox
    Friend WithEvents Spec_WCOB_Panel As Panel
    Friend WithEvents Spec_WCOB_only_TextBox As TextBox
    Friend WithEvents Label123 As Label
    Friend WithEvents Spec_WCOB_Label As Label
    Friend WithEvents Spec_WCOB_ComboBox As ComboBox
    Friend WithEvents Spec_WSCOB_Label As Label
    Friend WithEvents Spec_WSCOB_ComboBox As ComboBox
    Friend WithEvents Spec_WCOB_Ring_Label As Label
    Friend WithEvents Spec_WCOB_Ring_ComboBox As ComboBox
    Friend WithEvents Spec_ATT_Panel As Panel
    Friend WithEvents Spec_ATT_Label As Label
    Friend WithEvents Spec_ATT_ComboBox As ComboBox
    Friend WithEvents Spec_Flood_Panel As Panel
    Friend WithEvents Spec_Flood_Label As Label
    Friend WithEvents Spec_Flood_ComboBox As ComboBox
    Friend WithEvents Spec_Flood_FL_TextBox As TextBox
    Friend WithEvents Spec_Flood_FL_Label As Label
    Friend WithEvents Spec_LS1M_Panel As Panel
    Friend WithEvents Spec_LS1M_Label As Label
    Friend WithEvents Spec_LS1M_ComboBox As ComboBox
    Friend WithEvents Spec_PRU_Panel As Panel
    Friend WithEvents Spec_PRU_Label As Label
    Friend WithEvents Spec_PRU_ComboBox As ComboBox
    Friend WithEvents Spec_LoadCell_Panel As Panel
    Friend WithEvents Spec_LoadCellPos_ComboBox As ComboBox
    Friend WithEvents Spec_LoadCellPos_Label As Label
    Friend WithEvents Spec_LoadCell_Label As Label
    Friend WithEvents Spec_LoadCell_ComboBox As ComboBox
    Friend WithEvents DWG_TabPage As TabPage
    Friend WithEvents Label58 As Label
    Friend WithEvents DWG_PageNum_TextBox As TextBox
    Friend WithEvents Label60 As Label
    Friend WithEvents DWG_PrkName_ComboBox As ComboBox
    Friend WithEvents Label59 As Label
    Friend WithEvents DWG_Page_AddButton As Button
    Friend WithEvents DWG_Page_SubButton As Button
    Friend WithEvents DWG_Page_unChkAllButton As Button
    Friend WithEvents DWG_Page_CheckedListBox As CheckedListBox
    Friend WithEvents DWG_Page_ChkAllButton As Button
    Friend WithEvents Use_prk_CheckBox As CheckBox
    Friend WithEvents ProgramChange_TabPage As TabPage
    Friend WithEvents Use_Program_CheckBox As CheckBox
    Friend WithEvents use_ProgramChg_Panel3 As Panel
    Friend WithEvents PrmList_3_debug_CheckBox As CheckBox
    Friend WithEvents PrmList_3_excute_CheckBox As CheckBox
    Friend WithEvents PrmList_3_confirm_CheckBox As CheckBox
    Friend WithEvents PrmList_3_other_Checkbox As CheckBox
    Friend WithEvents PrmList_3_test_CheckBox As CheckBox
    Friend WithEvents PrmList_3_other_TextBox As TextBox
    Friend WithEvents use_ProgramChg_Panel2 As Panel
    Friend WithEvents PrmList_2_Other_CheckBox As CheckBox
    Friend WithEvents PrmList_2_Tower_CheckBox As CheckBox
    Friend WithEvents PrmList_2_COP_CheckBox As CheckBox
    Friend WithEvents PrmList_2_test_CheckBox As CheckBox
    Friend WithEvents PrmList_2_test_TextBox As TextBox
    Friend WithEvents PrmList_2_COP_TextBox As TextBox
    Friend WithEvents PrmList_2_tower_TextBox As TextBox
    Friend WithEvents PrmList_2_other_TextBox As TextBox
    Friend WithEvents Label52 As Label
    Friend WithEvents PrmList_5_review_CheckBox As CheckBox
    Friend WithEvents Label35 As Label
    Friend WithEvents Label33 As Label
    Friend WithEvents PrmList_1_reason_TextBox As TextBox
    Friend WithEvents Label32 As Label
    Friend WithEvents Label34 As Label
    Friend WithEvents CheckList As TabPage
    Friend WithEvents CheckList_GroupBox As GroupBox
    Friend WithEvents Label10 As Label
    Friend WithEvents ChkList_PaSheet_CheckBox As CheckBox
    Friend WithEvents Label11 As Label
    Friend WithEvents ChkList_OS_CheckBox As CheckBox
    Friend WithEvents ChkList_Elec_DateTimePicker As DateTimePicker
    Friend WithEvents Label12 As Label
    Friend WithEvents ChkList_Confirm_DateTimePicker As DateTimePicker
    Friend WithEvents ChkList_Confirm_CheckBox As CheckBox
    Friend WithEvents ChkList_OS_DateTimePicker As DateTimePicker
    Friend WithEvents Label13 As Label
    Friend WithEvents ChkList_PaSheet_DateTimePicker As DateTimePicker
    Friend WithEvents ChkList_Elec_CheckBox As CheckBox
    Friend WithEvents Use_ChkList_CheckBox As CheckBox
    Friend WithEvents Button9 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents ChkList_3_Panel As Panel
    Friend WithEvents ChkList_3_yes_RadioButton As RadioButton
    Friend WithEvents ChkList_3_no_RadioButton As RadioButton
    Friend WithEvents Label19 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents ChkList_3_yes_Man_TextBox As TextBox
    Friend WithEvents Label20 As Label
    Friend WithEvents ChkList_3_yes_Content_TextBox As TextBox
    Friend WithEvents Label21 As Label
    Friend WithEvents ChkList_3_yes_Result_TextBox As TextBox
    Friend WithEvents ChkList_2_Panel As Panel
    Friend WithEvents ChkList_2_yes_RadioButton As RadioButton
    Friend WithEvents ChkList_2_no_RadioButton As RadioButton
    Friend WithEvents Label17 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents ChkList_2_yes_Result_TextBox As TextBox
    Friend WithEvents ChkList_2_yes_Content_TextBox As TextBox
    Friend WithEvents ChkList_1_Panel As Panel
    Friend WithEvents ChkList_1_no_RadioButton As RadioButton
    Friend WithEvents ChkList_1_yes_RadioButton As RadioButton
    Friend WithEvents ChkList_1_yes_Content_TextBox As TextBox
    Friend WithEvents ChkList_1_yes_result_TextBox As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Basic_TabPage As TabPage
    Friend WithEvents Use_Basic_CheckBox As CheckBox
    Friend WithEvents NumericUpDown1 As NumericUpDown
    Friend WithEvents ReminderMarquee_Label As Label
    Friend WithEvents Basic_JobNoMOD_TextBox As TextBox
    Friend WithEvents Basic_JobNoNew_TextBox As TextBox
    Friend WithEvents Basic_JobNoOld_TextBox As TextBox
    Friend WithEvents Basic_JobName_TextBox As TextBox
    Friend WithEvents Basic_JobNoMOD_Label As Label
    Friend WithEvents Basic_DrawDate_DateTimePicker As DateTimePicker
    Friend WithEvents Label53 As Label
    Friend WithEvents Basic_ApproverChinese_ComboBox As ComboBox
    Friend WithEvents Basic_Local_ComboBox As ComboBox
    Friend WithEvents Basic_JobNoNew_Label As Label
    Friend WithEvents Basic_CheckerChinese_ComboBox As ComboBox
    Friend WithEvents Basic_DesingerChinese_ComboBox As ComboBox
    Friend WithEvents Basic_Local_Label As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Basic_JobNoOld_Label As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Load_TabPage As TabPage
    Friend WithEvents Load_TabControl As TabControl
    Friend WithEvents Spec_TabPage As TabPage
    Friend WithEvents Load_Spec_GroupBox As GroupBox
    Friend WithEvents JobMaker_LOAD_Spec_CheckBox As CheckBox
    Friend WithEvents CheckList_TabPage As TabPage
    Friend WithEvents Load_ChkList_GroupBox As GroupBox
    Friend WithEvents JobMaker_LOAD_ChkList_CheckBox As CheckBox
    Friend WithEvents JobMaker_TabControl As TabControl
    Friend WithEvents JMFileCho_Spec_TextBox As TextBox
    Friend WithEvents JMFileCho_ChkList_TextBox As TextBox
    Friend WithEvents Basic_ApproverEnglish_ComboBox As ComboBox
    Friend WithEvents Basic_CheckerEnglish_ComboBox As ComboBox
    Friend WithEvents Basic_DesingerEnglish_ComboBox As ComboBox
    Friend WithEvents Label222 As Label
    Friend WithEvents Label221 As Label
    Friend WithEvents Label220 As Label
    Friend WithEvents Label219 As Label
    Friend WithEvents Label218 As Label
    Friend WithEvents Label217 As Label
    Friend WithEvents ImpSetting_GroupBox As GroupBox
    Friend WithEvents HallIndicator_FlowLayoutPanel As FlowLayoutPanel
    Friend WithEvents Spec_VonicBz_Panel As Panel
    Friend WithEvents Spec_VonicBz_Label As Label
    Friend WithEvents Spec_VonicBz_ComboBox As ComboBox
    Friend WithEvents Spec_Fireman_Only_CheckBox As CheckBox
    Friend WithEvents Spec_SeismicSensor_ComboBox As ComboBox
    Friend WithEvents Load_Other_btn_GroupBox As GroupBox
    Friend WithEvents CheckList_OutputButton As Button
    Friend WithEvents DWG_OutputButton As Button
    Friend WithEvents Spec_OutputButton As Button
    Friend WithEvents Load_SpecDWG_btn_GroupBox As GroupBox
    Friend WithEvents All_OutputButton As Button
    Friend WithEvents DWG_StdPage_Button As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents PictureBox3 As PictureBox
    Friend WithEvents CheckList_FlowLayoutPanel As FlowLayoutPanel
    Friend WithEvents ProgramChange_FlowLayoutPanel As FlowLayoutPanel
    Friend WithEvents use_ProgramChg_Panel1 As Panel
    Friend WithEvents use_ProgramChg_Panel5 As Panel
    Friend WithEvents ResultFailOutput_TextBox As TextBox
    Friend WithEvents HIN_TestButton As Button
    Friend WithEvents JMFileCho_Spec_Button As Button
    Friend WithEvents JMFileCho_ChkList_Button As Button
    Friend WithEvents LoadSQL_TabPage As TabPage
    Friend WithEvents Load_SQLite_GroupBox As GroupBox
    Friend WithEvents JMFileCho_SQLite_Button As Button
    Friend WithEvents PictureBox4 As PictureBox
    Friend WithEvents JMFileCho_SQLite_TextBox As TextBox
    Friend WithEvents JobMaker_LOAD_SQLite_CheckBox As CheckBox
    Friend WithEvents Label149 As Label
    Friend WithEvents JM_DefaultPath_Spec_Label As Label
    Friend WithEvents JM_DefaultPath_CheckList_Label As Label
    Friend WithEvents Label173 As Label
    Friend WithEvents JM_DefaultPath_SQLite_Label As Label
    Friend WithEvents Label188 As Label
    Friend WithEvents JMFileConfirm_SQLite_Button As Button
    Friend WithEvents JobMaker_Close_Button As Button
    Friend WithEvents Spec_Control_ComboBox As ComboBox
    Friend WithEvents Spec_LiftNum_NumericUpDown As NumericUpDown
    Friend WithEvents Basic_GroupBox As GroupBox
    Friend WithEvents SpecBasic_GroupBox As GroupBox
    Friend WithEvents DWG_GroupBox As GroupBox
    Friend WithEvents MMIC_GroupBox As GroupBox
    Friend WithEvents Spec_BtmFL_Real_TextBox As TextBox
    Friend WithEvents Spec_TopFL_Real_TextBox As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents GWeb_GroupBox As GroupBox
    Friend WithEvents DWG_Construction_CheckedListBox As CheckedListBox
    Friend WithEvents DWG_Produce_CheckedListBox As CheckedListBox
    Friend WithEvents Label193 As Label
    Friend WithEvents Label192 As Label
    Friend WithEvents Label194 As Label
    Friend WithEvents DWG_VonicStd_ComboBox As ComboBox
    Friend WithEvents Spec_Fire_Only_CheckBox As CheckBox
    Friend WithEvents Label195 As Label
    Friend WithEvents Spec_Fire_Only_TextBox As TextBox
    Friend WithEvents Spec_Parking_Only_CheckBox As CheckBox
    Friend WithEvents Label56 As Label
    Friend WithEvents Spec_Parking_Only_TextBox As TextBox
    Friend WithEvents Spec_SeismicSW_Only_CheckBox As CheckBox
    Friend WithEvents Label215 As Label
    Friend WithEvents Spec_SeismicSW_Only_TextBox As TextBox
    Friend WithEvents Spec_SeismicSensor_Only_CheckBox As CheckBox
    Friend WithEvents Label214 As Label
    Friend WithEvents Spec_SeismicSensor_Only_TextBox As TextBox
    Friend WithEvents Spec_Seismic_Only_CheckBox As CheckBox
    Friend WithEvents Label196 As Label
    Friend WithEvents Spec_Seismic_Only_TextBox As TextBox
    Friend WithEvents Spec_CpiOLT_Only_CheckBox As CheckBox
    Friend WithEvents Label216 As Label
    Friend WithEvents Spec_CpiOLT_Only_TextBox As TextBox
    Friend WithEvents Spec_CarGong_VONIC_Only_CheckBox As CheckBox
    Friend WithEvents Label225 As Label
    Friend WithEvents Spec_CarGong_VONIC_Only_TextBox As TextBox
    Friend WithEvents Spec_CarGong_COB_Only_CheckBox As CheckBox
    Friend WithEvents Label224 As Label
    Friend WithEvents Spec_CarGong_COB_Only_TextBox As TextBox
    Friend WithEvents Spec_CarGong_TopBtm_Only_CheckBox As CheckBox
    Friend WithEvents Label79 As Label
    Friend WithEvents Spec_CarGong_TopBtm_Only_TextBox As TextBox
    Friend WithEvents Spec_CarGong_VONIC_CheckBox As CheckBox
    Friend WithEvents Spec_CarGong_COB_CheckBox As CheckBox
    Friend WithEvents Spec_CarGong_TopBtm_CheckBox As CheckBox
    Friend WithEvents Spec_CarGong_Top_CheckBox As CheckBox
    Friend WithEvents Spec_CarGong_Top_Only_CheckBox As CheckBox
    Friend WithEvents Label223 As Label
    Friend WithEvents Spec_CarGong_Top_Only_TextBox As TextBox
    Friend WithEvents Spec_CarGong_VONIC_TextBox As TextBox
    Friend WithEvents Spec_CarGong_COB_TextBox As TextBox
    Friend WithEvents Spec_CarGong_TopBtm_TextBox As TextBox
    Friend WithEvents Spec_CarGong_Top_TextBox As TextBox
    Friend WithEvents Spec_CRDType_ComboBox As ComboBox
    Friend WithEvents Spec_CRDType_Label As Label
    Friend WithEvents Spec_WSCOB_only_CheckBox As CheckBox
    Friend WithEvents Spec_WSCOB_only_TextBox As TextBox
    Friend WithEvents Label227 As Label
    Friend WithEvents Spec_WCOB_only_CheckBox As CheckBox
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents CheckList2_FlowLayoutPanel As FlowLayoutPanel
    Friend WithEvents ChkList_6_Panel As Panel
    Friend WithEvents Panel24 As Panel
    Friend WithEvents ChkList_6_yesItem_RadioButton As RadioButton
    Friend WithEvents ChkList_6_yesChk_RadioButton As RadioButton
    Friend WithEvents ChkList_6_yes_Content_TextBox As TextBox
    Friend WithEvents ChkList_6_no_RadioButton As RadioButton
    Friend WithEvents ChkList_6_yes_RadioButton As RadioButton
    Friend WithEvents Label27 As Label
    Friend WithEvents ChkList_4_Panel As Panel
    Friend WithEvents ChkList_4_ObjName_TextBox As TextBox
    Friend WithEvents ChkList_4_ObjBase_TextBox As TextBox
    Friend WithEvents Label22 As Label
    Friend WithEvents ChkList_4_SV_TextBox As TextBox
    Friend WithEvents Label23 As Label
    Friend WithEvents Label24 As Label
    Friend WithEvents Label26 As Label
    Friend WithEvents Label25 As Label
    Friend WithEvents ChkList_4_SVBase_TextBox As TextBox
    Friend WithEvents ChkList_5_Panel As Panel
    Friend WithEvents ChkList_5_nstd_RadioButton As RadioButton
    Friend WithEvents ChkList_5_std_RadioButton As RadioButton
    Friend WithEvents ChkList_5_no_RadioButton As RadioButton
    Friend WithEvents Label30 As Label
    Friend WithEvents ChkList_5_std_Content_TextBox As TextBox
    Friend WithEvents ChkList_5_nstd_Content_TextBox As TextBox
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents CheckList3_FlowLayoutPanel As FlowLayoutPanel
    Friend WithEvents ChkList_7_Panel As Panel
    Friend WithEvents ChkList_7_yes_RadioButton As RadioButton
    Friend WithEvents ChkList_7_no_RadioButton As RadioButton
    Friend WithEvents Label28 As Label
    Friend WithEvents ChkList_7_yes1_content_TextBox As TextBox
    Friend WithEvents ChkList_8_Panel As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ChkList_8_yes_RadioButton As RadioButton
    Friend WithEvents ChkList_8_no_RadioButton As RadioButton
    Friend WithEvents ChkList_8Item_RadioButton As RadioButton
    Friend WithEvents Label29 As Label
    Friend WithEvents ChkList_9_Panel As Panel
    Friend WithEvents ChkList_9_no_RadioButton As RadioButton
    Friend WithEvents ChkList_9_yes_RadioButton As RadioButton
    Friend WithEvents Label31 As Label
    Friend WithEvents TabControl3 As TabControl
    Friend WithEvents TabPage5 As TabPage
    Friend WithEvents TabPage6 As TabPage
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents use_ProgramChg_Panel4 As Panel
    Friend WithEvents Label36 As Label
    Friend WithEvents Panel11 As Panel
    Friend WithEvents PrmList_4_yes12_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no12_RadioButton As RadioButton
    Friend WithEvents Label37 As Label
    Friend WithEvents Panel7 As Panel
    Friend WithEvents PrmList_4_yes8_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no8_RadioButton As RadioButton
    Friend WithEvents Label38 As Label
    Friend WithEvents Panel12 As Panel
    Friend WithEvents PrmList_4_yes11_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no11_RadioButton As RadioButton
    Friend WithEvents Label39 As Label
    Friend WithEvents Panel6 As Panel
    Friend WithEvents PrmList_4_yes4_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no4_RadioButton As RadioButton
    Friend WithEvents Label40 As Label
    Friend WithEvents Panel13 As Panel
    Friend WithEvents PrmList_4_yes10_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no10_RadioButton As RadioButton
    Friend WithEvents Label41 As Label
    Friend WithEvents Panel8 As Panel
    Friend WithEvents PrmList_4_yes7_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no7_RadioButton As RadioButton
    Friend WithEvents Label42 As Label
    Friend WithEvents Panel14 As Panel
    Friend WithEvents PrmList_4_yes9_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no9_RadioButton As RadioButton
    Friend WithEvents Label43 As Label
    Friend WithEvents Panel5 As Panel
    Friend WithEvents PrmList_4_yes3_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no3_RadioButton As RadioButton
    Friend WithEvents Label48 As Label
    Friend WithEvents Panel9 As Panel
    Friend WithEvents PrmList_4_yes6_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no6_RadioButton As RadioButton
    Friend WithEvents Label47 As Label
    Friend WithEvents Panel4 As Panel
    Friend WithEvents PrmList_4_yes2_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no2_RadioButton As RadioButton
    Friend WithEvents Label46 As Label
    Friend WithEvents Panel10 As Panel
    Friend WithEvents PrmList_4_yes5_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no5_RadioButton As RadioButton
    Friend WithEvents Label45 As Label
    Friend WithEvents Panel3 As Panel
    Friend WithEvents PrmList_4_yes1_RadioButton As RadioButton
    Friend WithEvents PrmList_4_no1_RadioButton As RadioButton
    Friend WithEvents Label44 As Label
    Friend WithEvents Label51 As Label
    Friend WithEvents PrmList_4_content12_TextBox As TextBox
    Friend WithEvents Label50 As Label
    Friend WithEvents Label49 As Label
    Friend WithEvents Spec_BasicAll_TabControl As TabControl
    Friend WithEvents TabPage7 As TabPage
    Friend WithEvents TabPage8 As TabPage
    Friend WithEvents Label189 As Label
    Friend WithEvents Spec_ControlWay_Label As Label
    Friend WithEvents Spec_MachineType_Label As Label
    Friend WithEvents Spec_Purpose_NumericUpDown As NumericUpDown
    Friend WithEvents Spec_MachineType_NumericUpDown As NumericUpDown
    Friend WithEvents Spec_Purpose_Panel As Panel
    Friend WithEvents Spec_MachineType_Panel As Panel
    Friend WithEvents SpecBasic_p2_base_Panel As Panel
    Friend WithEvents Spec_Base_ComboBox As ComboBox
    Friend WithEvents Spec_ControlWay_Panel As Panel
    Friend WithEvents MMIC_Panel As Panel
    Friend WithEvents Panel17 As Panel
    Friend WithEvents mmicType1_ObjNameBase_TextBox As TextBox
    Friend WithEvents mmicType1_ObjName_TextBox As TextBox
    Friend WithEvents mmicType1_CarNo_TextBox As TextBox
    Friend WithEvents Panel15 As Panel
    Friend WithEvents mmic_ObjName_TextBox As TextBox
    Friend WithEvents mmic_CarNo_TextBox As TextBox
    Friend WithEvents MMIC_VD10_GroupBox As GroupBox
    Friend WithEvents MMIC_VD10_NumericUpDown As NumericUpDown
    Friend WithEvents MMIC_VD10_Base_TextBox As TextBox
    Friend WithEvents MMIC_VD10_Type_ComboBox As ComboBox
    Friend WithEvents Label132 As Label
    Friend WithEvents Label131 As Label
    Friend WithEvents Label114 As Label
    Friend WithEvents Label115 As Label
    Friend WithEvents Label113 As Label
    Friend WithEvents MMIC_VD10_Panel As Panel
    Friend WithEvents Label65 As Label
    Friend WithEvents MMIC_VD10_ROM_ComboBox As ComboBox
    Friend WithEvents MMIC_VD10_Quantity_ComboBox As ComboBox
    Friend WithEvents MMIC_SV_E_GroupBox As GroupBox
    Friend WithEvents MMIC_SV_E_NumericUpDown As NumericUpDown
    Friend WithEvents MMIC_SV_ECarObj_ComboBox As ComboBox
    Friend WithEvents Label106 As Label
    Friend WithEvents MMIC_SV_E_Panel As Panel
    Friend WithEvents Label107 As Label
    Friend WithEvents Label63 As Label
    Friend WithEvents MMIC_SV_EBase_ComboBox As ComboBox
    Friend WithEvents MMIC_SV_GroupBox As GroupBox
    Friend WithEvents Label231 As Label
    Friend WithEvents MMIC_SV_NumericUpDown As NumericUpDown
    Friend WithEvents Label130 As Label
    Friend WithEvents MMIC_SV_Type_ComboBox As ComboBox
    Friend WithEvents MMIC_SV_Base_TextBox As TextBox
    Friend WithEvents Label129 As Label
    Friend WithEvents Label103 As Label
    Friend WithEvents MMIC_SV_Panel As Panel
    Friend WithEvents Label104 As Label
    Friend WithEvents MMIC_MR_E_GroupBox As GroupBox
    Friend WithEvents MMIC_MR_E_NumericUpDown As NumericUpDown
    Friend WithEvents MMIC_MR_ECarObj_ComboBox As ComboBox
    Friend WithEvents Label100 As Label
    Friend WithEvents MMIC_MR_E_Panel As Panel
    Friend WithEvents Label101 As Label
    Friend WithEvents Label62 As Label
    Friend WithEvents MMIC_MR_EBase_ComboBox As ComboBox
    Friend WithEvents MMIC_MR_GroupBox As GroupBox
    Friend WithEvents Label229 As Label
    Friend WithEvents MMIC_MR_NumericUpDown As NumericUpDown
    Friend WithEvents MMIC_MR_Base_TextBox As TextBox
    Friend WithEvents Label64 As Label
    Friend WithEvents Label99 As Label
    Friend WithEvents Label95 As Label
    Friend WithEvents MMIC_MR_CP43x_ComboBox As ComboBox
    Friend WithEvents MMIC_MR_Panel As Panel
    Friend WithEvents Label128 As Label
    Friend WithEvents Spec_TW_TabControl As TabControl
    Friend WithEvents TabPage9 As TabPage
    Friend WithEvents TabPage10 As TabPage
    Friend WithEvents Spec_TW_FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents TabPage12 As TabPage
    Friend WithEvents Spec_TW_FlowLayoutPanel3 As FlowLayoutPanel
    Friend WithEvents TabPage13 As TabPage
    Friend WithEvents Spec_TW_FlowLayoutPanel4 As FlowLayoutPanel
    Friend WithEvents TabPage14 As TabPage
    Friend WithEvents Spec_TW_FlowLayoutPanel5 As FlowLayoutPanel
    Friend WithEvents TabPage15 As TabPage
    Friend WithEvents Spec_TW_FlowLayoutPanel6 As FlowLayoutPanel
    Friend WithEvents Spec_FrontRearDr_Panel As Panel
    Friend WithEvents Spec_FrontRearDr_Label As Label
    Friend WithEvents Spec_FrontRearDr_ComboBox As ComboBox
    Friend WithEvents TabPage11 As TabPage
    Friend WithEvents Spec_TW_unUse_FlowLayoutPanel As FlowLayoutPanel
    Friend WithEvents Spec_WTB_Panel As Panel
    Friend WithEvents Label144 As Label
    Friend WithEvents Spec_WTB_EQMac_ComboBox As ComboBox
    Friend WithEvents Label143 As Label
    Friend WithEvents Spec_WTB_EQIND_ComboBox As ComboBox
    Friend WithEvents Label142 As Label
    Friend WithEvents Spec_WTB_Indep_ComboBox As ComboBox
    Friend WithEvents Label141 As Label
    Friend WithEvents Spec_WTB_EQ_ComboBox As ComboBox
    Friend WithEvents Label140 As Label
    Friend WithEvents Spec_WTB_Alart_ComboBox As ComboBox
    Friend WithEvents Label137 As Label
    Friend WithEvents Spec_WTB_BZSW_ComboBox As ComboBox
    Friend WithEvents Label138 As Label
    Friend WithEvents Spec_WTB_EQSW_ComboBox As ComboBox
    Friend WithEvents Label139 As Label
    Friend WithEvents Spec_WTB_PKSW_ComboBox As ComboBox
    Friend WithEvents Label133 As Label
    Friend WithEvents Spec_WTB_EmerPow_ComboBox As ComboBox
    Friend WithEvents Label134 As Label
    Friend WithEvents Spec_WTB_FO_ComboBox As ComboBox
    Friend WithEvents Label135 As Label
    Friend WithEvents Spec_WTB_Urgent_ComboBox As ComboBox
    Friend WithEvents Label136 As Label
    Friend WithEvents Spec_WTB_Normal_ComboBox As ComboBox
    Friend WithEvents Label108 As Label
    Friend WithEvents Spec_WTB_ChkSW_ComboBox As ComboBox
    Friend WithEvents Label105 As Label
    Friend WithEvents Spec_WTB_FM_ComboBox As ComboBox
    Friend WithEvents Label102 As Label
    Friend WithEvents Spec_WTB_Stop_ComboBox As ComboBox
    Friend WithEvents Label98 As Label
    Friend WithEvents Spec_WTB_Error_ComboBox As ComboBox
    Friend WithEvents Label68 As Label
    Friend WithEvents Spec_WTB_ComboBox As ComboBox
    Friend WithEvents Spec_IF79x_Panel As Panel
    Friend WithEvents Label120 As Label
    Friend WithEvents Spec_IF79x_IDM0_ComboBox As ComboBox
    Friend WithEvents Label121 As Label
    Friend WithEvents Spec_IF79x_ID12_ComboBox As ComboBox
    Friend WithEvents Label119 As Label
    Friend WithEvents Spec_IF79x_ID5_ComboBox As ComboBox
    Friend WithEvents Label118 As Label
    Friend WithEvents Spec_IF79x_ID4_ComboBox As ComboBox
    Friend WithEvents Label117 As Label
    Friend WithEvents Spec_IF79x_ID0_ComboBox As ComboBox
    Friend WithEvents Label69 As Label
    Friend WithEvents Spec_IF79x_ComboBox As ComboBox
    Friend WithEvents Spec_EachStop_Panel As Panel
    Friend WithEvents Label71 As Label
    Friend WithEvents Spec_EachStop_ComboBox As ComboBox
    Friend WithEvents Panel115 As Panel
    Friend WithEvents Label_SPEC_INSTALL_OPE As Label
    Friend WithEvents Spec_install_ope_ComboBox As ComboBox
    Friend WithEvents SpecBasic_GroupBox2 As GroupBox
    Friend WithEvents Spec_Emer_Panel As Panel
    Friend WithEvents Spec_EmerNum_NumericUpDown As NumericUpDown
    Friend WithEvents Spec_EmerCapacity_Label As Label
    Friend WithEvents Spec_EmerSignal_Label As Label
    Friend WithEvents Spec_EmerAddress_ComboBox As ComboBox
    Friend WithEvents Spec_EmerInput_ComboBox As ComboBox
    Friend WithEvents Spec_EmerAddress_Label As Label
    Friend WithEvents Spec_emerGroup_TabControl As TabControl
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Spec_EmerNum_Label As Label
    Friend WithEvents Spec_Emer_Label As Label
    Friend WithEvents Spec_Emer_ComboBox As ComboBox
    Friend WithEvents Spec_EmerInput_Label As Label
    Friend WithEvents Spec_EmerCapacity_TextBox As TextBox
    Friend WithEvents Spec_EmerSignal_ComboBox As ComboBox
    Friend WithEvents Spec_OpeSw_Panel As Panel
    Friend WithEvents Spec_OpeSw_InputPos_ComboBox As ComboBox
    Friend WithEvents Spec_OpeSw_InputAddress_TextBox As TextBox
    Friend WithEvents Spec_OpeSw_InputPos_Label As Label
    Friend WithEvents Spec_OpeSw_DevicePos_TextBox As TextBox
    Friend WithEvents Spec_OpeSw_DevicePos_Label As Label
    Friend WithEvents Spec_OpeSw_Label As Label
    Friend WithEvents Spec_OpeSw_ComboBox As ComboBox
    Friend WithEvents JM_JobSelect_SQLite_TextBox As TextBox
    Friend WithEvents Label109 As Label
    Friend WithEvents JM_JobSelect_SQLite_ComboBox As ComboBox
    Friend WithEvents FinalCheck_TabPage As TabPage
    Friend WithEvents FinalCheck_Button As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents JM_JobSelect_Spec_ComboBox As ComboBox
    Friend WithEvents JM_JobSelect_Spec_TextBox As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents JM_JobSelect_CheckList_ComboBox As ComboBox
    Friend WithEvents JM_JobSelect_CheckList_TextBox As TextBox
    Friend WithEvents Spec_Operation_Panel As Panel
    Friend WithEvents Spec_Operation_Label As Label
    Friend WithEvents Spec_Operation_ComboBox As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents AutoLoad_TabPage As TabPage
    Friend WithEvents Load_AutoLoad_GroupBox As GroupBox
    Friend WithEvents Label54 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label57 As Label
    Friend WithEvents Label66 As Label
    Friend WithEvents JMFileCho_AutoLoad_Button As Button
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents JMFileCho_AutoLoad_TextBox As TextBox
    Friend WithEvents JobMaker_LOAD_AutoLoad_CheckBox As CheckBox
    Friend WithEvents JMFileConfirm_AutoLoad_Button As Button
    Friend WithEvents EepData_TabPage As TabPage
    Friend WithEvents EepData_TabControl As TabControl
    Friend WithEvents EepData_TabPage1 As TabPage
    Friend WithEvents EepData_MachineRoom_Label As Label
    Friend WithEvents EepData_MachineRoom_TextBox As TextBox
    Friend WithEvents EepData_TabPage2 As TabPage
    Friend WithEvents EepData_DrCloser_Label As Label
    Friend WithEvents EepData_DrCloser_TextBox As TextBox
    Friend WithEvents EepData_CarNo_TextBox As TextBox
    Friend WithEvents EepData_CarNo_Label As Label
    Friend WithEvents EepData_GroupNo_TextBox As TextBox
    Friend WithEvents EepData_GroupNo_Label As Label
    Friend WithEvents EepData_Purpose_TextBox As TextBox
    Friend WithEvents EepData_Purpose_Label As Label
    Friend WithEvents EepData_GspType_TextBox As TextBox
    Friend WithEvents EepData_GspType_Label As Label
    Friend WithEvents EepData_OpeType_TextBox As TextBox
    Friend WithEvents EepData_OpeType_Label As Label
    Friend WithEvents EepData_StopFL_TextBox As TextBox
    Friend WithEvents EepData_StopFL_Label As Label
    Friend WithEvents EepData_BtmFL_TextBox As TextBox
    Friend WithEvents EepData_BtmFL_Label As Label
    Friend WithEvents EepData_TopFL_TextBox As TextBox
    Friend WithEvents EepData_TopFL_Label As Label
    Friend WithEvents EepData_Capactity_TextBox As TextBox
    Friend WithEvents EepData_Capactity_Label As Label
    Friend WithEvents EepData_Speed_TextBox As TextBox
    Friend WithEvents EepData_Speed_Label As Label
    Friend WithEvents EepData_DrType_TextBox As TextBox
    Friend WithEvents EepData_DrType_Label As Label
    Friend WithEvents Use_EepData_CheckBox As CheckBox
    Friend WithEvents EepData_DrHold_Label As Label
    Friend WithEvents EepData_AutoFan_Label As Label
    Friend WithEvents EepData_AutoByPass_Label As Label
    Friend WithEvents EepData_Nudging_Label As Label
    Friend WithEvents EepData_Seismic_Label As Label
    Friend WithEvents EepData_MainFL_FR_Label As Label
    Friend WithEvents EepData_MainFL_Label As Label
    Friend WithEvents EepData_SpecMainFL_Label As Label
    Friend WithEvents EepData_Indep_Label As Label
    Friend WithEvents EepData_EnergyRe_Label As Label
    Friend WithEvents EepData_Landic_Label As Label
    Friend WithEvents EepData_DrFrontWidth_Label As Label
    Friend WithEvents EepData_Page1_GroupBox As GroupBox
    Friend WithEvents EepData_Page2_GroupBox As GroupBox
    Friend WithEvents EepData_DrHold_TextBox As TextBox
    Friend WithEvents EepData_DrFrontWidth_TextBox As TextBox
    Friend WithEvents EepData_Landic_TextBox As TextBox
    Friend WithEvents EepData_AutoFan_TextBox As TextBox
    Friend WithEvents EepData_AutoByPass_TextBox As TextBox
    Friend WithEvents EepData_EnergyRe_TextBox As TextBox
    Friend WithEvents EepData_Nudging_TextBox As TextBox
    Friend WithEvents EepData_Indep_TextBox As TextBox
    Friend WithEvents EepData_SpecMainFL_TextBox As TextBox
    Friend WithEvents EepData_Seismic_TextBox As TextBox
    Friend WithEvents EepData_MainFL_TextBox As TextBox
    Friend WithEvents EepData_MainFL_FR_TextBox As TextBox
    Friend WithEvents EepData_TabPage3 As TabPage
    Friend WithEvents EepData_Page3_GroupBox As GroupBox
    Friend WithEvents EepData_Overbalance_TextBox As TextBox
    Friend WithEvents EepData_DrCloseBtn_TextBox As TextBox
    Friend WithEvents EepData_Overbalance_Label As Label
    Friend WithEvents EepData_CarChime_TextBox As TextBox
    Friend WithEvents EepData_EscapeOpe_ForR_TextBox As TextBox
    Friend WithEvents EepData_EscapeOpe_ForR_Label As Label
    Friend WithEvents EepData_EscapeFL_TextBox As TextBox
    Friend WithEvents EepData_HallChime_TextBox As TextBox
    Friend WithEvents EepData_EscapeFL_Label As Label
    Friend WithEvents EepData_EscapeOpe_TextBox As TextBox
    Friend WithEvents EepData_ParkingOpe_TextBox As TextBox
    Friend WithEvents EepData_EscapeOpe_Label As Label
    Friend WithEvents EepData_SafetyShoe_TextBox As TextBox
    Friend WithEvents EepData_ParkingSW_TextBox As TextBox
    Friend WithEvents EepData_SafetyShoe_Label As Label
    Friend WithEvents EepData_PhotoEye_TextBox As TextBox
    Friend WithEvents EepData_ParkingFL_TextBox As TextBox
    Friend WithEvents EepData_PhotoEye_Label As Label
    Friend WithEvents EepData_ParkingFL_ForR_TextBox As TextBox
    Friend WithEvents EepData_DrCloseBtn_Label As Label
    Friend WithEvents EepData_ParkingFL_ForR_Label As Label
    Friend WithEvents EepData_CarChime_Label As Label
    Friend WithEvents EepData_HallChime_Label As Label
    Friend WithEvents EepData_ParkingFL_Label As Label
    Friend WithEvents EepData_ParkingOpe_Label As Label
    Friend WithEvents EepData_ParkingSW_Label As Label
    Friend WithEvents EepData_TabPage4 As TabPage
    Friend WithEvents EepData_Page4_GroupBox As GroupBox
    Friend WithEvents EepData_EmerOpe_TextBox As TextBox
    Friend WithEvents EepData_SheaveDia_TextBox As TextBox
    Friend WithEvents EepData_EmerOpe_Label As Label
    Friend WithEvents EepData_MachineType_TextBox As TextBox
    Friend WithEvents EepData_FMSSW_TextBox As TextBox
    Friend WithEvents EepData_FMSSW_Label As Label
    Friend WithEvents EepData_FMSOpe_TextBox As TextBox
    Friend WithEvents EepData_Gear_TextBox As TextBox
    Friend WithEvents EepData_FMSOpe_Label As Label
    Friend WithEvents EepData_FireOpe_TextBox As TextBox
    Friend WithEvents EepData_Inverter_TextBox As TextBox
    Friend WithEvents EepData_FireOpe_Label As Label
    Friend WithEvents EepData_Encoder_TextBox As TextBox
    Friend WithEvents EepData_MotorPole_TextBox As TextBox
    Friend WithEvents EepData_Encoder_Label As Label
    Friend WithEvents EepData_MotorDirection_TextBox As TextBox
    Friend WithEvents EepData_MotorVoltage_TextBox As TextBox
    Friend WithEvents EepData_MotorDirection_Label As Label
    Friend WithEvents EepData_MotorCapacity_TextBox As TextBox
    Friend WithEvents EepData_SheaveDia_Label As Label
    Friend WithEvents EepData_MotorCapacity_Label As Label
    Friend WithEvents EepData_MachineType_Label As Label
    Friend WithEvents EepData_Gear_Label As Label
    Friend WithEvents EepData_MotorVoltage_Label As Label
    Friend WithEvents EepData_Inverter_Label As Label
    Friend WithEvents EepData_MotorPole_Label As Label
    Friend WithEvents EepData_TabPage5 As TabPage
    Friend WithEvents EepData_Page5_GroupBox As GroupBox
    Friend WithEvents EepData_Rope_TextBox As TextBox
    Friend WithEvents EepData_FloodOpe_TextBox As TextBox
    Friend WithEvents EepData_Rope_Label As Label
    Friend WithEvents EepData_Vonic_TextBox As TextBox
    Friend WithEvents EepData_AttOpe_TextBox As TextBox
    Friend WithEvents EepData_AttOpe_Label As Label
    Friend WithEvents EepData_FRDr_TextBox As TextBox
    Friend WithEvents EepData_HIN1_TextBox As TextBox
    Friend WithEvents EepData_FRDr_Label As Label
    Friend WithEvents EepData_WCOB_Spec_TextBox As TextBox
    Friend WithEvents EepData_HIN2_TextBox As TextBox
    Friend WithEvents EepData_WCOB_Spec_Label As Label
    Friend WithEvents EepData_WSCOB_TextBox As TextBox
    Friend WithEvents EepData_HIN3_TextBox As TextBox
    Friend WithEvents EepData_WSCOB_Label As Label
    Friend WithEvents EepData_WCOB_TextBox As TextBox
    Friend WithEvents EepData_HIN4_TextBox As TextBox
    Friend WithEvents EepData_WCOB_Label As Label
    Friend WithEvents EepData_SCOB_TextBox As TextBox
    Friend WithEvents EepData_FloodOpe_Label As Label
    Friend WithEvents EepData_SCOB_Label As Label
    Friend WithEvents EepData_Vonic_Label As Label
    Friend WithEvents EepData_HIN1_Label As Label
    Friend WithEvents EepData_HIN4_Label As Label
    Friend WithEvents EepData_HIN2_Label As Label
    Friend WithEvents EepData_HIN3_Label As Label
    Friend WithEvents EepData_TabPage6 As TabPage
    Friend WithEvents EepData_Page6_GroupBox As GroupBox
    Friend WithEvents EepData_Travel_TextBox As TextBox
    Friend WithEvents EepData_Hight_TextBox As TextBox
    Friend WithEvents EepData_Travel_Label As Label
    Friend WithEvents EepData_Hight_Label As Label
    Friend WithEvents EntityCommand1 As Entity.Core.EntityClient.EntityCommand
    Friend WithEvents Label80 As Label
    Friend WithEvents Label78 As Label
    Friend WithEvents Label77 As Label
End Class
