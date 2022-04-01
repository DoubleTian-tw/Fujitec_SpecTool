Imports System.IO

'Public Class DynamicControlName
Module DynamicControlName

    Public JobMaker_EMER_TabPage As String = "JM_Emer_TabPage"

    Public Spec_EmerGroup_TextBox As String = "Spec_EmerGroup_TextBox"
    Public Spec_EmerCarName_TextBox As String = "Spec_EmerCarName_TextBox"
    Public Spec_EmerEscapeFL_TextBox As String = "Spec_EmerEscapeFL_TextBox"
    Public Spec_EmerReturnFL_TextBox As String = "Spec_EmerReturnFL_TextBox"
    Public Spec_EmerContinue_TextBox As String = "Spec_EmerContinue_TextBox"

    Public Spec_EmerGroup_Label As String = "Spec_EmerGroup_Label"
    Public Spec_EmerCarName_Label As String = "Spec_EmerCarName_Label"
    Public Spec_EmerEscapeFL_Label As String = "Spec_EmerEscapeFL_Label"
    Public Spec_EmerReturnFL_Label As String = "Spec_EmerReturnFL_Label"
    Public Spec_EmerContinue_Label As String = "Spec_EmerContinue_Label"


    'Elvic > Elvator Command
    Public Spec_elaCmd_liftNum_Label As String = "Spec_elaCmd_liftNum_Label"
    Public Spec_elaCmd_Parking_CheckBox As String = "Spec_elaCmd_Parking_CheckBox"
    Public Spec_elaCmd_VIP_CheckBox As String = "Spec_elaCmd_VIP_CheckBox"
    Public Spec_elaCmd_Indepent_CheckBox As String = "Spec_elaCmd_Indepent_CheckBox"
    Public Spec_elaCmd_FloorLockout_CheckBox As String = "Spec_elaCmd_FloorLockout_CheckBox"
    Public Spec_elaCmd_ExpressService_CheckBox As String = "Spec_elaCmd_ExpressService_CheckBox"
    Public Spec_elaCmd_ReturnFloor_CheckBox As String = "Spec_elaCmd_ReturnFloor_CheckBox"
    'Elvic > Group Command
    Public Spec_grpCmd_liftNum_Label As String = "Spec_grpCmd_liftNum_Label"
    Public Spec_grpCmd_UpPeak_CheckBox As String = "Spec_grpCmd_UpPeak_CheckBox"
    Public Spec_grpCmd_DownPeak_CheckBox As String = "Spec_grpCmd_DownPeak_CheckBox"
    Public Spec_grpCmd_LunchTime_CheckBox As String = "Spec_grpCmd_LunchTime_CheckBox"
    Public Spec_grpCmd_MainFL_CheckBox As String = "Spec_grpCmd_MainFL_CheckBox"
    Public Spec_grpCmd_Zoning_CheckBox As String = "Spec_grpCmd_Zoning_CheckBox"
    Public Spec_grpCmd_CarCall_CheckBox As String = "Spec_grpCmd_CarCall_CheckBox"
    'Elvic > Other Command
    Public Spec_otherCmd_liftNum_Label As String = "Spec_otherCmd_liftNum_Label"
    Public Spec_otherCmd_Seismic_CheckBox As String = "Spec_otherCmd_Seismic_CheckBox"
    Public Spec_otherCmd_FireAlarm_CheckBox As String = "Spec_otherCmd_FireAlarm_CheckBox"
    Public Spec_otherCmd_CRD_CheckBox As String = "Spec_otherCmd_CRD_CheckBox"


    Public JobMaker_HIN_TB As String = "JM_HIN_TextBox"
    Public JobMaker_HIN_FlowPanel As String = "JM_HIN_FlowLayoutPanel"
    Public JobMaker_HIN_AllFL_ChkB As String = "HIN_AllFL_CheckBox"
    Public JobMaker_HIN_ChoAuto_ChkB As String = "HIN_choAutoInsert_CheckBox"
    Public JobMaker_HIN_ChoAuto_CmbB As String = "HIN_choAutoInsert_ComboBox"
    Public JobMaker_HIN_FL_ChkB As String = "FL_HIN_CheckBox"
    Public JobMaker_HIN_FL_CmbB As String = "FL_HIN_ComboBox"

    Public mmicBase_CarNo As String = "mmicBase_CarNo"
    Public mmicBase_ObjName As String = "mmicBase_ObjName"
    Public mmicBase_ObjNameBase As String = "mmicBase_ObjNameBase"
    Public mmicEBase_CarNo As String = "mmicEBase_CarNo"
    Public mmicEBase_ObjName As String = "mmicEBase_ObjName"
    Public svBase_CarNo As String = "svBase_CarNo"
    Public svBase_ObjName As String = "svBase_ObjName"
    Public svBase_ObjNameBase As String = "svBase_ObjNameBase"
    Public svEBase_CarNo As String = "svEBase_CarNo"
    Public svEBase_ObjName As String = "svEBase_ObjName"
    Public vd10Base_CarNo As String = "vd10Base_CarNo"
    Public vd10Base_ObjName As String = "vd10Base_ObjName"

    Public intellPC_Label_CarNo As String = "intellPC_Label_CarNo"
    Public intellPC_Label_GSType As String = "intellPC_Label_GSType"
    Public intellPC_Label_IPAddress As String = "intellPC_Label_IPAddress"
    Public intellPC_Soft_CarNo As String = "intellPC_Soft_CarNo"
    Public intellPC_Soft_CDRom As String = "intellPC_Soft_CDRom"
    Public intellPC_Job_CarNo As String = "intellPC_Job_CarNo"
    Public intellPC_Job_CDRom As String = "intellPC_Job_CDRom"

    Public JobMaker_BasicSpecControler_Array As Control()

    Public JobMaker_LiftInfoName_Array() As String
    Public JobMaker_LiftInfoName_output_Array() As String

    Public JobMaker_Elvic_elaCmd_InfoName_Array() As String
    Public JobMaker_Elvic_grpCmd_InfoName_Array() As String
    Public JobMaker_Elvic_otherCmd_InfoName_Array() As String

    Public JobMaker_EmerTBInfoName_Array(),
               JobMaker_EmerLBInfoName_Array() As String
    Public JobMaker_HINInfoName_Array() As String
    Public JobMaker_MMIC_MrBase_InfoName_Array(), JobMaker_MMIC_Mr_InfoName_Array(),
           JobMaker_MMIC_MrEBase_InfoName_Array(),
           JobMaker_MMIC_SvBase_InfoName_Array(), JobMaker_MMIC_Sv_InfoName_Array(),
           JobMaker_MMIC_SvEBase_InfoName_Array(),
           JobMaker_MMIC_VD10Base_InfoName_Array() As String
    Public IntellPC_Label_InfoName_Array(), IntellPC_Soft_InfoName_Array(),
           IntellPC_Job_InfoName_Array() As String

    Public Sub JobMaker_LiftInfo()

        'With JobMaker_Form
        '    JobMaker_BasicSpecControler_Array =
        '            { .Spec_LiftName_TextBox, .Spec_LiftMem_ComboBox,
        '              .Spec_Control_ComboBox, .Spec_TopFL_ComboBox,
        '              .Spec_BtmFL_ComboBox, .Spec_StopFL_ComboBox,
        '              .Spec_Speed_ComboBox, .Spec_FLName_TextBox,
        '              .Spec_MachineType_ComboBox, .Spec_Purpose_ComboBox, .Spec_FLEX_ComboBox}
        'End With
        With JobMaker_Form
            JobMaker_LiftInfoName_Array = { .Spec_LiftName_TextBox.Name, .Spec_LiftMem_ComboBox.Name,
                                            .Spec_Control_ComboBox.Name,
                                            .Spec_TopFL_ComboBox.Name, .Spec_TopFL_Real_ComboBox.Name,
                                            .Spec_BtmFL_ComboBox.Name, .Spec_BtmFL_Real_ComboBox.Name,
                                            .Spec_StopFL_ComboBox.Name,
                                            .Spec_Speed_ComboBox.Name, .Spec_FLName_TextBox.Name,
                                            .Spec_OverBalance_ComboBox.Name,
                                            .Spec_MachineType_ComboBox.Name, .Spec_Purpose_ComboBox.Name, .Spec_FLEX_ComboBox.Name}
        End With
        With JobMaker_Form
            JobMaker_LiftInfoName_output_Array = { .Spec_LiftName_TextBox.Name, .Spec_LiftMem_ComboBox.Name,
                                                       .Spec_Control_ComboBox.Name, .Spec_TopFL_ComboBox.Name,
                                                       .Spec_BtmFL_ComboBox.Name, .Spec_StopFL_ComboBox.Name,
                                                       .Spec_Speed_ComboBox.Name, .Spec_FLName_TextBox.Name}
        End With

    End Sub

    Public Sub JobMaker_ElvicInfo()
        JobMaker_Elvic_elaCmd_InfoName_Array = {Spec_elaCmd_Parking_CheckBox, Spec_elaCmd_VIP_CheckBox,
                                                    Spec_elaCmd_Indepent_CheckBox, Spec_elaCmd_FloorLockout_CheckBox,
                                                    Spec_elaCmd_ExpressService_CheckBox, Spec_elaCmd_ReturnFloor_CheckBox}
        JobMaker_Elvic_grpCmd_InfoName_Array = {Spec_grpCmd_UpPeak_CheckBox, Spec_grpCmd_DownPeak_CheckBox,
                                                    Spec_grpCmd_LunchTime_CheckBox, Spec_grpCmd_MainFL_CheckBox,
                                                    Spec_grpCmd_Zoning_CheckBox, Spec_grpCmd_CarCall_CheckBox}
        JobMaker_Elvic_otherCmd_InfoName_Array = {Spec_otherCmd_Seismic_CheckBox,
                                                      Spec_otherCmd_FireAlarm_CheckBox,
                                                      Spec_otherCmd_CRD_CheckBox}
    End Sub
    Public Sub JobMaker_EmerInfo()
        JobMaker_EmerTBInfoName_Array = {Spec_EmerGroup_TextBox, Spec_EmerCarName_TextBox, Spec_EmerEscapeFL_TextBox,
                                             Spec_EmerReturnFL_TextBox, Spec_EmerContinue_TextBox}
        JobMaker_EmerLBInfoName_Array = {Spec_EmerGroup_Label, Spec_EmerCarName_Label, Spec_EmerEscapeFL_Label,
                                             Spec_EmerReturnFL_Label, Spec_EmerContinue_Label}
    End Sub

    Public Sub JobMaker_HINInfo()
        JobMaker_HINInfoName_Array = {JobMaker_HIN_AllFL_ChkB, JobMaker_HIN_ChoAuto_ChkB, JobMaker_HIN_ChoAuto_CmbB,
                                          JobMaker_HIN_FL_ChkB, JobMaker_HIN_FL_CmbB}
    End Sub

    Public Sub JobMaker_MMICInfo()
        JobMaker_MMIC_MrBase_InfoName_Array = {mmicBase_CarNo, mmicBase_ObjName, mmicBase_ObjNameBase}
        JobMaker_MMIC_Mr_InfoName_Array = {mmicBase_CarNo, mmicBase_ObjName}
        JobMaker_MMIC_MrEBase_InfoName_Array = {mmicEBase_CarNo, mmicEBase_ObjName}
        JobMaker_MMIC_SvBase_InfoName_Array = {svBase_CarNo, svBase_ObjName, svBase_ObjNameBase}
        JobMaker_MMIC_Sv_InfoName_Array = {svBase_CarNo, svBase_ObjName}
        JobMaker_MMIC_SvEBase_InfoName_Array = {svEBase_CarNo, svEBase_ObjName}
        JobMaker_MMIC_VD10Base_InfoName_Array = {vd10Base_CarNo, vd10Base_ObjName}
    End Sub
    Public Sub JobMaker_IntellPCInfo()
        IntellPC_Label_InfoName_Array = {intellPC_Label_CarNo, intellPC_Label_GSType, intellPC_Label_IPAddress}
        IntellPC_Soft_InfoName_Array = {intellPC_Soft_CarNo, intellPC_Soft_CDRom}
        IntellPC_Job_InfoName_Array = {intellPC_Job_CarNo, intellPC_Job_CDRom}
    End Sub

    'End Class
End Module
