Imports System.IO

''' <summary>
''' 動態產生控制項Controls時的Names
''' </summary>
Public Class DynamicControlName

    Public JobMaker_EMER_TabPage As String = "JM_Emer_TabPage"

    Public Spec_MachineType_ComboBox As String = "Spec_MachineType_ComboBox"
    Public Spec_ControlWay_ComboBox As String = "Spec_ControlWay_ComboBox"
    Public Spec_Purpose_ComboBox As String = "Spec_Purpose_ComboBox"
    Public Spec_FLEX_ComboBox As String = "Spec_FLEX_ComboBox"

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

    Public JobMaker_HIN_TB As String = "JM_HIN_TextBox"
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


    Public JobMaker_LiftInfoName_Array() As String
    Public JobMaker_MachinTypeInfoName_Array() As String
    Public JobMaker_ControlWayInfoName_Array() As String
    Public JobMaker_PurposeInfoName_Array() As String
    Public JobMaker_FLEXInfoName_Array() As String

    Public JobMaker_EmerTBInfoName_Array(),
           JobMaker_EmerLBInfoName_Array() As String
    Public JobMaker_HINInfoName_Array() As String
    Public JobMaker_MMIC_MrBase_InfoName_Array(), JobMaker_MMIC_Mr_InfoName_Array(),
           JobMaker_MMIC_MrEBase_InfoName_Array(),
           JobMaker_MMIC_SvBase_InfoName_Array(), JobMaker_MMIC_Sv_InfoName_Array(),
           JobMaker_MMIC_SvEBase_InfoName_Array(),
           JobMaker_MMIC_VD10Base_InfoName_Array() As String

    Public Sub JobMaker_LiftInfo()

        'ReDim JobMaker_LiftInfoName_Array(JobMaker_Form.SpecBasic_LiftItem_Panel.Controls.Count - 1)

        'Dim i As Integer
        'For Each ctrlName As Control In JobMaker_Form.SpecBasic_LiftItem_Panel.Controls
        '    i += 1
        '    JobMaker_LiftInfoName_Array(i - 1) = ctrlName.Name
        '    'MsgBox(JobMaker_LiftInfoName_Array(i - 1))
        'Next
        With JobMaker_Form
            JobMaker_LiftInfoName_Array = { .Spec_LiftName_TextBox.Name, .Spec_LiftMem_ComboBox.Name,
                                            .Spec_Control_ComboBox.Name, .Spec_TopFL_ComboBox.Name,
                                            .Spec_BtmFL_ComboBox.Name, .Spec_StopFL_ComboBox.Name,
                                            .Spec_Speed_ComboBox.Name, .Spec_FLName_TextBox.Name}
        End With
        JobMaker_MachinTypeInfoName_Array = {Spec_MachineType_ComboBox}
        JobMaker_ControlWayInfoName_Array = {Spec_ControlWay_ComboBox}
        JobMaker_PurposeInfoName_Array = {Spec_Purpose_ComboBox}
        JobMaker_FLEXInfoName_Array = {Spec_FLEX_ComboBox}

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


End Class
