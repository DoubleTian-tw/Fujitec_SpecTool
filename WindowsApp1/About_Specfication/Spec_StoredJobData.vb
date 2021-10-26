﻿Imports System.Data.SQLite

Public Class Spec_StoredJobData
    Dim sqlite_connect As SQLiteConnection
    Dim sqlite_cmd As SQLiteCommand

    Dim sqlite_Transaction As SQLiteTransaction
    Dim sqlite_dataReader As SQLiteDataReader

    '[JobMaker > 基本 ] --------------------------------------------
    Public Basic_Use_ChkBox As String = "Basic_Use_ChkBox"
    Public Basic_Local As String = "Basic_Local"
    Public Basic_DesignerChinese As String = "Basic_DesignerChinese"
    Public Basic_DesignerEnglish As String = "Basic_DesignerEnglish"
    Public Basic_CheckerChinese As String = "Basic_CheckerChinese"
    Public Basic_CheckerEnglish As String = "Basic_CheckerEnglish"
    Public Basic_JobNo_New As String = "Basic_JobNo_New"
    Public Basic_JobNo_Old As String = "Basic_JobNo_Old"
    Public Basic_JobNo_Mod As String = "Basic_JobNo_Mod"
    Public Basic_JobName As String = "Basic_JobName"
    Public Basic_DateTimePicker As String = "Basic_DateTimePicker"
    'Public Basic_Spec_AllPages As String = "Basic_Spec_AllPages"
    'Public Basic_FM_AllPages As String = "Basic_FM_AllPages"
    'Public Basic_MachineType As String = "Basic_MachineType"
    'Public Basic_FLEX As String = "Basic_FLEX"
    'Public Basic_OperationType As String = "Basic_OperationType"
    'Public Basic_PRK_Name As String = "Basic_PRK_Name"
    '------------------------------------------- [JobMaker > 基本 ] 

    '[JobMaker > CheckList ] -------------------------------------
    Public ChkList_Use_ChkBox As String = "ChkList_Use_ChkBox"
    Public ChkList_PA_DateTimePicker As String = "ChkList_PA_DateTimePicker"
    Public ChkList_OS_DateTimePicker As String = "ChkList_OS_DateTimePicker"
    Public ChkList_CFM_DateTimePicker As String = "ChkList_CFM_DateTimePicker"
    Public ChkList_ELE_DateTimePicker As String = "ChkList_ELE_DateTimePicker"
    Public ChkList_PA_ChkBox As String = "ChkList_PA_ChkBox"
    Public ChkList_OS_ChkBox As String = "ChkList_OS_ChkBox"
    Public ChkList_CFM_ChkBox As String = "ChkList_CFM_ChkBox"
    Public ChkList_ELE_ChkBox As String = "ChkList_ELE_ChkBox"

    Public ChkList_Q1No_RadioBox As String = "ChkList_Q1No_RadioBox"
    Public ChkList_Q1Yes_RadioBox As String = "ChkList_Q1Yes_RadioBox"
    Public ChkList_Q2No_RadioBox As String = "ChkList_Q2No_RadioBox"
    Public ChkList_Q2Yes_RadioBox As String = "ChkList_Q2Yes_RadioBox"
    Public ChkList_Q3No_RadioBox As String = "ChkList_Q3No_RadioBox"
    Public ChkList_Q3Yes_RadioBox As String = "ChkList_Q3Yes_RadioBox"
    Public ChkList_Q5No_RadioBox As String = "ChkList_Q5No_RadioBox"
    Public ChkList_Q5Std_RadioBox As String = "ChkList_Q5Std_RadioBox"
    Public ChkList_Q5NoStd_RadioBox As String = "ChkList_Q5NoStd_RadioBox"
    Public ChkList_Q6No_RadioBox As String = "ChkList_Q6No_RadioBox"
    Public ChkList_Q6Yes_RadioBox As String = "ChkList_Q6Yes_RadioBox"
    Public ChkList_Q6YesChk_RadioBox As String = "ChkList_Q6YesChk_RadioBox"
    Public ChkList_Q6YesItem_RadioBox As String = "ChkList_Q6YesItem_RadioBox"
    Public ChkList_Q7No_RadioBox As String = "ChkList_Q7No_RadioBox"
    Public ChkList_Q7Yes_RadioBox As String = "ChkList_Q7Yes_RadioBox"
    Public ChkList_Q8No_RadioBox As String = "ChkList_Q8No_RadioBox"
    Public ChkList_Q8Yes_RadioBox As String = "ChkList_Q8Yes_RadioBox"
    Public ChkList_Q8Item_RadioBox As String = "ChkList_Q8Item_RadioBox"
    Public ChkList_Q9No_RadioBox As String = "ChkList_Q9No_RadioBox"
    Public ChkList_Q9Yes_RadioBox As String = "ChkList_Q9Yes_RadioBox"

    Public ChkList_Q1Yes_Content As String = "ChkList_Q1Yes_Content"
    Public ChkList_Q1Yes_Result As String = "ChkList_Q1Yes_Result"
    Public ChkList_Q2Yes_Content As String = "ChkList_Q2Yes_Content"
    Public ChkList_Q2Yes_Result As String = "ChkList_Q2Yes_Result"
    Public ChkList_Q3Yes_Man As String = "ChkList_Q3Yes_Man"
    Public ChkList_Q3Yes_Content As String = "ChkList_Q3Yes_Content"
    Public ChkList_Q3Yes_Result As String = "ChkList_Q3Yes_Result"
    Public ChkList_Q4MMIC As String = "ChkList_Q4MMIC"
    Public ChkList_Q4MmicBase As String = "ChkList_Q4MmicBase"
    Public ChkList_Q4SV As String = "ChkList_Q4SV"
    Public ChkList_Q4SVmicBase As String = "ChkList_Q4SVmicBase"
    Public ChkList_Q5Std_Content As String = "ChkList_Q5Std_Content"
    Public ChkList_Q5nStd_Content As String = "ChkList_Q5nStd_Content"
    Public ChkList_Q6Yes_Content As String = "ChkList_Q6Yes_Content"
    Public ChkList_Q7Yes_Content As String = "ChkList_Q7Yes_Content"
    '------------------------------------- [JobMaker > CheckList ] 

    '[JobMaker > Program Change程式變更 ] ------------------------------------
    Public ChkList_Prgm_Use_ChkBox As String = "ChkList_Prgm_Use_ChkBox"
    Public ChkList_Prgm_1_reason As String = "ChkList_Prgm_1_reason"
    Public ChkList_Prgm_2_Test_ChkBox As String = "ChkList_Prgm_2_Test_ChkBox"
    Public ChkList_Prgm_2_COP_ChkBox As String = "ChkList_Prgm_2_COP_ChkBox"
    Public ChkList_Prgm_2_Tower_ChkBox As String = "ChkList_Prgm_2_Tower_ChkBox"
    Public ChkList_Prgm_2_Other_ChkBox As String = "ChkList_Prgm_2_Other_ChkBox"
    Public ChkList_Prgm_2_Test_Content As String = "ChkList_Prgm_2_Test_Content"
    Public ChkList_Prgm_2_COP_Content As String = "ChkList_Prgm_2_COP_Content"
    Public ChkList_Prgm_2_Tower_Content As String = "ChkList_Prgm_2_Tower_Content"
    Public ChkList_Prgm_2_Other_Content As String = "ChkList_Prgm_2_Other_Content"

    Public ChkList_Prgm_3_Test_ChkBox As String = "ChkList_Prgm_3_Test_ChkBox"
    Public ChkList_Prgm_3_Debug_ChkBox As String = "ChkList_Prgm_3_Debug_ChkBox"
    Public ChkList_Prgm_3_CFM_ChkBox As String = "ChkList_Prgm_3_CFM_ChkBox"
    Public ChkList_Prgm_3_EXE_ChkBox As String = "ChkList_Prgm_3_EXE_ChkBox"
    Public ChkList_Prgm_3_Other_ChkBox As String = "ChkList_Prgm_3_Other_ChkBox"
    Public ChkList_Prgm_3_OtherContent As String = "ChkList_Prgm_3_OtherContent"

    Public ChkList_Prgm_4_1Yes_ChkBox As String = "ChkList_Prgm_4_1Yes_ChkBox"
    Public ChkList_Prgm_4_1No_ChkBox As String = "ChkList_Prgm_4_1No_ChkBox"
    Public ChkList_Prgm_4_2Yes_ChkBox As String = "ChkList_Prgm_4_2Yes_ChkBox"
    Public ChkList_Prgm_4_2No_ChkBox As String = "ChkList_Prgm_4_2No_ChkBox"
    Public ChkList_Prgm_4_3Yes_ChkBox As String = "ChkList_Prgm_4_3Yes_ChkBox"
    Public ChkList_Prgm_4_3No_ChkBox As String = "ChkList_Prgm_4_3No_ChkBox"
    Public ChkList_Prgm_4_4Yes_ChkBox As String = "ChkList_Prgm_4_4Yes_ChkBox"
    Public ChkList_Prgm_4_4No_ChkBox As String = "ChkList_Prgm_4_4No_ChkBox"
    Public ChkList_Prgm_4_5Yes_ChkBox As String = "ChkList_Prgm_4_5Yes_ChkBox"
    Public ChkList_Prgm_4_5No_ChkBox As String = "ChkList_Prgm_4_5No_ChkBox"
    Public ChkList_Prgm_4_6Yes_ChkBox As String = "ChkList_Prgm_4_6Yes_ChkBox"
    Public ChkList_Prgm_4_6No_ChkBox As String = "ChkList_Prgm_4_6No_ChkBox"
    Public ChkList_Prgm_4_7Yes_ChkBox As String = "ChkList_Prgm_4_7Yes_ChkBox"
    Public ChkList_Prgm_4_7No_ChkBox As String = "ChkList_Prgm_4_7No_ChkBox"
    Public ChkList_Prgm_4_8Yes_ChkBox As String = "ChkList_Prgm_4_8Yes_ChkBox"
    Public ChkList_Prgm_4_8No_ChkBox As String = "ChkList_Prgm_4_8No_ChkBox"
    Public ChkList_Prgm_4_9Yes_ChkBox As String = "ChkList_Prgm_4_9Yes_ChkBox"
    Public ChkList_Prgm_4_9No_ChkBox As String = "ChkList_Prgm_4_9No_ChkBox"
    Public ChkList_Prgm_4_10Yes_ChkBox As String = "ChkList_Prgm_4_10Yes_ChkBox"
    Public ChkList_Prgm_4_10No_ChkBox As String = "ChkList_Prgm_4_10No_ChkBox"
    Public ChkList_Prgm_4_11Yes_ChkBox As String = "ChkList_Prgm_4_11Yes_ChkBox"
    Public ChkList_Prgm_4_11No_ChkBox As String = "ChkList_Prgm_4_11No_ChkBox"
    Public ChkList_Prgm_4_12Yes_ChkBox As String = "ChkList_Prgm_4_12Yes_ChkBox"
    Public ChkList_Prgm_4_12No_ChkBox As String = "ChkList_Prgm_4_12No_ChkBox"
    Public ChkList_Prgm_4_TestContent As String = "ChkList_Prgm_4_TestContent"
    '------------------------------------[ JobMaker > Program Change程式變更 ] 

    '[ JobMaker > DWG 送狀 ]-------------------------------------------------
    Public DWG_Vonic As String = "DWG_Vonic"
    Public DWG_Use_ChkBox As String = "DWG_Use_ChkBox"
    Public DWG_Pages As String = "DWG_Pages"
    '-------------------------------------------------[ JobMaker > DWG 送狀 ]


    '[ JobMaker > SPEC仕樣 > Basic ] -------------------------------------------
    Public SpecBasic_Use_ChkBox As String = "SpecBasic_Use_ChkBox"
    Public SpecBasic_LiftNumber As String = "SpecBasic_LiftNumber"
    Public SpecBasic_MachineType As String = "SpecBasic_MachineType"
    Public SpecBasic_MachineType_Number As String = "SpecBasic_MachineType_Number"
    Public SpecBasic_ControlWay As String = "SpecBasic_ControlWay"
    Public SpecBasic_Purpose As String = "SpecBasic_Purpose"
    Public SpecBasic_Purpose_Number As String = "SpecBasic_Purpose_Number"


    'Public SpecBasic_LiftName As String = "SpecBasic_LiftName"
    'Public SpecBasic_LiftMem As String = "SpecBasic_LiftMem"
    'Public SpecBasic_Control As String = "SpecBasic_Control"
    'Public SpecBasic_TopFL As String = "SpecBasic_TopFL"
    'Public SpecBasic_BtmFL As String = "SpecBasic_BtmFL"
    'Public SpecBasic_StopFL As String = "SpecBasic_StopFL"
    'Public SpecBasic_Speed As String = "SpecBasic_Speed"
    'Public SpecBasic_FLName As String = "SpecBasic_FLName"

    'Public SpecBasic_LiftContain() As String = {SpecBasic_LiftName, SpecBasic_LiftMem, SpecBasic_Control, SpecBasic_TopFL,
    '                                            SpecBasic_BtmFL, SpecBasic_StopFL, SpecBasic_Speed, SpecBasic_FLName}
    '------------------------------------------- [ JobMaker > SPEC仕樣 > Basic ] 

    '[ JobMaker > SPEC仕樣 > TW台灣]-------------------------------------------------
    Public SPEC_TW_IDU_CHKBOX As String = "SPEC_TW_IDU_CHKBOX"
    Public SPEC_TW_FP17_CHKBOX As String = "SPEC_TW_FP17_CHKBOX"
    Public SPEC_MACHINE_TYPE As String = "SPEC_MACHINE_TYPE"
    Public SPEC_AUTO_DR As String = "SPEC_AUTO_DR"
    Public SPEC_AUTO_DR_PHOTOEYE As String = "SPEC_AUTO_DR_PHOTOEYE"
    Public SPEC_AUTO_DR_SAFETY As String = "SPEC_AUTO_DR_SAFETY"
    Public SPEC_CANCELL_CALL As String = "SPEC_CANCELL_CALL"
    Public SPEC_CANCELL_CALL_SCOB As String = "SPEC_CANCELL_CALL_SCOB"
    Public SPEC_CANCELL_BEHIND As String = "SPEC_CANCELL_BEHIND"
    Public SPEC_LAMP_CHK As String = "SPEC_LAMP_CHK"
    'Public SPEC_EC_BOOK As String = "SPEC_EC_BOOK"
    'Public SPEC_INSTALL_BOOK As String = "SPEC_INSTALL_BOOK"
    Public SPEC_AUTO_FAN As String = "SPEC_AUTO_FAN"
    Public SPEC_AUTO_FAN_ION_WITHOUT As String = "SPEC_AUTO_FAN_ION_WITHOUT"
    'Public SPEC_AUTO_LIGHT As String = "SPEC_AUTO_LIGHT"
    'Public SPEC_RUN_OPEN As String = "SPEC_RUN_OPEN"
    Public SPEC_CC_CANCEL As String = "SPEC_CC_CANCEL"
    Public SPEC_AUTO_PASS As String = "SPEC_AUTO_PASS"
    'Public SPEC_AUTO_LEVEL As String = "SPEC_AUTO_LEVEL"
    Public SPEC_OPERATION As String = "SPEC_OPERATION"
    Public SPEC_INDEP_OPE As String = "SPEC_INDEP_OPE"
    Public SPEC_INDEP_OPE_CMD As String = "SPEC_INDEP_OPE_CMD"
    Public SPEC_UCMP As String = "SPEC_UCMP"
    Public SPEC_HIN_CPI As String = "SPEC_HIN_CPI"
    Public SPEC_FIRE_OPE As String = "SPEC_FIRE_OPE"
    Public SPEC_FIRE_OPE_SIGNAL As String = "SPEC_FIRE_OPE_SIGNAL"
    Public SPEC_FIRE_ONLY_CHECKBOX As String = "SPEC_FIRE_ONLY_CHECKBOX"
    Public SPEC_FIRE_ONLY_TEXTBOX As String = "SPEC_FIRE_ONLY_TEXTBOX"
    Public SPEC_FIREMAN As String = "SPEC_FIREMAN"
    Public SPEC_FIREMAN_ESCAPE_FL As String = "SPEC_FIREMAN_ESCAPE_FL"
    Public SPEC_FIREMAN_ONLY_CHKBOX As String = "SPEC_FIREMAN_ONLY_CHKBOX"
    Public SPEC_FIREMAN_ONLY As String = "SPEC_FIREMAN_ONLY"

    Public SPEC_PARKING As String = "SPEC_PARKING"
    Public SPEC_PARKING_ONLY_CHECKBOX As String = "SPEC_PARKING_ONLY_CHECKBOX"
    Public SPEC_PARKING_ONLY_TEXTBOX As String = "SPEC_PARKING_ONLY_TEXTBOX"
    Public SPEC_PARKING_FL As String = "SPEC_PARKING_FL"
    Public SPEC_PARKING_ELVIC As String = "SPEC_PARKING_ELVIC"
    Public SPEC_PARKING_WTB As String = "SPEC_PARKING_WTB"
    Public SPEC_PARKING_SHUTDOWN As String = "SPEC_PARKING_SHUTDOWN"
    Public SPEC_PARKING_COB As String = "SPEC_PARKING_COB"
    Public SPEC_PARKING_HALL As String = "SPEC_PARKING_HALL"
    Public SPEC_SEISMIC As String = "SPEC_SEISMIC"
    Public SPEC_SEISMIC_ONLY_CHECKBOX As String = "SPEC_SEISMIC_ONLY_CHECKBOX"
    Public SPEC_SEISMIC_ONLY_TEXTBOX As String = "SPEC_SEISMIC_ONLY_TEXTBOX"
    Public SPEC_SEISMIC_SENSOR As String = "SPEC_SEISMIC_SENSOR"
    Public SPEC_SEISMIC_SENSOR_ONLY_CHECKBOX As String = "SPEC_SEISMIC_SENSOR_ONLY_CHECKBOX"
    Public SPEC_SEISMIC_SENSOR_ONLY_TEXTBOX As String = "SPEC_SEISMIC_SENSOR_ONLY_TEXTBOX"
    Public SPEC_SEISMIC_CANCEL_SW As String = "SPEC_SEISMIC_CANCEL_SW"
    Public SPEC_SEISMIC_CANCEL_SW_ONLY_CHECKBOX As String = "SPEC_SEISMIC_CANCEL_SW_ONLY_CHECKBOX"
    Public SPEC_SEISMIC_CANCEL_SW_ONLY_TEXTBOX As String = "SPEC_SEISMIC_CANCEL_SW_ONLY_TEXTBOX"
    Public SPEC_CPI As String = "SPEC_CPI"
    Public SPEC_CPI_SEISMIC As String = "SPEC_CPI_SEISMIC"
    Public SPEC_CPI_FIRE As String = "SPEC_CPI_FIRE"
    Public SPEC_CPI_EMER As String = "SPEC_CPI_EMER"
    Public SPEC_CPI_FM As String = "SPEC_CPI_FM"
    Public SPEC_CPI_OLT As String = "SPEC_CPI_OLT"
    Public SPEC_CPI_OLT_ONLY_CHECKBOX As String = "SPEC_CPI_OLT_ONLY_CHECKBOX"
    Public SPEC_CPI_OLT_ONLY_TEXTBOX As String = "SPEC_CPI_OLT_ONLY_TEXTBOX"

    Public SPEC_CAR_GONG As String = "SPEC_CAR_GONG"
    Public SPEC_CAR_GONG_CARTOP As String = "SPEC_CAR_GONG_CARTOP"
    Public SPEC_CAR_GONG_CARTOP_CHECKBOX As String = "SPEC_CAR_GONG_CARTOP_CHECKBOX"
    Public SPEC_CAR_GONG_CARTOP_ONLY_CHECKBOX As String = "SPEC_CAR_GONG_CARTOP_ONLY_CHECKBOX"
    Public SPEC_CAR_GONG_CARTOP_ONLY_TEXTBOX As String = "SPEC_CAR_GONG_CARTOP_ONLY_TEXTBOX"

    Public SPEC_CAR_GONG_CARTOPBTM As String = "SPEC_CAR_GONG_CARTOPBTM"
    Public SPEC_CAR_GONG_CARTOPBTM_CHECKBOX As String = "SPEC_CAR_GONG_CARTOPBTM_CHECKBOX"
    Public SPEC_CAR_GONG_CARTOPBTM_ONLY_CHECKBOX As String = "SPEC_CAR_GONG_CARTOPBTM_ONLY_CHECKBOX"
    Public SPEC_CAR_GONG_CARTOPBTM_ONLY_TEXTBOX As String = "SPEC_CAR_GONG_CARTOPBTM_ONLY_TEXTBOX"

    Public SPEC_CAR_GONG_COB As String = "SPEC_CAR_GONG_COB"
    Public SPEC_CAR_GONG_COB_CHECKBOX As String = "SPEC_CAR_GONG_COB_CHECKBOX"
    Public SPEC_CAR_GONG_COB_ONLY_CHECKBOX As String = "SPEC_CAR_GONG_COB_ONLY_CHECKBOX"
    Public SPEC_CAR_GONG_COB_ONLY_TEXTBOX As String = "SPEC_CAR_GONG_COB_ONLY_TEXTBOX"

    Public SPEC_CAR_GONG_VONIC As String = "SPEC_CAR_GONG_VONIC"
    Public SPEC_CAR_GONG_VONIC_CHECKBOX As String = "SPEC_CAR_GONG_VONIC_CHECKBOX"
    Public SPEC_CAR_GONG_VONIC_ONLY_CHECKBOX As String = "SPEC_CAR_GONG_VONIC_ONLY_CHECKBOX"
    Public SPEC_CAR_GONG_VONIC_ONLY_TEXTBOX As String = "SPEC_CAR_GONG_VONIC_ONLY_TEXTBOX"

    Public SPEC_HALL_GONG As String = "SPEC_HALL_GONG"
    Public SPEC_HPI As String = "SPEC_HPI"
    Public SPEC_HPI_OLT As String = "SPEC_HPI_OLT"
    Public SPEC_HPI_MAIN As String = "SPEC_HPI_MAIN"
    Public SPEC_HPI_INDEP As String = "SPEC_HPI_INDEP"
    Public SPEC_HPI_EMER As String = "SPEC_HPI_EMER"
    Public SPEC_DR_HOLD As String = "SPEC_DR_HOLD"
    Public SPEC_CRD As String = "SPEC_CRD"
    Public SPEC_CRD_TYPE As String = "SPEC_CRD_TYPE"
    Public SPEC_CRD_SPEC As String = "SPEC_CRD_SPEC"
    Public SPEC_CRD_RVS_CALL As String = "SPEC_CRD_RVS_CALL"
    Public SPEC_CRD_ANTI As String = "SPEC_CRD_ANTI"
    Public SPEC_CRD_AUTOREGI As String = "SPEC_CRD_AUTOREGI"
    Public SPEC_CRD_ID4 As String = "SPEC_CRD_ID4"
    Public SPEC_CRD_ID5 As String = "SPEC_CRD_ID5"

    Public SPEC_EMER As String = "SPEC_EMER"
    Public SPEC_EMER_NUMBER As String = "SPEC_EMER_NUMBER"
    Public SPEC_EMER_SIGNAL As String = "SPEC_EMER_SIGNAL"
    Public SPEC_EMER_CAPACITY As String = "SPEC_EMER_CAPACITY"
    Public SPEC_EMER_INPUT As String = "SPEC_EMER_INPUT"
    Public SPEC_EMER_ADDRESS As String = "SPEC_EMER_ADDRESS"
    'Public SPEC_EMER_GROUP As String = "SPEC_EMER_GROUP"
    'Public SPEC_EMER_CARNAME As String = "SPEC_EMER_CARNAME"
    'Public SPEC_EMER_ESCAPE_FL As String = "SPEC_EMER_ESCAPE_FL"
    'Public SPEC_EMER_RETURN_FL As String = "SPEC_EMER_RETURN_FL"
    'Public SPEC_EMER_CONTINUE As String = "SPEC_EMER_CONTINUE"
    Public SPEC_LANDIC As String = "SPEC_LANDIC"
    Public SPEC_MFL_RETURN As String = "SPEC_MFL_RETURN"
    Public SPEC_MFL_RETURN_FL As String = "SPEC_MFL_RETURN_FL"
    Public SPEC_VONIC As String = "SPEC_VONIC"
    Public SPEC_VONIC_STANDARD As String = "SPEC_VONIC_STANDARD"
    Public SPEC_ELVIC As String = "SPEC_ELVIC"
    Public SPEC_ELVIC_1_PARKING As String = "SPEC_ELVIC_1_PARKING"
    Public SPEC_ELVIC_1_FL_LOCKOUT As String = "SPEC_ELVIC_1_FL_LOCKOUT"
    Public SPEC_ELVIC_1_VIP As String = "SPEC_ELVIC_1_VIP"
    Public SPEC_ELVIC_1_EXPRESS As String = "SPEC_ELVIC_1_EXPRESS"
    Public SPEC_ELVIC_1_INDEP As String = "SPEC_ELVIC_1_INDEP"
    Public SPEC_ELVIC_1_RETURN As String = "SPEC_ELVIC_1_RETURN"
    Public SPEC_ELVIC_2_TRAFFIC As String = "SPEC_ELVIC_2_TRAFFIC"
    Public SPEC_ELVIC_2_TRAFFIC_UPPEAK As String = "SPEC_ELVIC_2_TRAFFIC_UPPEAK"
    Public SPEC_ELVIC_2_TRAFFIC_DNPEAK As String = "SPEC_ELVIC_2_TRAFFIC_DNPEAK"
    Public SPEC_ELVIC_2_TRAFFIC_LUNCH As String = "SPEC_ELVIC_2_TRAFFIC_LUNCH"
    Public SPEC_ELVIC_2_MFL As String = "SPEC_ELVIC_2_MFL"
    Public SPEC_ELVIC_2_ZONING_EXPRESS As String = "SPEC_ELVIC_2_ZONING_EXPRESS"
    Public SPEC_ELVIC_2_FL_LOCKOUT As String = "SPEC_ELVIC_2_FL_LOCKOUT"
    Public SPEC_ELVIC_2_CARCALL As String = "SPEC_ELVIC_2_CARCALL"
    Public SPEC_ELVIC_3_FIRE As String = "SPEC_ELVIC_3_FIRE"
    Public SPEC_ELVIC_3_WAVIC As String = "SPEC_ELVIC_3_WAVIC"
    Public SPEC_ELVIC_3_CARD As String = "SPEC_ELVIC_3_CARD"

    Public SPEC_WCOB As String = "SPEC_WCOB"
    Public SPEC_WCOB_ONLY_CHECKBOX As String = "SPEC_WCOB_ONLY_CHECKBOX"
    Public SPEC_WCOB_ONLY_TEXTBOX As String = "SPEC_WCOB_ONLY_TEXTBOX"
    Public SPEC_WSCOB As String = "SPEC_WSCOB"
    Public SPEC_WSCOB_ONLY_CHECKBOX As String = "SPEC_WSCOB_ONLY_CHECKBOX"
    Public SPEC_WSCOB_ONLY_TEXTBOX As String = "SPEC_WSCOB_ONLY_TEXTBOX"
    Public SPEC_WCOB_RING As String = "SPEC_WCOB_RING"
    'Public SPEC_WCOB_ONLY As String = "SPEC_WCOB_ONLY"

    Public SPEC_HLL As String = "SPEC_HLL"
    Public SPEC_ATT As String = "SPEC_ATT"
    Public SPEC_FLOOD As String = "SPEC_FLOOD"
    Public SPEC_FLOOD_FL As String = "SPEC_FLOOD_FL"
    Public SPEC_LS1M As String = "SPEC_LS1M"
    Public SPEC_PRU As String = "SPEC_PRU"
    Public SPEC_LOAD_CELL As String = "SPEC_LOAD_CELL"
    Public SPEC_LOAD_CELL_POSITION As String = "SPEC_LOAD_CELL_POSITION"
    Public SPEC_WTB As String = "SPEC_WTB"
    Public SPEC_WTB_ERROR As String = "SPEC_WTB"
    Public SPEC_WTB_STOP As String = "SPEC_WTB_STOP"
    Public SPEC_WTB_FIREMAN As String = "SPEC_WTB_FIREMAN"
    Public SPEC_WTB_NORMAL As String = "SPEC_WTB_NORMAL"
    Public SPEC_WTB_URGENT As String = "SPEC_WTB_URGENT"
    Public SPEC_WTB_FO As String = "SPEC_WTB_FO"
    Public SPEC_WTB_EMER As String = "SPEC_WTB_EMER"
    Public SPEC_WTB_ALART As String = "SPEC_WTB_ALART"
    Public SPEC_WTB_EQ As String = "SPEC_WTB_EQ"
    Public SPEC_WTB_INDEP As String = "SPEC_WTB_INDEP"
    Public SPEC_WTB_EQSW As String = "SPEC_WTB_EQSW"
    Public SPEC_WTB_BZSW As String = "SPEC_WTB_BZSW"
    Public SPEC_WTB_CHKSW As String = "SPEC_WTB_CHKSW"
    Public SPEC_WTB_PKSW As String = "SPEC_WTB_PKSW"
    Public SPEC_WTB_EQIND As String = "SPEC_WTB_EQIND"
    Public SPEC_WTB_EQMAC As String = "SPEC_WTB_EQMAC"

    Public SPEC_FRONT_REAR_DR As String = "SPEC_FRONT_REAR_DR"
    Public SPEC_EACH_STOP As String = "SPEC_EACH_STOP"
    Public SPEC_INSTALL_OPE As String = "SPEC_INSTALL_OPE"
    Public SPEC_VONICBZ As String = "SPEC_VONICBZ"
    'Public SPEC_FORCE_CLOSE As String = "SPEC_FORCE_CLOSE"
    Public SPEC_OPE_SW As String = "SPEC_OPE_SW"
    Public SPEC_OPE_SW_POS As String = "SPEC_OPE_SW_POS"
    Public SPEC_OPE_SW_ADDRESS As String = "SPEC_OPE_SW_ADDRESS"
    '-------------------------------------------------[ JobMaker > SPEC仕樣 ]

    '[ JobMaker > 重要設定 ] -------------------------------------------------
    Public IMPORTANT_Use_ChkBox As String = "IMPORTANT_Use_ChkBox"
    Public IMPORTANT_FAN As String = "IMPORTANT_FAN"
    Public IMPORTANT_BALANCE As String = "IMPORTANT_BALANCE"
    Public IMPORTANT_WCOB As String = "IMPORTANT_WCOB"
    Public IMPORTANT_DOOR As String = "IMPORTANT_DOOR"
    Public IMPORTANT_HIN As String = "IMPORTANT_HIN" 'not yet
    Public IMPORTANT_HIN_FL As String = "IMPORTANT_HIN_FL" 'not yet
    '------------------------------------------------- [ JobMaker > 重要設定 ] 

    '[ JobMaker > MMIC ] -------------------------------------------------
    Public MMIC_Use_ChkBox As String = "MMIC_Use_ChkBox"
    Public MMIC_MachineType As String = "MMIC_MACHINETYPE"
    Public MMIC_FLEX As String = "MMIC_FLEX"

    Public MMIC_MR_BASE As String = "MMIC_MR_BASE"
    Public MMIC_MR_CP43x As String = "MMIC_MR_CP43x"
    Public MMIC_MR_Number As String = "MMIC_MR_Number"
    Public MMIC_MR_CarNo As String = "MMIC_MR_CarNo"
    Public MMIC_MR_ObjName As String = "MMIC_MR_ObjName"
    Public MMIC_MR_ObjNameBase As String = "MMIC_MR_ObjNameBase"

    Public MMIC_MR_EBase As String = "MMIC_MR_EBase"
    Public MMIC_MR_ECarObj As String = "MMIC_MR_EBase"
    Public MMIC_MR_ENumber As String = "MMIC_MR_ENumber"
    Public MMIC_MR_ECarNo As String = "MMIC_MR_ECarNo"
    Public MMIC_MR_EObjName As String = "MMIC_MR_EObjName"

    Public MMIC_SV_BASE As String = "MMIC_SV_BASE"
    Public MMIC_SV_TYPE As String = "MMIC_SV_TYPE"
    Public MMIC_SV_Number As String = "MMIC_SV_Number"
    Public MMIC_SV_CarNo As String = "MMIC_SV_CarNo"
    Public MMIC_SV_ObjName As String = "MMIC_SV_ObjName"
    Public MMIC_SV_ObjNameBase As String = "MMIC_SV_ObjNameBase"

    Public MMIC_SV_EBase As String = "MMIC_SV_EBase"
    Public MMIC_SV_ECarObj As String = "MMIC_SV_ECarObj"
    Public MMIC_SV_ENumber As String = "MMIC_SV_ENumber"
    Public MMIC_SV_ECarNo As String = "MMIC_SV_ECarNo"
    Public MMIC_SV_EObjName As String = "MMIC_SV_EObjName"


    Public MMIC_VD10_ROM As String = "MMIC_VD10_ROM"
    Public MMIC_VD10_Quantity As String = "MMIC_VD10_Quantity"
    Public MMIC_VD10_TYPE As String = "MMIC_VD10_TYPE"
    Public MMIC_VD10_BASE As String = "MMIC_VD10_BASE"
    Public MMIC_VD10_Number As String = "MMIC_VD10_Number"
    Public MMIC_VD10_CarNo As String = "MMIC_VD10_CarNo"
    Public MMIC_VD10_ObjName As String = "MMIC_VD10_ObjName"
    '------------------------------------------------- [ JobMaker > MMIC ] 

    Public SQLite_connectionPath_Tool As String = "M:\DESIGN\BACK UP\yc_tian\Tool Application\SQLite\" 'SQLite的檔案位置
    Public SQLite_connectionPath_Job As String = $"{SQLite_connectionPath_Tool}JOB\" 'SQLite的檔案位置
    'Public SQLite_connectionPath_Tool As String = "M:\DESIGN\BACK UP\yc_tian\SQLite\" 'SQLite的檔案位置
    'Public SQLite_connectionPath_Job As String = "M:\DESIGN\BACK UP\yc_tian\SQLite\JOB\" 'SQLite的檔案位置
    Public SQLite_ToolDBMS_Name As String = "Tool_Database.sqlite"
    Public SQLite_StdJobDataDBMS_Name As String = "Standard_StoredJobData.sqlite"
    Public SQLite_JobDBMS_Name As String

    Public SQLite_tableName_AllProgramType As String = "AllProgramType"
    Public SQLite_tableName_Basic As String = "BasicSetting"
    Public SQLite_tableName_CheckList As String = "CheckListSetting"
    Public SQLite_tableName_Program As String = "CheckList_PrgmSetting"
    Public SQLite_tableName_DWG As String = "DWGSetting"
    Public SQLite_tableName_SpecBasic As String = "SpecBasicSetting"
    Public SQLite_tableName_SpecTW As String = "SpecTWSetting"
    Public SQLite_tableName_Important As String = "SpecImportantSetting"
    Public SQLite_tableName_MMIC As String = "SpecMMICSetting"

    ''' <summary>
    ''' SQLite語法
    ''' </summary>
    Private SQLite_storedGrammer As String
    ''' <summary>
    ''' 選擇要插入或是更新資料，True為Insert / False為update
    ''' </summary>
    Private updateOrInsert_bool As Boolean
    ''' <summary>
    ''' 是否覆蓋? True:是/False:否
    ''' </summary>
    Private coverFile_bool As Boolean

    '更新或新建SQLite ------------------------------------------
    ''' <summary>
    ''' 新建或覆蓋檔案內容
    ''' </summary>
    ''' <param name="job_dbms">Job檔案名稱</param>
    ''' <param name="coverFile">是否覆蓋? True:是/False:否</param>
    Public Sub Update_Stored(job_dbms As String, coverFile As Boolean)
        Try
            SQLite_JobDBMS_Name = job_dbms
            coverFile_bool = coverFile

            If JobMaker_Form.Use_Basic_CheckBox.Checked Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"更新 「基本」 開始 ======================= {vbCrLf}"
                Basic_Stored()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================更新 「基本」 結束 {vbCrLf}"
            End If
            If JobMaker_Form.Use_ChkList_CheckBox.Checked Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"更新 「CheckList」 開始 ======================= {vbCrLf}"
                CheckList_Stored()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================更新 「CheckList」 結束 {vbCrLf}"
            End If
            If JobMaker_Form.Use_Program_CheckBox.Checked Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"更新 「程式變更」 開始 ======================= {vbCrLf}"
                ProgramChange_Stored()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================更新 「程式變更」 結束 {vbCrLf}"
            End If
            'If JobMaker_Form.Use_prk_CheckBox.Checked Then
            '    JobMaker_Form.ResultOutput_TextBox.Text += $"更新 「送狀」 開始 ======================= {vbCrLf}"
            '    DWG_Stored()
            '    JobMaker_Form.ResultOutput_TextBox.Text += $"=======================更新 「送狀」 結束 {vbCrLf}"
            'End If
            If JobMaker_Form.Use_SpecBasic_CheckBox.Checked Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"更新 「仕樣」 開始 ======================= {vbCrLf}"
                SpecBasic_Stored()
                SpecTW_Stored()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================更新 「仕樣」 結束 {vbCrLf}"
            End If
            If JobMaker_Form.Use_Imp_CheckBox.Checked Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"更新 「重要設定」 開始 ======================= {vbCrLf}"
                Important_Stored()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================更新 「重要設定」 結束 {vbCrLf}"
            End If
            If JobMaker_Form.Use_mmic_CheckBox.Checked Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"更新 「MMIC」 開始 ======================= {vbCrLf}"
                MMIC_Stored()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================更新 「MMIC」 結束 {vbCrLf}"
            End If

            MsgBox($"寫入成功",, "Fine")
        Catch e As Exception
            MsgBox($"寫入失敗 : {e}",, "Fail")
            JobMaker_Form.ResultOutput_TextBox.Text += $"寫入失敗 : {e} {vbCrLf}"
        End Try

    End Sub

    Private Sub Basic_Stored()
        '基本 ----------------------------------------------------------
        '是否使用分頁
        update_DbmsData(Basic_Use_ChkBox,
                        JobMaker_Form.Use_Basic_CheckBox.Checked,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Local
        update_DbmsData(Basic_Local,
                        JobMaker_Form.Basic_Local_ComboBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'JobNo(New)
        update_DbmsData(Basic_JobNo_New,
                        JobMaker_Form.Basic_JobNoNew_TextBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'JobNo(Old)
        update_DbmsData(Basic_JobNo_Old,
                        JobMaker_Form.Basic_JobNoOld_TextBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'JobNo(Mod)
        update_DbmsData(Basic_JobNo_Mod,
                        JobMaker_Form.Basic_JobNoMOD_TextBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'JobName
        update_DbmsData(Basic_JobName,
                        JobMaker_Form.Basic_JobName_TextBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'DesignerChinese
        update_DbmsData(Basic_DesignerChinese,
                        JobMaker_Form.Basic_DesingerChinese_ComboBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'DesignerEnglish
        update_DbmsData(Basic_DesignerEnglish,
                        JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'CheckerChinese
        update_DbmsData(Basic_CheckerChinese,
                        JobMaker_Form.Basic_CheckerChinese_ComboBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'CheckerEnglish
        update_DbmsData(Basic_CheckerEnglish,
                        JobMaker_Form.Basic_CheckerEnglish_ComboBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Date Time Picker
        update_DbmsData(Basic_DateTimePicker,
                        JobMaker_Form.Basic_DrawDate_DateTimePicker.Value.ToString,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '----------------------------------------------------------基本 
    End Sub
    Private Sub CheckList_Stored()
        'CheckList ---------------------------------------------------
        '是否使用分頁
        update_DbmsData(ChkList_Use_ChkBox,
                        JobMaker_Form.Use_ChkList_CheckBox.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'PA DateTimePicker
        update_DbmsData(ChkList_PA_DateTimePicker,
                        JobMaker_Form.ChkList_PaSheet_DateTimePicker.Value.ToString,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'OS DateTimePicker
        update_DbmsData(ChkList_OS_DateTimePicker,
                        JobMaker_Form.ChkList_OS_DateTimePicker.Value.ToString,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'CFM DateTimePicker
        update_DbmsData(ChkList_CFM_DateTimePicker,
                        JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.ToString,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELE DateTimePicker
        update_DbmsData(ChkList_ELE_DateTimePicker,
                        JobMaker_Form.ChkList_Elec_DateTimePicker.Value.ToString,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'PA CheckBox
        update_DbmsData(ChkList_PA_ChkBox,
                        JobMaker_Form.ChkList_PaSheet_CheckBox.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'OS CheckBox
        update_DbmsData(ChkList_OS_ChkBox,
                        JobMaker_Form.ChkList_OS_CheckBox.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'CFM CheckBox
        update_DbmsData(ChkList_CFM_ChkBox,
                        JobMaker_Form.ChkList_Confirm_CheckBox.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELE CheckBox
        update_DbmsData(ChkList_ELE_ChkBox,
                        JobMaker_Form.ChkList_Elec_CheckBox.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'Q1 no RadioButton
        update_DbmsData(ChkList_Q1No_RadioBox,
                        JobMaker_Form.ChkList_1_no_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q1 yes RadioButton
        update_DbmsData(ChkList_Q1Yes_RadioBox,
                        JobMaker_Form.ChkList_1_yes_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q1 yes 討論內容
        update_DbmsData(ChkList_Q1Yes_Content,
                        JobMaker_Form.ChkList_1_yes_Content_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q1 yes 結果
        update_DbmsData(ChkList_Q1Yes_Result,
                        JobMaker_Form.ChkList_1_yes_Content_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'Q2 no RadioButton
        update_DbmsData(ChkList_Q2No_RadioBox,
                        JobMaker_Form.ChkList_2_no_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q2 yes RadioButton
        update_DbmsData(ChkList_Q2Yes_RadioBox,
                        JobMaker_Form.ChkList_2_yes_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q2 yes 指出內容
        update_DbmsData(ChkList_Q2Yes_Content,
                        JobMaker_Form.ChkList_2_yes_Content_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q2 yes 結果
        update_DbmsData(ChkList_Q2Yes_Result,
                        JobMaker_Form.ChkList_2_yes_Result_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'Q3 no RadioButton
        update_DbmsData(ChkList_Q3No_RadioBox,
                        JobMaker_Form.ChkList_3_no_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q3 yes RadioButton
        update_DbmsData(ChkList_Q3Yes_RadioBox,
                        JobMaker_Form.ChkList_3_yes_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q3 yes 討論者
        update_DbmsData(ChkList_Q3Yes_Man,
                        JobMaker_Form.ChkList_3_yes_Man_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'Q3 yes 內容
        update_DbmsData(ChkList_Q3Yes_Content,
                        JobMaker_Form.ChkList_3_yes_Content_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'Q3 yes 結論
        update_DbmsData(ChkList_Q3Yes_Result,
                        JobMaker_Form.ChkList_3_yes_Result_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'Q5 no CheckBox
        update_DbmsData(ChkList_Q5No_RadioBox,
                        JobMaker_Form.ChkList_5_no_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q5 std CheckBox
        update_DbmsData(ChkList_Q5Std_RadioBox,
                        JobMaker_Form.ChkList_5_std_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q5 noStd CheckBox
        update_DbmsData(ChkList_Q5NoStd_RadioBox,
                        JobMaker_Form.ChkList_5_nstd_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q5 std 標準
        update_DbmsData(ChkList_Q5NoStd_RadioBox,
                        JobMaker_Form.ChkList_5_std_Content_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q5 noStd 工直
        update_DbmsData(ChkList_Q5NoStd_RadioBox,
                        JobMaker_Form.ChkList_5_nstd_Content_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'Q6 no CheckBox
        update_DbmsData(ChkList_Q6No_RadioBox,
                        JobMaker_Form.ChkList_6_no_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q6 yes CheckBox
        update_DbmsData(ChkList_Q6Yes_RadioBox,
                        JobMaker_Form.ChkList_6_yes_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q6 yes Check CheckBox
        update_DbmsData(ChkList_Q6YesChk_RadioBox,
                        JobMaker_Form.ChkList_6_yesChk_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q6 yes Item CheckBox
        update_DbmsData(ChkList_Q6YesItem_RadioBox,
                        JobMaker_Form.ChkList_6_yesItem_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q6 yes 檢驗項目
        update_DbmsData(ChkList_Q6Yes_Content,
                        JobMaker_Form.ChkList_6_yes_Content_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q7 no CheckBox
        update_DbmsData(ChkList_Q7No_RadioBox,
                        JobMaker_Form.ChkList_7_no_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q7 yes CheckBox
        update_DbmsData(ChkList_Q7Yes_RadioBox,
                        JobMaker_Form.ChkList_7_yes_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q7 yes 文書
        update_DbmsData(ChkList_Q7Yes_Content,
                        JobMaker_Form.ChkList_7_yes1_content_TextBox.Text,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q8 no CheckBox
        update_DbmsData(ChkList_Q8No_RadioBox,
                        JobMaker_Form.ChkList_8_no_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q8 yes CheckBox
        update_DbmsData(ChkList_Q8Yes_RadioBox,
                        JobMaker_Form.ChkList_8_yes_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q8 item CheckBox
        update_DbmsData(ChkList_Q8Item_RadioBox,
                        JobMaker_Form.ChkList_8Item_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q9 no CheckBox
        update_DbmsData(ChkList_Q9No_RadioBox,
                        JobMaker_Form.ChkList_9_no_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Q9 yes CheckBox
        update_DbmsData(ChkList_Q9Yes_RadioBox,
                        JobMaker_Form.ChkList_9_yes_RadioButton.Checked,
                        SQLite_tableName_CheckList,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '--------------------------------------------------- CheckList
    End Sub
    Private Sub ProgramChange_Stored()
        'CheckList ---------------------------------------------------
        '是否使用分頁
        update_DbmsData(ChkList_Prgm_Use_ChkBox,
                        JobMaker_Form.Use_Program_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '1.變更理由
        update_DbmsData(ChkList_Prgm_1_reason,
                        JobMaker_Form.PrmList_1_reason_TextBox.Text,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '2.測試裝置CheckBox
        update_DbmsData(ChkList_Prgm_2_Test_ChkBox,
                        JobMaker_Form.PrmList_2_test_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '2.控制盤CheckBox
        update_DbmsData(ChkList_Prgm_2_COP_ChkBox,
                        JobMaker_Form.PrmList_2_COP_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '2.研修測試塔CheckBox
        update_DbmsData(ChkList_Prgm_2_Tower_ChkBox,
                        JobMaker_Form.PrmList_2_Tower_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '2.其他CheckBox
        update_DbmsData(ChkList_Prgm_2_Other_ChkBox,
                        JobMaker_Form.PrmList_2_Other_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '2.測試裝置TextBox
        update_DbmsData(ChkList_Prgm_2_Test_Content,
                        JobMaker_Form.PrmList_2_test_TextBox.Text,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '2.控制盤TextBox
        update_DbmsData(ChkList_Prgm_2_COP_Content,
                        JobMaker_Form.PrmList_2_COP_TextBox.Text,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '2.研修測試塔TextBox
        update_DbmsData(ChkList_Prgm_2_Tower_Content,
                        JobMaker_Form.PrmList_2_tower_TextBox.Text,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '2.其他TextBox
        update_DbmsData(ChkList_Prgm_2_Other_Content,
                        JobMaker_Form.PrmList_2_other_TextBox.Text,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '3.Debug CheckBox
        update_DbmsData(ChkList_Prgm_3_Debug_ChkBox,
                        JobMaker_Form.PrmList_3_debug_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '3.Test CheckBox
        update_DbmsData(ChkList_Prgm_3_Test_ChkBox,
                        JobMaker_Form.PrmList_3_test_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '3.Confrim CheckBox
        update_DbmsData(ChkList_Prgm_3_CFM_ChkBox,
                        JobMaker_Form.PrmList_3_confirm_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '3.Execution CheckBox
        update_DbmsData(ChkList_Prgm_3_EXE_ChkBox,
                        JobMaker_Form.PrmList_3_excute_CheckBox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '3.Other CheckBox
        update_DbmsData(ChkList_Prgm_3_Other_ChkBox,
                        JobMaker_Form.PrmList_3_other_Checkbox.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '3.Other TextBox
        update_DbmsData(ChkList_Prgm_3_OtherContent,
                        JobMaker_Form.PrmList_3_other_TextBox.Text,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.1 AutoYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_1Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes1_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.1 AutoNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_1No_ChkBox,
                        JobMaker_Form.PrmList_4_no1_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.2 InputYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_2Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes2_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.2 InputNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_2No_ChkBox,
                        JobMaker_Form.PrmList_4_no2_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.3 IniYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_3Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes3_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.3 IniNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_3No_ChkBox,
                        JobMaker_Form.PrmList_4_no3_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.4 CaseYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_4Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes4_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.4 CaseNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_4No_ChkBox,
                        JobMaker_Form.PrmList_4_no4_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.5 IfYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_5Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes5_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '4.5 IfNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_5No_ChkBox,
                        JobMaker_Form.PrmList_4_no5_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.6 LoopYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_6Yes_ChkBox,
                        JobMaker_Form.PrmList_4_no6_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.6 LoopNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_6No_ChkBox,
                        JobMaker_Form.PrmList_4_no6_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.7 RangeYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_7Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes7_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.7 RangeNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_7No_ChkBox,
                        JobMaker_Form.PrmList_4_no7_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.8 CastingYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_8Yes_ChkBox,
                        JobMaker_Form.PrmList_4_no8_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.8 CastingNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_8No_ChkBox,
                        JobMaker_Form.PrmList_4_yes8_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.9 0Yes RadioBtn
        update_DbmsData(ChkList_Prgm_4_9Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes9_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.9 0No RadioBtn
        update_DbmsData(ChkList_Prgm_4_9No_ChkBox,
                        JobMaker_Form.PrmList_4_no9_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.10 CountYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_10Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes10_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.10 CountNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_10No_ChkBox,
                        JobMaker_Form.PrmList_4_no10_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.11 AddressYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_11Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes11_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.11 AddressNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_11No_ChkBox,
                        JobMaker_Form.PrmList_4_no11_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.12 CustomYes RadioBtn
        update_DbmsData(ChkList_Prgm_4_12Yes_ChkBox,
                        JobMaker_Form.PrmList_4_yes12_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.12 CustomNo RadioBtn
        update_DbmsData(ChkList_Prgm_4_12No_ChkBox,
                        JobMaker_Form.PrmList_4_no12_RadioButton.Checked,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '4.12 Content RadioBtn
        update_DbmsData(ChkList_Prgm_4_TestContent,
                        JobMaker_Form.PrmList_4_content12_TextBox.Text,
                        SQLite_tableName_Program,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
    End Sub
    Private Sub DWG_Stored()
        'DWG ---------------------------------------------------
        '是否使用分頁
        update_DbmsData(DWG_Use_ChkBox,
                        JobMaker_Form.Use_prk_CheckBox.Checked,
                        SQLite_tableName_DWG,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'VONIC標準
        update_DbmsData(DWG_Vonic,
                        JobMaker_Form.DWG_VonicStd_ComboBox.Text,
                        SQLite_tableName_DWG,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '輸出項目必要打勾 DWG_Page_CheckListBox 
        If coverFile_bool = False Then
            For pageInsert_i = 1 To JobMaker_Form.DWG_Page_CheckedListBox.Items.Count - 1
                Insert_DbmsData(DWG_Pages,
                                SQLite_tableName_DWG,
                                SQLite_connectionPath_Job,
                                SQLite_JobDBMS_Name)
                update_DbmsData(DWG_Pages,
                                JobMaker_Form.DWG_Page_CheckedListBox.Items(pageInsert_i - 1).ToString,
                                SQLite_tableName_DWG,
                                SQLite_connectionPath_Job,
                                SQLite_JobDBMS_Name)
            Next

        Else

            '計算sqlite與windows form中的checklistbox數目 ---------------------
            Dim sqlite_dwgPages_count, chkListBox_dwgPages_count As Integer
            sqlite_dwgPages_count = read_DbmsData_CountRow(DWG_Pages,
                                                           SQLite_tableName_DWG,
                                                           SQLite_connectionPath_Job,
                                                           SQLite_JobDBMS_Name)
            chkListBox_dwgPages_count = JobMaker_Form.DWG_Page_CheckedListBox.Items.Count
            '--------------------- 計算sqlite與windows form中的checklistbox數目 

            '如果 數量不同 或 數量相同但文字不同 表示需要更新 ---------------------
            Dim overwrite_dwgPage_bool As Boolean
            '--------------------- 如果 數量不同 或 數量相同但文字不同 表示需要更新 

            If sqlite_dwgPages_count <> chkListBox_dwgPages_count Then
                '數量不同
                overwrite_dwgPage_bool = True
            Else
                '數量相同但文字不同
                For dwgPage_i = 1 To chkListBox_dwgPages_count
                    If JobMaker_Form.DWG_Page_CheckedListBox.Items(dwgPage_i - 1).ToString <>
                        read_DbmsData_RowID(DWG_Pages, SQLite_tableName_DWG, SQLite_connectionPath_Job,
                                            SQLite_JobDBMS_Name, dwgPage_i) Then
                        overwrite_dwgPage_bool = True
                        Exit For
                    Else
                        overwrite_dwgPage_bool = False
                    End If
                Next
            End If

            If overwrite_dwgPage_bool Then
                '當下更新的送狀Page CheckListBox與紀錄中的比較，如果有一處不同就全數刪除設="" -------
                update_DbmsData(DWG_Pages,
                                "",
                                SQLite_tableName_DWG,
                                SQLite_connectionPath_Job,
                                SQLite_JobDBMS_Name)
                '-------當下更新的送狀Page CheckListBox與紀錄中的比較，如果有一處不同就全數刪除設="" 

                '更新新的CheckListBox ----------------------------------------------------------------
                For dwgPageUpdate_i = 1 To chkListBox_dwgPages_count
                    'If JobMaker_Form.DWG_Page_CheckedListBox.GetItemCheckState(dwgPageUpdate_i - 1) = CheckState.Checked Then
                    update_DbmsData(DWG_Pages,
                                    JobMaker_Form.DWG_Page_CheckedListBox.Items.Item(dwgPageUpdate_i - 1).ToString,
                                    SQLite_tableName_DWG,
                                    SQLite_connectionPath_Job,
                                    SQLite_JobDBMS_Name,
                                    dwgPageUpdate_i)
                    'End If
                Next
                '---------------------------------------------------------------- 更新新的CheckListBox 
            End If
        End If
        '--------------------------------------------------- DWG 
    End Sub
    Private Sub SpecBasic_Stored()
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        'SPEC ---------------------------------------------------
        '是否使用分頁
        update_DbmsData(SpecBasic_Use_ChkBox,
                        JobMaker_Form.Use_SpecBasic_CheckBox.Checked,
                        SQLite_tableName_SpecBasic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '機種 / 控制方式 
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.Spec_MachineType_NumericUpDown,
                                    {dyCtrlName.Spec_MachineType_ComboBox, dyCtrlName.Spec_ControlWay_ComboBox}.Count,
                                    {dyCtrlName.Spec_MachineType_ComboBox, dyCtrlName.Spec_ControlWay_ComboBox},
                                    JobMaker_Form.Spec_MachineType_Panel,
                                    SpecBasic_MachineType_Number,
                                    SQLite_tableName_SpecBasic)
        '機種 / 控制方式 數量
        update_DbmsData(SpecBasic_MachineType_Number,
                        JobMaker_Form.Spec_MachineType_NumericUpDown.Value,
                        SQLite_tableName_SpecBasic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '用途 
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.Spec_Purpose_NumericUpDown,
                                    {dyCtrlName.Spec_Purpose_ComboBox}.Count,
                                    {dyCtrlName.Spec_Purpose_ComboBox},
                                    JobMaker_Form.Spec_Purpose_Panel,
                                    SpecBasic_Purpose_Number,
                                    SQLite_tableName_SpecBasic)
        '用途 數量
        update_DbmsData(SpecBasic_Purpose_Number,
                        JobMaker_Form.Spec_Purpose_NumericUpDown.Value,
                        SQLite_tableName_SpecBasic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'panel中的號機基本資訊 -------------------------------------------------------------------

        dyCtrlName.JobMaker_LiftInfo()

        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.Spec_LiftNum_NumericUpDown,
                                    dyCtrlName.JobMaker_LiftInfoName_Array.Count,
                                    dyCtrlName.JobMaker_LiftInfoName_Array,
                                    JobMaker_Form.SpecBasic_LiftItem_Dynamic_Panel,
                                    SpecBasic_LiftNumber,
                                    SQLite_tableName_SpecBasic)

        '電梯總數
        update_DbmsData(SpecBasic_LiftNumber,
                        JobMaker_Form.Spec_LiftNum_NumericUpDown.Value,
                        SQLite_tableName_SpecBasic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '------------------------------------------------------------------- panel中的號機基本資訊 


    End Sub
    Private Sub SpecTW_Stored()
        'SPEC ---------------------------------------------------
        'IDU CheckBox
        update_DbmsData(SPEC_TW_IDU_CHKBOX,
                        JobMaker_Form.Use_SpecTWIDU_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'IDU CheckBox
        update_DbmsData(SPEC_MACHINE_TYPE,
                        JobMaker_Form.Spec_Base_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'FP17 CheckBox
        update_DbmsData(SPEC_TW_FP17_CHKBOX,
                        JobMaker_Form.Use_SpecTWFP17_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '機種
        update_DbmsData(SPEC_MACHINE_TYPE,
                        JobMaker_Form.Spec_Base_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門時限自動調節
        update_DbmsData(SPEC_AUTO_DR,
                        JobMaker_Form.Spec_DRAuto_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門時限自動調節-光電裝置
        update_DbmsData(SPEC_AUTO_DR_PHOTOEYE,
                        JobMaker_Form.Spec_PhotoEye_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門時限自動調節-機械式裝置
        update_DbmsData(SPEC_AUTO_DR_SAFETY,
                        JobMaker_Form.Spec_MechSafety_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '取消嬉戲呼叫
        update_DbmsData(SPEC_CANCELL_CALL,
                        JobMaker_Form.Spec_CancellCall_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '取消嬉戲呼叫-副COB
        update_DbmsData(SPEC_CANCELL_CALL_SCOB,
                        JobMaker_Form.Spec_SCOB_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '逆呼無效
        update_DbmsData(SPEC_CANCELL_BEHIND,
                        JobMaker_Form.Spec_CancellBehind_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '燈點檢模式
        update_DbmsData(SPEC_LAMP_CHK,
                        JobMaker_Form.Spec_LampChk_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '風扇連動
        update_DbmsData(SPEC_AUTO_FAN,
                        JobMaker_Form.Spec_AutoFan_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '風扇連動-離子除菌
        update_DbmsData(SPEC_AUTO_FAN_ION_WITHOUT,
                        JobMaker_Form.Spec_ION_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂呼叫取消
        update_DbmsData(SPEC_CC_CANCEL,
                        JobMaker_Form.Spec_CCCancell_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '自動滿員通過
        update_DbmsData(SPEC_AUTO_PASS,
                        JobMaker_Form.Spec_AutoPass_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '操作方式
        'update_DbmsData(SPEC_OPERATION,
        '                JobMaker_Form.Spec_Operation_ComboBox.Text,
        '                SQLite_tableName_SpecTW,
        '                SQLite_connectionPath_Job,
        '                SQLite_JobDBMS_Name)
        '專用運轉
        update_DbmsData(SPEC_INDEP_OPE,
                        JobMaker_Form.Spec_Indep_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '戶開行走保護
        update_DbmsData(SPEC_UCMP,
                        JobMaker_Form.Spec_UCMP_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'HIN/CPI
        update_DbmsData(SPEC_HIN_CPI,
                        JobMaker_Form.Spec_HinCpi_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '火災管制運轉
        update_DbmsData(SPEC_FIRE_OPE,
                        JobMaker_Form.Spec_Fire_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '火災管制運轉-訊號
        update_DbmsData(SPEC_FIRE_OPE_SIGNAL,
                        JobMaker_Form.Spec_FireSignal_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '火災管制運轉-Only checkBox
        update_DbmsData(SPEC_FIRE_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Fire_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '火災管制運轉-Only TextBox
        update_DbmsData(SPEC_FIRE_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Fireman_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '消防梯運轉
        update_DbmsData(SPEC_FIREMAN,
                        JobMaker_Form.Spec_Fireman_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '消防梯運轉-避難階
        update_DbmsData(SPEC_FIREMAN_ESCAPE_FL,
                        JobMaker_Form.Spec_EscapeFL_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '消防梯運轉-Only n 號機 CheckBox
        update_DbmsData(SPEC_FIREMAN_ONLY_CHKBOX,
                        JobMaker_Form.Spec_Fireman_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '消防梯運轉-Only n 號機
        update_DbmsData(SPEC_FIREMAN_ONLY,
                        JobMaker_Form.Spec_Fireman_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉
        update_DbmsData(SPEC_PARKING,
                        JobMaker_Form.Spec_Parking_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉-Only CheckBox
        update_DbmsData(SPEC_PARKING_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Parking_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉-Only TextBox
        update_DbmsData(SPEC_PARKING_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Parking_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉-停車階
        update_DbmsData(SPEC_PARKING_FL,
                        JobMaker_Form.Spec_Parking_FL_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉-ELVIC
        update_DbmsData(SPEC_PARKING_ELVIC,
                        JobMaker_Form.Spec_ParkingFL_ELVIC_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉-WTB
        update_DbmsData(SPEC_PARKING_WTB,
                        JobMaker_Form.Spec_ParkingFL_WTB_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉-休止
        update_DbmsData(SPEC_PARKING_SHUTDOWN,
                        JobMaker_Form.Spec_ParkingFL_DR_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉-COB
        update_DbmsData(SPEC_PARKING_COB,
                        JobMaker_Form.Spec_ParkingFL_COB_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '停車階運轉-HALL
        update_DbmsData(SPEC_PARKING_HALL,
                        JobMaker_Form.Spec_ParkingFL_HALL_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉
        update_DbmsData(SPEC_SEISMIC,
                        JobMaker_Form.Spec_Seismic_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉-ONLY CHECKBOX
        update_DbmsData(SPEC_SEISMIC_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Seismic_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉-ONLY TEXTBOX
        update_DbmsData(SPEC_SEISMIC_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Seismic_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉-感知器N段
        update_DbmsData(SPEC_SEISMIC_CANCEL_SW,
                        JobMaker_Form.Spec_SeismicSensor_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉-感知器N段 ONLY CHECKBOX
        update_DbmsData(SPEC_SEISMIC_CANCEL_SW_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_SeismicSensor_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉-感知器N段 ONLY TEXTBOX
        update_DbmsData(SPEC_SEISMIC_CANCEL_SW_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_SeismicSensor_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉-自動解除開關
        update_DbmsData(SPEC_SEISMIC_CANCEL_SW,
                        JobMaker_Form.Spec_SeismicSW_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉-自動解除開關 ONLY CHECKBOX
        update_DbmsData(SPEC_SEISMIC_CANCEL_SW_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_SeismicSW_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '地震管制運轉-自動解除開關 ONLY TEXTBOX
        update_DbmsData(SPEC_SEISMIC_CANCEL_SW_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_SeismicSW_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '車廂管制運轉燈
        update_DbmsData(SPEC_CPI,
                        JobMaker_Form.Spec_CPI_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂管制運轉燈-地震
        update_DbmsData(SPEC_CPI_SEISMIC,
                        JobMaker_Form.Spec_CpiSeismic_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂管制運轉燈-火災
        update_DbmsData(SPEC_CPI_FIRE,
                        JobMaker_Form.Spec_CpiFire_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂管制運轉燈-自發
        update_DbmsData(SPEC_CPI_EMER,
                        JobMaker_Form.Spec_CpiEmer_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂管制運轉燈-緊急
        update_DbmsData(SPEC_CPI_FM,
                        JobMaker_Form.Spec_CpiFM_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂管制運轉燈-滿載
        update_DbmsData(SPEC_CPI_OLT,
                        JobMaker_Form.Spec_CpiOLT_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂管制運轉燈-滿載 ONLY CHECKBOX
        update_DbmsData(SPEC_CPI_OLT_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_CpiOLT_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂管制運轉燈-滿載 ONLY TEXTBOX
        update_DbmsData(SPEC_CPI_OLT_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_CpiOLT_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)


        '車廂上到著鈴
        update_DbmsData(SPEC_CAR_GONG,
                        JobMaker_Form.Spec_CarGong_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [TOP] CHECKBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOP_CHECKBOX,
                        JobMaker_Form.Spec_CarGong_Top_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [TOP] TEXTBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOP,
                        JobMaker_Form.Spec_CarGong_Top_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [TOP] ONLY CHECKBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOP_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_CarGong_Top_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [TOP] ONLY TEXTBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOP_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_CarGong_Top_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)


        '車廂上到著鈴- CAR [TOP BTM] CHECKBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOPBTM_CHECKBOX,
                        JobMaker_Form.Spec_CarGong_TopBtm_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [TOP BTM] TEXTBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOPBTM,
                        JobMaker_Form.Spec_CarGong_TopBtm_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [TOP BTM] ONLY CHECKBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOPBTM_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_CarGong_TopBtm_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [TOP BTM] ONLY TEXTBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOPBTM_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_CarGong_TopBtm_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)


        '車廂上到著鈴- CAR [COB] CHECKBOX
        update_DbmsData(SPEC_CAR_GONG_COB_CHECKBOX,
                        JobMaker_Form.Spec_CarGong_COB_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [COB] TEXTBOX
        update_DbmsData(SPEC_CAR_GONG_CARTOPBTM,
                        JobMaker_Form.Spec_CarGong_COB_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [COB] ONLY CHECKBOX
        update_DbmsData(SPEC_CAR_GONG_COB_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_CarGong_COB_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [COB] ONLY TEXTBOX
        update_DbmsData(SPEC_CAR_GONG_COB_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_CarGong_COB_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)


        '車廂上到著鈴- CAR [VONIC] CHECKBOX
        update_DbmsData(SPEC_CAR_GONG_VONIC_CHECKBOX,
                        JobMaker_Form.Spec_CarGong_VONIC_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [VONIC] TEXTBOX
        update_DbmsData(SPEC_CAR_GONG_VONIC,
                        JobMaker_Form.Spec_CarGong_VONIC_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [VONIC] ONLY CHECKBOX
        update_DbmsData(SPEC_CAR_GONG_VONIC_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_CarGong_VONIC_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂上到著鈴- CAR [VONIC] ONLY TEXTBOX
        update_DbmsData(SPEC_CAR_GONG_VONIC_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_CarGong_VONIC_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '乘場到著鈴
        update_DbmsData(SPEC_HALL_GONG,
                        JobMaker_Form.Spec_HallGong_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場信號文字
        update_DbmsData(SPEC_HPI,
                        JobMaker_Form.Spec_HPIMsg_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場信號文字-滿載
        update_DbmsData(SPEC_HPI_OLT,
                        JobMaker_Form.Spec_HpiOLT_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場信號文字-保養
        update_DbmsData(SPEC_HPI_MAIN,
                        JobMaker_Form.Spec_HpiMain_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場信號文字-專用
        update_DbmsData(SPEC_HPI_INDEP,
                        JobMaker_Form.Spec_HpiIndep_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場信號文字-緊急
        update_DbmsData(SPEC_HPI_EMER,
                        JobMaker_Form.Spec_HpiFM_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門延長按鈕
        update_DbmsData(SPEC_DR_HOLD,
                        JobMaker_Form.Spec_DrHold_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '刷卡機
        update_DbmsData(SPEC_CRD,
                        JobMaker_Form.Spec_CRD_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '刷卡機-分層全層
        update_DbmsData(SPEC_CRD_TYPE,
                        JobMaker_Form.Spec_CRDType_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '刷卡機-ID:4
        update_DbmsData(SPEC_CRD_ID4,
                        JobMaker_Form.Spec_CRDID4_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '刷卡機-ID:5
        update_DbmsData(SPEC_CRD_ID5,
                        JobMaker_Form.Spec_CRDID5_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '自家發
        update_DbmsData(SPEC_EMER,
                        JobMaker_Form.Spec_Emer_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '自家發-訊號
        update_DbmsData(SPEC_EMER_SIGNAL,
                        JobMaker_Form.Spec_EmerSignal_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '自家發-緊急容量
        update_DbmsData(SPEC_EMER_CAPACITY,
                        JobMaker_Form.Spec_EmerCapacity_NumericUpDown.Value,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '自家發-入力點
        update_DbmsData(SPEC_EMER_INPUT,
                        JobMaker_Form.Spec_EmerInput_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '自家發-Address
        update_DbmsData(SPEC_EMER_ADDRESS,
                        JobMaker_Form.Spec_EmerAddress_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '自家發 TabPage中的基本資訊  -------------------------------------------------------------------
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_EmerInfo()

        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.DoubleLayer_Panel,
                                    JobMaker_Form.Spec_EmerNum_NumericUpDown,
                                    dyCtrlName.JobMaker_EmerTBInfoName_Array.Count,
                                    dyCtrlName.JobMaker_EmerTBInfoName_Array,
                                    JobMaker_Form.Spec_emerGroup_TabControl,
                                    SPEC_EMER_NUMBER,
                                    SQLite_tableName_SpecTW)
        '自家發-群數
        update_DbmsData(SPEC_EMER_NUMBER,
                        JobMaker_Form.Spec_EmerNum_NumericUpDown.Value,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '------------------------------------------------------------------- 自家發 TabPage中的基本資訊 


        'LANDIC
        update_DbmsData(SPEC_LANDIC,
                        JobMaker_Form.Spec_Landic_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '基準階賦歸
        update_DbmsData(SPEC_MFL_RETURN,
                        JobMaker_Form.Spec_MFLReturn_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '基準階賦歸-基準階
        update_DbmsData(SPEC_MFL_RETURN_FL,
                        JobMaker_Form.Spec_MFLReturn_FL_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '語音撥放器VONIC
        update_DbmsData(SPEC_VONIC,
                        JobMaker_Form.Spec_Vonic_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '語音撥放器VONIC-標準
        update_DbmsData(SPEC_VONIC_STANDARD,
                        JobMaker_Form.Spec_Vonic_standard_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC
        update_DbmsData(SPEC_ELVIC,
                        JobMaker_Form.Spec_Elvic_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-PARKING OPE
        update_DbmsData(SPEC_ELVIC_1_PARKING,
                        JobMaker_Form.Spec_Elvic_Parking_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELIVC-FLOOR LOCK OUT
        update_DbmsData(SPEC_ELVIC_1_FL_LOCKOUT,
                        JobMaker_Form.Spec_Elvic_FloorLockOut_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-VIP OPE
        update_DbmsData(SPEC_ELVIC_1_VIP,
                        JobMaker_Form.Spec_Elvic_VIP_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-INDEPENDENT OPE
        update_DbmsData(SPEC_ELVIC_1_INDEP,
                        JobMaker_Form.Spec_Elvic_Indep_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-RETURN TO DESIGNATED FLOOR
        update_DbmsData(SPEC_ELVIC_1_RETURN,
                        JobMaker_Form.Spec_Elvic_ReturnFL_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-CHANGE TRAFFIC PATTERN
        update_DbmsData(SPEC_ELVIC_2_TRAFFIC,
                        JobMaker_Form.Spec_Elvic_Traffic_Peak_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-UP PEAK 
        update_DbmsData(SPEC_ELVIC_2_TRAFFIC_UPPEAK,
                        JobMaker_Form.Spec_Elvic_Traffic_UpPeak_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-DOWN PEAK 
        update_DbmsData(SPEC_ELVIC_2_TRAFFIC_DNPEAK,
                        JobMaker_Form.Spec_Elvic_Traffic_DownPeak_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-LUNCH TIME 
        update_DbmsData(SPEC_ELVIC_2_TRAFFIC_LUNCH,
                        JobMaker_Form.Spec_Elvic_Traffic_Lunch_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-CHANGE MAIN FLOOR
        update_DbmsData(SPEC_ELVIC_2_MFL,
                        JobMaker_Form.Spec_Elvic_MainFL_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-ZONING FOR EXPRESS OPE
        update_DbmsData(SPEC_ELVIC_2_ZONING_EXPRESS,
                        JobMaker_Form.Spec_Elvic_Zoning_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-FLOOR LOCK OUT
        update_DbmsData(SPEC_ELVIC_2_FL_LOCKOUT,
                        JobMaker_Form.Spec_Elvic_FloorLockOut_GR_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-CAR CALL DISCONNECT
        update_DbmsData(SPEC_ELVIC_2_CARCALL,
                        JobMaker_Form.Spec_Elvic_CarCall_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-FIRE OPE. COMMAND
        update_DbmsData(SPEC_ELVIC_3_FIRE,
                        JobMaker_Form.Spec_Elvic_Fire_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-WAVIC OPE. COMMAND
        update_DbmsData(SPEC_ELVIC_3_WAVIC,
                        JobMaker_Form.Spec_Elvic_Wavic_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-CARE READER COMMAND
        update_DbmsData(SPEC_ELVIC_3_CARD,
                        JobMaker_Form.Spec_Elvic_CRD_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場廳燈
        update_DbmsData(SPEC_HLL,
                        JobMaker_Form.Spec_HLL_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣
        update_DbmsData(SPEC_WCOB,
                        JobMaker_Form.Spec_WCOB_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣-Only CHECKBOX
        update_DbmsData(SPEC_WCOB_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_WCOB_only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣-Only TEXTBOX
        update_DbmsData(SPEC_WCOB_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_WCOB_only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣-SCOB
        update_DbmsData(SPEC_WSCOB,
                        JobMaker_Form.Spec_WSCOB_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣-SCOB ONLY CHECKBOX
        update_DbmsData(SPEC_WSCOB_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_WSCOB_only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣-SCOB ONLY TEXTBOX
        update_DbmsData(SPEC_WSCOB_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_WSCOB_only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣-鳴動
        update_DbmsData(SPEC_WCOB_RING,
                        JobMaker_Form.Spec_WCOB_Ring_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '運轉手盤運轉
        update_DbmsData(SPEC_ATT,
                        JobMaker_Form.Spec_ATT_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '浸水管制運轉
        update_DbmsData(SPEC_FLOOD,
                        JobMaker_Form.Spec_Flood_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '浸水管制運轉-停止階
        update_DbmsData(SPEC_FLOOD_FL,
                        JobMaker_Form.Spec_Flood_FL_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'LS1M頂部緊急停止開關
        update_DbmsData(SPEC_LS1M,
                        JobMaker_Form.Spec_LS1M_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '電力回升
        update_DbmsData(SPEC_PRU,
                        JobMaker_Form.Spec_PRU_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell
        update_DbmsData(SPEC_LOAD_CELL,
                        JobMaker_Form.Spec_LoadCell_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell-裝置在
        update_DbmsData(SPEC_LOAD_CELL_POSITION,
                        JobMaker_Form.Spec_LoadCellPos_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB
        update_DbmsData(SPEC_WTB,
                        JobMaker_Form.Spec_WTB_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-故障燈
        update_DbmsData(SPEC_WTB_ERROR,
                        JobMaker_Form.Spec_WTB_Error_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-休止燈
        update_DbmsData(SPEC_WTB_STOP,
                        JobMaker_Form.Spec_WTB_Stop_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-消防燈
        update_DbmsData(SPEC_WTB_FIREMAN,
                        JobMaker_Form.Spec_WTB_FM_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-正常燈
        update_DbmsData(SPEC_WTB_NORMAL,
                        JobMaker_Form.Spec_WTB_Normal_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-緊急電源燈
        update_DbmsData(SPEC_WTB_URGENT,
                        JobMaker_Form.Spec_WTB_Urgent_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-火災燈
        update_DbmsData(SPEC_WTB_FO,
                        JobMaker_Form.Spec_WTB_FO_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-自家發燈
        update_DbmsData(SPEC_WTB_EMER,
                        JobMaker_Form.Spec_WTB_EmerPow_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-警示燈
        update_DbmsData(SPEC_WTB_ALART,
                        JobMaker_Form.Spec_WTB_Alart_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-地震燈
        update_DbmsData(SPEC_WTB_EQ,
                        JobMaker_Form.Spec_WTB_EQ_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-專用燈
        update_DbmsData(SPEC_WTB_INDEP,
                        JobMaker_Form.Spec_WTB_Indep_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-地震開關
        update_DbmsData(SPEC_WTB_EQSW,
                        JobMaker_Form.Spec_WTB_EQSW_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-bz解除開關
        update_DbmsData(SPEC_WTB_BZSW,
                        JobMaker_Form.Spec_WTB_BZSW_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-check開關
        update_DbmsData(SPEC_WTB_CHKSW,
                        JobMaker_Form.Spec_WTB_ChkSW_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-停車開關
        update_DbmsData(SPEC_WTB_PKSW,
                        JobMaker_Form.Spec_WTB_PKSW_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-地震指示器
        update_DbmsData(SPEC_WTB_EQIND,
                        JobMaker_Form.Spec_WTB_EQIND_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WTB-地震強度
        update_DbmsData(SPEC_WTB_EQMAC,
                        JobMaker_Form.Spec_WTB_EQMac_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '正背門
        update_DbmsData(SPEC_FRONT_REAR_DR,
                        JobMaker_Form.Spec_FrontRearDr_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '各停開關
        update_DbmsData(SPEC_EACH_STOP,
                        JobMaker_Form.Spec_EachStop_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '拒付運轉
        update_DbmsData(SPEC_INSTALL_OPE,
                        JobMaker_Form.Spec_install_ope_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'vonic蜂鳴器
        update_DbmsData(SPEC_VONICBZ,
                        JobMaker_Form.Spec_VonicBz_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '單群控切換
        update_DbmsData(SPEC_OPE_SW,
                        JobMaker_Form.Spec_OpeSw_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '單群控切換-入力點Position
        update_DbmsData(SPEC_OPE_SW_POS,
                        JobMaker_Form.Spec_OpeSw_InputPos_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '單群控切換-入力點Address
        update_DbmsData(SPEC_OPE_SW_ADDRESS,
                        JobMaker_Form.Spec_OpeSw_InputAddress_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
    End Sub
    Private Sub Important_Stored()
        'IDU CheckBox
        update_DbmsData(IMPORTANT_Use_ChkBox,
                        JobMaker_Form.Use_Imp_CheckBox.Checked,
                        SQLite_tableName_Important,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '風扇連動
        'update_DbmsData(IMPORTANT_FAN,
        '                JobMaker_Form.Imp_FAN_ComboBox.Text,
        '                SQLite_tableName_Important,
        '                SQLite_connectionPath_Job,
        '                SQLite_JobDBMS_Name)
        'OVER BALANCE
        update_DbmsData(IMPORTANT_BALANCE,
                        JobMaker_Form.Imp_OverBalance_ComboBox.Text,
                        SQLite_tableName_Important,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'WCOB
        update_DbmsData(IMPORTANT_WCOB,
                        JobMaker_Form.Imp_WHB_ComboBox.Text,
                        SQLite_tableName_Important,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'DOOR TYPE
        update_DbmsData(IMPORTANT_DOOR,
                        JobMaker_Form.Imp_DoorType_TextBox.Text,
                        SQLite_tableName_Important,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
    End Sub
    Private Sub MMIC_Stored()
        'MMIC CheckBox
        update_DbmsData(MMIC_Use_ChkBox,
                        JobMaker_Form.Use_mmic_CheckBox.Checked,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC MR BASE
        update_DbmsData(MMIC_MR_BASE,
                        JobMaker_Form.MMIC_MR_Base_TextBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC MR CP43
        update_DbmsData(MMIC_MR_CP43x,
                        JobMaker_Form.MMIC_MR_CP43x_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'MMIC MR EEPROM BASE
        update_DbmsData(MMIC_MR_EBase,
                        JobMaker_Form.MMIC_MR_EBase_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC MR EEPROM Car Obj
        update_DbmsData(MMIC_MR_ECarObj,
                        JobMaker_Form.MMIC_MR_ECarObj_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'MMIC SV TYPE
        update_DbmsData(MMIC_SV_TYPE,
                        JobMaker_Form.MMIC_SV_Type_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC SV BASE
        update_DbmsData(MMIC_SV_BASE,
                        JobMaker_Form.MMIC_SV_Base_TextBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'MMIC SV EEPROM BASE
        update_DbmsData(MMIC_SV_EBase,
                        JobMaker_Form.MMIC_SV_EBase_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC SV EEPROM Car Obj
        update_DbmsData(MMIC_SV_ECarObj,
                        JobMaker_Form.MMIC_SV_ECarObj_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC SV EEPROM Number
        update_DbmsData(MMIC_SV_ENumber,
                        JobMaker_Form.MMIC_SV_E_NumericUpDown.Value,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC VD10 ROM DEVICE
        update_DbmsData(MMIC_VD10_ROM,
                        JobMaker_Form.MMIC_VD10_ROM_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC VD10 Quantity
        update_DbmsData(MMIC_VD10_Quantity,
                        JobMaker_Form.MMIC_VD10_Quantity_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC VD10 Type
        update_DbmsData(MMIC_VD10_TYPE,
                        JobMaker_Form.MMIC_VD10_Type_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC VD10 BASE
        update_DbmsData(MMIC_VD10_BASE,
                        JobMaker_Form.MMIC_VD10_Base_TextBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)




        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_MMICInfo()
        'MMIC MR CarNo / ObjName / ObjName Base
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.MMIC_MR_NumericUpDown,
                                    dyCtrlName.JobMaker_MMIC_MrBase_InfoName_Array.Count,
                                    dyCtrlName.JobMaker_MMIC_MrBase_InfoName_Array,
                                    JobMaker_Form.MMIC_MR_Panel,
                                    MMIC_MR_Number,
                                    SQLite_tableName_MMIC)
        'MMIC MR Number
        update_DbmsData(MMIC_MR_Number,
                        JobMaker_Form.MMIC_MR_NumericUpDown.Value,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)


        'MMIC MR EEPROM CarNo / ObjName
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.MMIC_MR_E_NumericUpDown,
                                    dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array.Count,
                                    dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array,
                                    JobMaker_Form.MMIC_MR_E_Panel,
                                    MMIC_MR_ENumber,
                                    SQLite_tableName_MMIC)
        'MMIC MR EEPROM Number
        update_DbmsData(MMIC_MR_ENumber,
                        JobMaker_Form.MMIC_MR_E_NumericUpDown.Value,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)


        'MMIC SV CarNo / ObjName / ObjName Base
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.MMIC_SV_NumericUpDown,
                                    dyCtrlName.JobMaker_MMIC_SvBase_InfoName_Array.Count,
                                    dyCtrlName.JobMaker_MMIC_SvBase_InfoName_Array,
                                    JobMaker_Form.MMIC_SV_Panel,
                                    MMIC_SV_Number,
                                    SQLite_tableName_MMIC)
        'MMIC SV Number
        update_DbmsData(MMIC_SV_Number,
                        JobMaker_Form.MMIC_SV_NumericUpDown.Value,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)


        'MMIC SV EEPROM CarNo / ObjName / ObjName Base
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.MMIC_SV_E_NumericUpDown,
                                    dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array.Count,
                                    dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array,
                                    JobMaker_Form.MMIC_SV_E_Panel,
                                    MMIC_SV_ENumber,
                                    SQLite_tableName_MMIC)
        'MMIC SV EEPROM Number
        update_DbmsData(MMIC_SV_ENumber,
                        JobMaker_Form.MMIC_SV_E_NumericUpDown.Value,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)



        'MMIC VD10 CarNo  / ObjName
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.MMIC_VD10_NumericUpDown,
                                    dyCtrlName.JobMaker_MMIC_VD10Base_InfoName_Array.Count,
                                    dyCtrlName.JobMaker_MMIC_VD10Base_InfoName_Array,
                                    JobMaker_Form.MMIC_VD10_Panel,
                                    MMIC_VD10_Number,
                                    SQLite_tableName_MMIC)
        'MMIC VD10 Number
        update_DbmsData(MMIC_VD10_Number,
                        JobMaker_Form.MMIC_VD10_NumericUpDown.Value,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
    End Sub
    '------------------------------------------ 更新或新建SQLite 

    '載入SQLite ------------------------------------------
    Public Sub Load_Stored(job_dbms As String)
        Try
            SQLite_JobDBMS_Name = job_dbms

            Dim Basic_ChkBox, CheckList_ChkBox, Program_ChkBox,
                DWG_ChkBox, Spec_ChkBox, Imp_ChkBox, MMIC_ChkBox As String

            Basic_ChkBox =
                read_DbmsData(Basic_Use_ChkBox,
                              SQLite_tableName_Basic,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If Basic_ChkBox = "True" Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"讀取 「基本」 開始 ======================= {vbCrLf}"
                Basic_Load()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================讀取 「基本」 結束 {vbCrLf}"
            End If
            '---------------------------------------------
            CheckList_ChkBox =
                read_DbmsData(ChkList_Use_ChkBox,
                              SQLite_tableName_CheckList,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If CheckList_ChkBox = "True" Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"讀取 「CheckList」 開始 ======================= {vbCrLf}"
                CheckList_Load()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================讀取 「CheckList」 結束 {vbCrLf}"
            End If
            '---------------------------------------------
            Program_ChkBox =
                read_DbmsData(ChkList_Prgm_Use_ChkBox,
                              SQLite_tableName_Program,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If Program_ChkBox = "True" Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"讀取 「程式變更」 開始 ======================= {vbCrLf}"
                ProgramChange_Load()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================讀取 「程式變更」 結束 {vbCrLf}"
            End If
            '---------------------------------------------
            'DWG_ChkBox =
            '    read_DbmsData(DWG_Use_ChkBox,
            '                  SQLite_tableName_DWG,
            '                  SQLite_connectionPath_Job,
            '                  SQLite_JobDBMS_Name)
            'If DWG_ChkBox = "True" Then
            '    JobMaker_Form.ResultOutput_TextBox.Text += $"讀取 「送狀」 開始 ======================= {vbCrLf}"
            '    DWG_Load()
            '    JobMaker_Form.ResultOutput_TextBox.Text += $"=======================讀取 「送狀」 結束 {vbCrLf}"
            'End If
            '---------------------------------------------
            Spec_ChkBox =
                read_DbmsData(SpecBasic_Use_ChkBox,
                              SQLite_tableName_SpecBasic,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If Spec_ChkBox = "True" Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"讀取 「仕樣」 開始 ======================= {vbCrLf}"
                SpecBasic_Load()
                SpecTW_Load()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================讀取 「仕樣」 結束 {vbCrLf}"
            End If
            '---------------------------------------------
            Imp_ChkBox =
                read_DbmsData(IMPORTANT_Use_ChkBox,
                              SQLite_tableName_Important,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If Imp_ChkBox = "True" Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"讀取 「重要設定」 開始 ======================= {vbCrLf}"
                Important_Load()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================讀取 「重要設定」 結束 {vbCrLf}"
            End If
            '---------------------------------------------
            MMIC_ChkBox =
                read_DbmsData(MMIC_Use_ChkBox,
                              SQLite_tableName_MMIC,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If MMIC_ChkBox = "True" Then
                JobMaker_Form.ResultOutput_TextBox.Text += $"讀取 「MMIC」 開始 ======================= {vbCrLf}"
                MMIC_Load()
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================讀取 「MMIC」 結束 {vbCrLf}"
            End If

            MsgBox("載入成功",, "Fine")
        Catch e As Exception
            MsgBox($"載入失敗 : {e}",, "Fail")
        End Try
    End Sub
    Private Sub Basic_Load()
        'Basic Use CheckBox
        Dim temp_basic_use_chkbox As String
        temp_basic_use_chkbox =
            read_DbmsData(Basic_Use_ChkBox,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        If temp_basic_use_chkbox <> "" Then
            JobMaker_Form.Use_Basic_CheckBox.Checked = temp_basic_use_chkbox
        End If

        'Local
        JobMaker_Form.Basic_Local_ComboBox.Text =
            read_DbmsData(Basic_Local,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        'JobNo(New)
        JobMaker_Form.Basic_JobNoNew_TextBox.Text =
            read_DbmsData(Basic_JobNo_New,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        'JobNo(Old)
        JobMaker_Form.Basic_JobNoOld_TextBox.Text =
            read_DbmsData(Basic_JobNo_Old,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'JobNo(Mod)
        JobMaker_Form.Basic_JobNoMOD_TextBox.Text =
            read_DbmsData(Basic_JobNo_Mod,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'JobName
        JobMaker_Form.Basic_JobName_TextBox.Text =
            read_DbmsData(Basic_JobName,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'DesignerChinese
        JobMaker_Form.Basic_DesingerChinese_ComboBox.Text =
            read_DbmsData(Basic_DesignerChinese,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'DesignerEnglish
        JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text =
            read_DbmsData(Basic_DesignerEnglish,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'CheckerChinese
        JobMaker_Form.Basic_CheckerChinese_ComboBox.Text =
            read_DbmsData(Basic_CheckerChinese,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'CheckerEnglish
        JobMaker_Form.Basic_CheckerEnglish_ComboBox.Text =
            read_DbmsData(Basic_CheckerEnglish,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'Date Time Picker
        JobMaker_Form.Basic_DrawDate_DateTimePicker.Value =
            DateTime.Parse(read_DbmsData(Basic_DateTimePicker,
                                         SQLite_tableName_Basic,
                                         SQLite_connectionPath_Job,
                                         SQLite_JobDBMS_Name))
    End Sub
    Private Sub CheckList_Load()
        'CheckList Use Checkbox
        JobMaker_Form.Use_ChkList_CheckBox.Checked =
            read_DbmsData(ChkList_Use_ChkBox,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        'PA DateTimePicker
        JobMaker_Form.ChkList_PaSheet_DateTimePicker.Value =
            read_DbmsData(ChkList_PA_DateTimePicker,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        'OS DateTimePicker
        JobMaker_Form.ChkList_OS_DateTimePicker.Value =
            read_DbmsData(ChkList_OS_DateTimePicker,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        'CFM DateTimePicker
        JobMaker_Form.ChkList_Confirm_DateTimePicker.Value =
            read_DbmsData(ChkList_CFM_DateTimePicker,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'ELE DateTimePicker
        JobMaker_Form.ChkList_Elec_DateTimePicker.Value =
            read_DbmsData(ChkList_ELE_DateTimePicker,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'PA CheckBox
        JobMaker_Form.ChkList_PaSheet_CheckBox.Checked =
            read_DbmsData(ChkList_PA_ChkBox,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'OS CheckBox
        JobMaker_Form.ChkList_OS_CheckBox.Checked =
            read_DbmsData(ChkList_OS_ChkBox,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'CFM CheckBox
        JobMaker_Form.ChkList_Confirm_CheckBox.Checked =
            read_DbmsData(ChkList_CFM_ChkBox,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        'ELE CheckBox
        JobMaker_Form.ChkList_Elec_CheckBox.Checked =
            read_DbmsData(ChkList_ELE_ChkBox,
                          SQLite_tableName_CheckList,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        'Q1 no RadioButton
        JobMaker_Form.ChkList_1_no_RadioButton.Checked =
           read_DbmsData(ChkList_Q1No_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q1 yes RadioButton
        JobMaker_Form.ChkList_1_yes_RadioButton.Checked =
           read_DbmsData(ChkList_Q1Yes_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q1 yes 討論結果
        JobMaker_Form.ChkList_1_yes_Content_TextBox.Text =
           read_DbmsData(ChkList_Q1Yes_Content,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q1 yes 結果
        JobMaker_Form.ChkList_1_yes_result_TextBox.Text =
           read_DbmsData(ChkList_Q1Yes_Result,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q2 no RadioButton
        JobMaker_Form.ChkList_2_no_RadioButton.Checked =
           read_DbmsData(ChkList_Q2No_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q2 yes RadioButton
        JobMaker_Form.ChkList_2_yes_RadioButton.Checked =
           read_DbmsData(ChkList_Q2Yes_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q2 yes 指出內容
        JobMaker_Form.ChkList_2_yes_Content_TextBox.Text =
           read_DbmsData(ChkList_Q2Yes_Content,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q2 yes 結果
        JobMaker_Form.ChkList_2_yes_Result_TextBox.Text =
           read_DbmsData(ChkList_Q2Yes_Result,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'Q3 no RadioButton
        JobMaker_Form.ChkList_3_no_RadioButton.Checked =
           read_DbmsData(ChkList_Q3No_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q3 yes RadioButton
        JobMaker_Form.ChkList_3_yes_RadioButton.Checked =
           read_DbmsData(ChkList_Q3Yes_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q3 yes 討論者
        JobMaker_Form.ChkList_3_yes_Man_TextBox.Text =
           read_DbmsData(ChkList_Q3Yes_Man,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q3 yes 內容
        JobMaker_Form.ChkList_3_yes_Content_TextBox.Text =
           read_DbmsData(ChkList_Q3Yes_Content,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q3 yes 結論
        JobMaker_Form.ChkList_3_yes_Result_TextBox.Text =
           read_DbmsData(ChkList_Q3Yes_Result,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q4 MMIC 
        JobMaker_Form.ChkList_4_ObjName_TextBox.Text =
           read_DbmsData(ChkList_Q4MMIC,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q4 SV
        JobMaker_Form.ChkList_4_SV_TextBox.Text =
           read_DbmsData(ChkList_Q4SV,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q5 無 RadioButton
        JobMaker_Form.ChkList_5_no_RadioButton.Checked =
           read_DbmsData(ChkList_Q5No_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q5 標準 RadioButton
        JobMaker_Form.ChkList_5_std_RadioButton.Checked =
           read_DbmsData(ChkList_Q5Std_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q5 工直 RadioButton
        JobMaker_Form.ChkList_5_nstd_RadioButton.Checked =
           read_DbmsData(ChkList_Q5NoStd_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q5 標準 TextBox
        JobMaker_Form.ChkList_5_std_Content_TextBox.Text =
           read_DbmsData(ChkList_Q5Std_Content,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q5 工直 TextBox
        JobMaker_Form.ChkList_5_nstd_Content_TextBox.Text =
           read_DbmsData(ChkList_Q5nStd_Content,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q6 no RadioButton
        JobMaker_Form.ChkList_6_no_RadioButton.Checked =
           read_DbmsData(ChkList_Q6No_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q6 yes RadioButton
        JobMaker_Form.ChkList_6_yes_RadioButton.Checked =
           read_DbmsData(ChkList_Q6Yes_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q6 yes Check RadioButton
        JobMaker_Form.ChkList_6_yesChk_RadioButton.Checked =
           read_DbmsData(ChkList_Q6YesChk_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q6 yes Item RadioButton
        JobMaker_Form.ChkList_6_yesItem_RadioButton.Checked =
           read_DbmsData(ChkList_Q6YesItem_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q6 yes 檢驗目標 TextBox
        JobMaker_Form.ChkList_6_yes_Content_TextBox.Text =
           read_DbmsData(ChkList_Q6Yes_Content,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q7 no RadioButton
        JobMaker_Form.ChkList_7_no_RadioButton.Checked =
           read_DbmsData(ChkList_Q7No_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q7 yes RadioButton
        JobMaker_Form.ChkList_7_yes_RadioButton.Checked =
           read_DbmsData(ChkList_Q7Yes_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q7 yes 文書 Textbox
        JobMaker_Form.ChkList_7_yes_RadioButton.Checked =
           read_DbmsData(ChkList_Q7Yes_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q8 no RadioButton
        JobMaker_Form.ChkList_8_no_RadioButton.Checked =
           read_DbmsData(ChkList_Q8No_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q8 yes RadioButton
        JobMaker_Form.ChkList_8_yes_RadioButton.Checked =
           read_DbmsData(ChkList_Q8Yes_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q8 item RadioButton
        JobMaker_Form.ChkList_8Item_RadioButton.Checked =
           read_DbmsData(ChkList_Q8Item_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q9 no RadioButton
        JobMaker_Form.ChkList_9_no_RadioButton.Checked =
           read_DbmsData(ChkList_Q9No_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q9 yes RadioButton
        JobMaker_Form.ChkList_9_yes_RadioButton.Checked =
           read_DbmsData(ChkList_Q9Yes_RadioBox,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
    End Sub
    Private Sub ProgramChange_Load()

        'Program Change Use CheckBox
        JobMaker_Form.Use_Program_CheckBox.Checked =
           read_DbmsData(ChkList_Use_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '1. 變更理由 TextBox
        JobMaker_Form.PrmList_1_reason_TextBox.Text =
           read_DbmsData(ChkList_Prgm_1_reason,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.測試裝置CheckBox
        JobMaker_Form.PrmList_2_test_CheckBox.Checked =
           read_DbmsData(ChkList_Prgm_2_Test_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.控制盤CheckBox
        JobMaker_Form.PrmList_2_COP_CheckBox.Checked =
           read_DbmsData(ChkList_Prgm_2_COP_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.研修測試塔CheckBox
        JobMaker_Form.PrmList_2_Tower_CheckBox.Checked =
           read_DbmsData(ChkList_Prgm_2_Tower_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.其他CheckBox
        JobMaker_Form.PrmList_2_Other_CheckBox.Checked =
           read_DbmsData(ChkList_Prgm_2_Other_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.測試裝置TextBox
        JobMaker_Form.PrmList_2_test_TextBox.Text =
           read_DbmsData(ChkList_Prgm_2_Test_Content,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.控制盤TextBox
        JobMaker_Form.PrmList_2_COP_TextBox.Text =
           read_DbmsData(ChkList_Prgm_2_COP_Content,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.研修測試塔TextBox
        JobMaker_Form.PrmList_2_tower_TextBox.Text =
           read_DbmsData(ChkList_Prgm_2_Tower_Content,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.其他TextBox
        JobMaker_Form.PrmList_2_other_TextBox.Text =
           read_DbmsData(ChkList_Prgm_2_Other_Content,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '3.Debug CheckBox
        JobMaker_Form.PrmList_3_debug_CheckBox.Checked =
           read_DbmsData(ChkList_Prgm_3_Debug_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '3.Test CheckBox
        JobMaker_Form.PrmList_3_test_CheckBox.Checked =
           read_DbmsData(ChkList_Prgm_3_Test_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '3.Confrim CheckBox
        JobMaker_Form.PrmList_3_confirm_CheckBox.Checked =
           read_DbmsData(ChkList_Prgm_3_CFM_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '3.Execution CheckBox
        JobMaker_Form.PrmList_3_excute_CheckBox.Checked =
           read_DbmsData(ChkList_Prgm_3_EXE_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '3.Other CheckBox
        JobMaker_Form.PrmList_3_other_Checkbox.Checked =
           read_DbmsData(ChkList_Prgm_3_Other_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '3.Other TextBox
        JobMaker_Form.PrmList_3_other_TextBox.Text =
           read_DbmsData(ChkList_Prgm_3_OtherContent,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.1 Auto Yes RadioBtn
        JobMaker_Form.PrmList_4_yes1_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_1Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.1 Auto No RadioBtn
        JobMaker_Form.PrmList_4_no1_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_1No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.2 Input Yes RadioBtn
        JobMaker_Form.PrmList_4_yes2_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_2Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.2 Input No RadioBtn
        JobMaker_Form.PrmList_4_no2_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_2No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.3 Ini Yes RadioBtn
        JobMaker_Form.PrmList_4_yes3_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_3Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.3 Ini No RadioBtn
        JobMaker_Form.PrmList_4_no3_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_3No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.4 Case Yes RadioBtn
        JobMaker_Form.PrmList_4_yes4_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_4Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.4 Case No RadioBtn
        JobMaker_Form.PrmList_4_no4_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_4No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.5 If Yes RadioBtn
        JobMaker_Form.PrmList_4_yes5_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_5Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.5 If No RadioBtn
        JobMaker_Form.PrmList_4_no5_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_5No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.6 Loop Yes RadioBtn
        JobMaker_Form.PrmList_4_yes6_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_6Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.6 Loop No RadioBtn
        JobMaker_Form.PrmList_4_no6_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_6No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.7 Range Yes RadioBtn
        JobMaker_Form.PrmList_4_yes7_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_7Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.7 Range No RadioBtn
        JobMaker_Form.PrmList_4_no7_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_7No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.8 Casting Yes RadioBtn
        JobMaker_Form.PrmList_4_yes8_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_8Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.8 Casting No RadioBtn
        JobMaker_Form.PrmList_4_no8_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_8No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.9 0 Yes RadioBtn
        JobMaker_Form.PrmList_4_yes9_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_9Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.9 0 No RadioBtn
        JobMaker_Form.PrmList_4_no9_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_9No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.10 Count Yes RadioBtn
        JobMaker_Form.PrmList_4_yes10_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_10Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.10 Count No RadioBtn
        JobMaker_Form.PrmList_4_no10_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_10No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.11 Address Yes RadioBtn
        JobMaker_Form.PrmList_4_yes11_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_11Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.11 Address No RadioBtn
        JobMaker_Form.PrmList_4_no11_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_11No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.12 Custom Yes RadioBtn
        JobMaker_Form.PrmList_4_yes12_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_12Yes_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.12 Custom No RadioBtn
        JobMaker_Form.PrmList_4_no12_RadioButton.Checked =
           read_DbmsData(ChkList_Prgm_4_12No_ChkBox,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4 Content RadioBtn
        JobMaker_Form.PrmList_4_content12_TextBox.Text =
           read_DbmsData(ChkList_Prgm_4_TestContent,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
    End Sub
    Private Sub DWG_Load()
        'DWG Use CheckBox
        JobMaker_Form.Use_prk_CheckBox.Checked =
           read_DbmsData(DWG_Use_ChkBox,
                         SQLite_tableName_DWG,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'Vonic標準
        JobMaker_Form.DWG_VonicStd_ComboBox.Text =
           read_DbmsData(DWG_Vonic,
                         SQLite_tableName_DWG,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '輸出項目必要打V
        read_DbmsData_catalogPage(DWG_Pages,
                                  SQLite_tableName_DWG,
                                  JobMaker_Form.DWG_Page_CheckedListBox,
                                  SQLite_connectionPath_Job,
                                  SQLite_JobDBMS_Name)
        '工務/現場/製造
        Dim dwg_page_count As Integer
        dwg_page_count = JobMaker_Form.DWG_Page_CheckedListBox.Items.Count

        With JobMaker_Form
            For i As Integer = 1 To dwg_page_count
                .DWG_Construction_CheckedListBox.
                    Items.Add("")
                .DWG_Produce_CheckedListBox.
                    Items.Add("")
            Next
        End With
    End Sub
    Private Sub SpecBasic_Load()
        Dim dyCtrlName As DynamicControlName = New DynamicControlName

        'Spec Use CheckBox
        Dim temp_specBasic_use_chkbox As String
        temp_specBasic_use_chkbox =
            read_DbmsData(SpecBasic_Use_ChkBox,
                          SQLite_tableName_SpecBasic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        If temp_specBasic_use_chkbox <> "" Then
            JobMaker_Form.Use_SpecBasic_CheckBox.Checked = temp_specBasic_use_chkbox
        End If
        '電梯總數 Textbox
        JobMaker_Form.Spec_LiftNum_NumericUpDown.Value =
           read_DbmsData(SpecBasic_LiftNumber,
                         SQLite_tableName_SpecBasic,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        dyCtrlName.JobMaker_LiftInfo()
        dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_LiftNum_NumericUpDown,
                                  JobMaker_Form.SpecBasic_LiftItem_Dynamic_Panel,
                                  dyCtrlName.JobMaker_LiftInfoName_Array.Count,
                                  dyCtrlName.JobMaker_LiftInfoName_Array,
                                  SQLite_tableName_SpecBasic)

        '機種 / 控制方式
        JobMaker_Form.Spec_MachineType_NumericUpDown.Value =
           read_DbmsData(SpecBasic_MachineType_Number,
                         SQLite_tableName_SpecBasic,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_MachineType_NumericUpDown,
                                  JobMaker_Form.Spec_MachineType_Panel,
                                  {dyCtrlName.Spec_MachineType_ComboBox, dyCtrlName.Spec_ControlWay_ComboBox}.Count,
                                  {dyCtrlName.Spec_MachineType_ComboBox, dyCtrlName.Spec_ControlWay_ComboBox},
                                  SQLite_tableName_SpecBasic)

        '用途 Textbox
        JobMaker_Form.Spec_Purpose_NumericUpDown.Value =
           read_DbmsData(SpecBasic_Purpose_Number,
                         SQLite_tableName_SpecBasic,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_Purpose_NumericUpDown,
                                  JobMaker_Form.Spec_Purpose_Panel,
                                  {dyCtrlName.Spec_Purpose_ComboBox}.Count,
                                  {dyCtrlName.Spec_Purpose_ComboBox},
                                  SQLite_tableName_SpecBasic)



        'panel中的號機基本資訊 -------------------------------------------------------------------



        'If JobMaker_Form.Spec_LiftNum_NumericUpDown.Value <> 0 Then
        '    For lift_i As Integer = 1 To CInt(JobMaker_Form.Spec_LiftNum_NumericUpDown.Value)
        '        For Each tempCtrl As Control In JobMaker_Form.LiftNum_Panel.Controls
        '            '共有幾台號機
        '            '八組TextBox
        '            For lift_j As Integer = 1 To dyCtrlName.JobMaker_LiftInfoName_Array.Count
        '                If tempCtrl.Name = $"{dyCtrlName.JobMaker_LiftInfoName_Array(lift_j - 1)}_{lift_i}" Then
        '                    tempCtrl.Text =
        '                        read_DbmsData_RowID(dyCtrlName.JobMaker_LiftInfoName_Array(lift_j - 1),
        '                                            SQLite_tableName_SpecBasic,
        '                                            SQLite_connectionPath_Job,
        '                                            SQLite_JobDBMS_Name,
        '                                            lift_i)
        '                End If
        '                'Next
        '            Next
        '        Next
        '    Next
        'End If
        '------------------------------------------------------------------- panel中的號機基本資訊 
    End Sub
    Private Sub SpecTW_Load()
        'Spec TW Use CheckBox
        Dim temp_spec_tw_idu_chkbox As String
        temp_spec_tw_idu_chkbox =
            read_DbmsData(SPEC_TW_IDU_CHKBOX,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        If temp_spec_tw_idu_chkbox <> "" Then
            JobMaker_Form.Use_SpecTWIDU_CheckBox.Checked = temp_spec_tw_idu_chkbox
        End If

        Dim temp_spec_tw_fp17_chkbox As String
        temp_spec_tw_fp17_chkbox =
            read_DbmsData(SPEC_TW_FP17_CHKBOX,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        If temp_spec_tw_fp17_chkbox <> "" Then
            JobMaker_Form.Use_SpecTWFP17_CheckBox.Checked = temp_spec_tw_fp17_chkbox
        End If

        '機種
        JobMaker_Form.Spec_Base_ComboBox.Text =
           read_DbmsData(SPEC_MACHINE_TYPE,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '開門時限自動調節
        JobMaker_Form.Spec_DRAuto_ComboBox.Text =
           read_DbmsData(SPEC_AUTO_DR,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '開門時限自動調節-光電裝置
        JobMaker_Form.Spec_PhotoEye_ComboBox.Text =
           read_DbmsData(SPEC_AUTO_DR_PHOTOEYE,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '開門時限自動調節-機械式裝置
        JobMaker_Form.Spec_MechSafety_ComboBox.Text =
           read_DbmsData(SPEC_AUTO_DR_SAFETY,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '取消嬉戲呼叫
        JobMaker_Form.Spec_CancellCall_ComboBox.Text =
           read_DbmsData(SPEC_CANCELL_CALL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '取消嬉戲呼叫-副COB
        JobMaker_Form.Spec_SCOB_ComboBox.Text =
           read_DbmsData(SPEC_CANCELL_CALL_SCOB,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '逆呼無效
        JobMaker_Form.Spec_CancellBehind_ComboBox.Text =
           read_DbmsData(SPEC_CANCELL_BEHIND,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '燈點檢模式
        JobMaker_Form.Spec_LampChk_ComboBox.Text =
           read_DbmsData(SPEC_LAMP_CHK,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '風扇連動
        JobMaker_Form.Spec_AutoFan_ComboBox.Text =
           read_DbmsData(SPEC_AUTO_FAN,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '風扇連動-離子除菌
        JobMaker_Form.Spec_ION_ComboBox.Text =
           read_DbmsData(SPEC_AUTO_FAN_ION_WITHOUT,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '車廂呼叫取消
        JobMaker_Form.Spec_CCCancell_ComboBox.Text =
           read_DbmsData(SPEC_CC_CANCEL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '自動滿員通過
        JobMaker_Form.Spec_AutoPass_ComboBox.Text =
           read_DbmsData(SPEC_AUTO_PASS,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '操作方式
        'JobMaker_Form.Spec_Operation_ComboBox.Text =
        '   read_DbmsData(SPEC_OPERATION,
        '                 SQLite_tableName_SpecTW,
        '                 SQLite_connectionPath_Job,
        '                 SQLite_JobDBMS_Name)
        '專用運轉
        JobMaker_Form.Spec_Indep_ComboBox.Text =
           read_DbmsData(SPEC_INDEP_OPE,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '戶開行走保護
        JobMaker_Form.Spec_UCMP_ComboBox.Text =
           read_DbmsData(SPEC_UCMP,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'HIN/CPI
        JobMaker_Form.Spec_HinCpi_ComboBox.Text =
           read_DbmsData(SPEC_HIN_CPI,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '火災管制運轉
        JobMaker_Form.Spec_Fire_ComboBox.Text =
           read_DbmsData(SPEC_FIRE_OPE,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '火災管制運轉-訊號
        JobMaker_Form.Spec_FireSignal_ComboBox.Text =
           read_DbmsData(SPEC_FIRE_OPE_SIGNAL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '火災管制運轉-Only CheckBox
        Dim spec_fire_only_checkbox_state As String
        spec_fire_only_checkbox_state =
            read_DbmsData(SPEC_FIRE_ONLY_CHECKBOX,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If spec_fire_only_checkbox_state <> "" Then
            JobMaker_Form.Spec_Fire_Only_CheckBox.Checked = CBool(spec_fire_only_checkbox_state)
        End If
        '火災管制運轉-Only TextBox
        JobMaker_Form.Spec_Fire_Only_TextBox.Text =
           read_DbmsData(SPEC_FIRE_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '消防梯運轉
        JobMaker_Form.Spec_Fireman_ComboBox.Text =
           read_DbmsData(SPEC_FIREMAN,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '消防梯運轉-避難階
        JobMaker_Form.Spec_EscapeFL_TextBox.Text =
           read_DbmsData(SPEC_FIREMAN_ESCAPE_FL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '消防梯運轉-Only n 號機 CheckBox
        Dim Spec_Fireman_Only_CheckBox_state As String
        Spec_Fireman_Only_CheckBox_state =
            read_DbmsData(SPEC_FIREMAN_ONLY_CHKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        If Spec_Fireman_Only_CheckBox_state <> "" Then
            JobMaker_Form.Spec_Fireman_Only_CheckBox.Checked = CBool(Spec_Fireman_Only_CheckBox_state)
        End If

        '消防梯運轉-Only n 號機
        JobMaker_Form.Spec_Fireman_Only_TextBox.Text =
           read_DbmsData(SPEC_FIREMAN_ONLY,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '停車階運轉
        JobMaker_Form.Spec_Parking_ComboBox.Text =
           read_DbmsData(SPEC_PARKING,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '停車階運轉-Only CheckBox
        Dim spec_parking_only_checkbox_state As String
        spec_parking_only_checkbox_state =
            read_DbmsData(SPEC_PARKING_ONLY_CHECKBOX,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If spec_parking_only_checkbox_state <> "" Then
            JobMaker_Form.Spec_Parking_Only_CheckBox.Checked = CBool(spec_parking_only_checkbox_state)
        End If
        '停車階運轉-Only TextBox
        JobMaker_Form.Spec_Parking_Only_TextBox.Text =
           read_DbmsData(SPEC_PARKING_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '停車階運轉-停車階
        JobMaker_Form.Spec_Parking_FL_TextBox.Text =
           read_DbmsData(SPEC_PARKING_FL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '停車階運轉-ELVIC
        JobMaker_Form.Spec_ParkingFL_ELVIC_ComboBox.Text =
           read_DbmsData(SPEC_PARKING_ELVIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '停車階運轉-WTB
        JobMaker_Form.Spec_ParkingFL_WTB_ComboBox.Text =
           read_DbmsData(SPEC_PARKING_WTB,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '停車階運轉-休止
        JobMaker_Form.Spec_ParkingFL_DR_ComboBox.Text =
           read_DbmsData(SPEC_PARKING_SHUTDOWN,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '停車階運轉-COB
        JobMaker_Form.Spec_ParkingFL_COB_ComboBox.Text =
           read_DbmsData(SPEC_PARKING_COB,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '停車階運轉-HALL
        JobMaker_Form.Spec_ParkingFL_HALL_ComboBox.Text =
           read_DbmsData(SPEC_PARKING_HALL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '地震管制運轉
        JobMaker_Form.Spec_Seismic_ComboBox.Text =
           read_DbmsData(SPEC_SEISMIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '地震管制運轉 ONLY CHECKBOX
        Dim spec_seismic_only_checkBox_state As String
        spec_seismic_only_checkBox_state =
            read_DbmsData(SPEC_SEISMIC_ONLY_CHECKBOX,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If spec_seismic_only_checkBox_state <> "" Then
            JobMaker_Form.Spec_Seismic_Only_CheckBox.Checked = CBool(spec_seismic_only_checkBox_state)
        End If
        '地震管制運轉 ONLY TEXTBOX
        JobMaker_Form.Spec_Seismic_Only_TextBox.Text =
           read_DbmsData(SPEC_SEISMIC_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '地震管制運轉-感知器N段
        JobMaker_Form.Spec_SeismicSensor_ComboBox.Text =
           read_DbmsData(SPEC_SEISMIC_CANCEL_SW,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '地震管制運轉-感知器N段 ONLY CHECKBOX
        Dim spec_seismicSensor_only_checkBox_state As String
        spec_seismicSensor_only_checkBox_state =
            read_DbmsData(SPEC_SEISMIC_CANCEL_SW_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        If spec_seismicSensor_only_checkBox_state <> "" Then
            JobMaker_Form.Spec_SeismicSensor_Only_CheckBox.Checked = CBool(spec_seismicSensor_only_checkBox_state)
        End If
        '地震管制運轉-感知器N段 ONLY TEXTBOX
        JobMaker_Form.Spec_SeismicSensor_Only_TextBox.Text =
           read_DbmsData(SPEC_SEISMIC_CANCEL_SW_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '地震管制運轉-自動解除開關
        JobMaker_Form.Spec_SeismicSW_ComboBox.Text =
           read_DbmsData(SPEC_SEISMIC_CANCEL_SW,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '地震管制運轉-自動解除開關 ONLY CHECKBOX
        Dim spec_seismicSW_only_checkbox_state As String
        spec_seismicSW_only_checkbox_state =
            read_DbmsData(SPEC_SEISMIC_CANCEL_SW_ONLY_CHECKBOX,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If spec_seismicSW_only_checkbox_state <> "" Then
            JobMaker_Form.Spec_SeismicSW_Only_CheckBox.Checked = CBool(spec_seismicSW_only_checkbox_state)
        End If
        '地震管制運轉-自動解除開關 ONLY TEXTBOX
        JobMaker_Form.Spec_SeismicSW_Only_TextBox.Text =
           read_DbmsData(SPEC_SEISMIC_CANCEL_SW_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '車廂管制運轉燈
        JobMaker_Form.Spec_CPI_ComboBox.Text =
           read_DbmsData(SPEC_CPI,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂管制運轉燈-地震
        JobMaker_Form.Spec_CpiSeismic_ComboBox.Text =
           read_DbmsData(SPEC_CPI_SEISMIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂管制運轉燈-火災
        JobMaker_Form.Spec_CpiFire_ComboBox.Text =
           read_DbmsData(SPEC_CPI_FIRE,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂管制運轉燈-自發
        JobMaker_Form.Spec_CpiEmer_ComboBox.Text =
           read_DbmsData(SPEC_CPI_EMER,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂管制運轉燈-緊急
        JobMaker_Form.Spec_CpiFM_ComboBox.Text =
           read_DbmsData(SPEC_CPI_FM,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂管制運轉燈-滿載
        JobMaker_Form.Spec_CpiOLT_ComboBox.Text =
           read_DbmsData(SPEC_CPI_OLT,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂管制運轉燈-滿載 ONLY CHECKBOX
        Dim spec_cpiOLT_only_checkbox_state As String
        spec_cpiOLT_only_checkbox_state =
            read_DbmsData(SPEC_CPI_OLT_ONLY_CHECKBOX,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If spec_cpiOLT_only_checkbox_state <> "" Then
            JobMaker_Form.Spec_CpiOLT_Only_CheckBox.Checked = CBool(spec_cpiOLT_only_checkbox_state)
        End If
        '車廂管制運轉燈-滿載 ONLY TEXTBOX
        JobMaker_Form.Spec_CpiOLT_Only_TextBox.Text =
           read_DbmsData(SPEC_CPI_OLT_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴
        JobMaker_Form.Spec_CarGong_ComboBox.Text =
           read_DbmsData(SPEC_CAR_GONG,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '車廂上到著鈴-CAR [TOP] CHECKBOX
        Dim spec_carGong_top_checkbox_state As String
        spec_carGong_top_checkbox_state =
            read_DbmsData(SPEC_CAR_GONG_CARTOP_CHECKBOX,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If spec_carGong_top_checkbox_state <> "" Then
            JobMaker_Form.Spec_CarGong_Top_CheckBox.Checked = CBool(spec_carGong_top_checkbox_state)
        End If
        '車廂上到著鈴-CAR [TOP] TEXTBOX
        JobMaker_Form.Spec_CarGong_Top_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_CARTOP,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [TOP] ONLY CHECKBOX
        JobMaker_Form.Spec_CarGong_Top_Only_CheckBox.Text =
           read_DbmsData(SPEC_CAR_GONG_CARTOP_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [TOP] ONLY TEXTBOX
        JobMaker_Form.Spec_CarGong_Top_Only_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_CARTOP_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)


        '車廂上到著鈴-CAR [TOP BTM] CHECKBOX
        Dim spec_carGong_topBtm_checkbox_state As String
        spec_carGong_topBtm_checkbox_state =
            read_DbmsData(SPEC_CAR_GONG_CARTOPBTM_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        If spec_carGong_topBtm_checkbox_state <> "" Then
            JobMaker_Form.Spec_CarGong_TopBtm_CheckBox.Checked = CBool(spec_carGong_topBtm_checkbox_state)
        End If

        '車廂上到著鈴-CAR [TOP BTM] TEXTBOX
        JobMaker_Form.Spec_CarGong_TopBtm_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_CARTOPBTM,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [TOP BTM] ONLY CHECKBOX
        JobMaker_Form.Spec_CarGong_TopBtm_Only_CheckBox.Text =
           read_DbmsData(SPEC_CAR_GONG_CARTOPBTM_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [TOP BTM] ONLY TEXTBOX
        JobMaker_Form.Spec_CarGong_TopBtm_Only_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_CARTOPBTM_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)


        '車廂上到著鈴-CAR [COB] CHECKBOX
        Dim Spec_CarGong_COB_CheckBox_state As String
        Spec_CarGong_COB_CheckBox_state =
            read_DbmsData(SPEC_CAR_GONG_COB_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        If Spec_CarGong_COB_CheckBox_state <> "" Then
            JobMaker_Form.Spec_CarGong_COB_CheckBox.Checked = CBool(Spec_CarGong_COB_CheckBox_state)
        End If
        '車廂上到著鈴-CAR [COB] TEXTBOX
        JobMaker_Form.Spec_CarGong_COB_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_COB,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [COB] ONLY CHECKBOX
        JobMaker_Form.Spec_CarGong_COB_Only_CheckBox.Text =
           read_DbmsData(SPEC_CAR_GONG_COB_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [COB] ONLY TEXTBOX
        JobMaker_Form.Spec_CarGong_COB_Only_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_COB_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)


        '車廂上到著鈴-CAR [VONIC] CHECKBOX
        Dim Spec_CarGong_VONIC_CheckBox_state As String
        Spec_CarGong_VONIC_CheckBox_state =
            read_DbmsData(SPEC_CAR_GONG_VONIC_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        If Spec_CarGong_VONIC_CheckBox_state <> "" Then
            JobMaker_Form.Spec_CarGong_VONIC_CheckBox.Checked = CBool(Spec_CarGong_VONIC_CheckBox_state)
        End If
        '車廂上到著鈴-CAR [VONIC] TEXTBOX
        JobMaker_Form.Spec_CarGong_VONIC_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_VONIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [VONIC] ONLY CHECKBOX
        JobMaker_Form.Spec_CarGong_VONIC_Only_CheckBox.Text =
           read_DbmsData(SPEC_CAR_GONG_VONIC_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [VONIC] ONLY TEXTBOX
        JobMaker_Form.Spec_CarGong_VONIC_Only_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_VONIC_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)


        '乘場到著鈴
        JobMaker_Form.Spec_HallGong_ComboBox.Text =
           read_DbmsData(SPEC_HALL_GONG,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '乘場信號文字
        JobMaker_Form.Spec_HPIMsg_ComboBox.Text =
           read_DbmsData(SPEC_HPI,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '乘場信號文字-滿載
        JobMaker_Form.Spec_HpiOLT_ComboBox.Text =
           read_DbmsData(SPEC_HPI_OLT,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '乘場信號文字-保養
        JobMaker_Form.Spec_HpiMain_ComboBox.Text =
           read_DbmsData(SPEC_HPI_MAIN,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '乘場信號文字-專用
        JobMaker_Form.Spec_HpiIndep_ComboBox.Text =
           read_DbmsData(SPEC_HPI_INDEP,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '乘場信號文字-緊急
        JobMaker_Form.Spec_HpiFM_ComboBox.Text =
           read_DbmsData(SPEC_HPI_EMER,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '開門延長按鈕
        JobMaker_Form.Spec_DrHold_ComboBox.Text =
           read_DbmsData(SPEC_DR_HOLD,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '刷卡機
        JobMaker_Form.Spec_CRD_ComboBox.Text =
           read_DbmsData(SPEC_CRD,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '刷卡機-分層全層
        JobMaker_Form.Spec_CRDType_ComboBox.Text =
           read_DbmsData(SPEC_CRD_TYPE,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '刷卡機-ID:4
        JobMaker_Form.Spec_CRDID4_ComboBox.Text =
           read_DbmsData(SPEC_CRD_ID4,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '刷卡機-ID:5
        JobMaker_Form.Spec_CRDID5_ComboBox.Text =
           read_DbmsData(SPEC_CRD_ID5,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '自家發
        JobMaker_Form.Spec_Emer_ComboBox.Text =
           read_DbmsData(SPEC_EMER,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '自家發-群數
        JobMaker_Form.Spec_EmerNum_NumericUpDown.Value =
           read_DbmsData(SPEC_EMER_NUMBER,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '自家發-訊號
        JobMaker_Form.Spec_EmerSignal_ComboBox.Text =
           read_DbmsData(SPEC_EMER_SIGNAL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '自家發-緊急容量
        JobMaker_Form.Spec_EmerCapacity_NumericUpDown.Value =
           read_DbmsData(SPEC_EMER_CAPACITY,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '自家發-入力點
        JobMaker_Form.Spec_EmerInput_ComboBox.Text =
           read_DbmsData(SPEC_EMER_INPUT,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '自家發-Address
        JobMaker_Form.Spec_EmerAddress_ComboBox.Text =
           read_DbmsData(SPEC_EMER_ADDRESS,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '自家發-自動產生群組項目 基本資訊 -------------------------------------------------------------------

        Dim dyCtrlName As DynamicControlName = New DynamicControlName
            dyCtrlName.JobMaker_EmerInfo()

            dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_EmerNum_NumericUpDown,
                                      JobMaker_Form.Spec_emerGroup_TabControl,
                                      dyCtrlName.JobMaker_EmerTBInfoName_Array.Count,
                                      dyCtrlName.JobMaker_EmerTBInfoName_Array,
                                      SQLite_tableName_SpecTW)

        'If JobMaker_Form.Spec_EmerNum_NumericUpDown.Value <> 0 Then
        '    For group_i As Integer = 1 To CInt(JobMaker_Form.Spec_EmerNum_NumericUpDown.Value)
        '        For Each mTabControl As Control In JobMaker_Form.Spec_emerGroup_TabControl.Controls
        '            For Each mTabPage As Control In mTabControl.Controls
        '                For tb_j As Integer = 1 To dyCtrlName.JobMaker_EmerTBInfoName_Array.Count
        '                    If mTabPage.Name = $"{dyCtrlName.JobMaker_EmerTBInfoName_Array(tb_j - 1)}_{group_i}" Then
        '                        mTabPage.Text = read_DbmsData_RowID(dyCtrlName.JobMaker_LiftInfoName_Array(tb_j - 1),
        '                                                            SQLite_tableName_SpecTW,
        '                                                            SQLite_connectionPath_Job,
        '                                                            SQLite_JobDBMS_Name,
        '                                                            group_i)
        '                    End If

        '                Next
        '            Next
        '        Next
        '    Next
        'End If
        '------------------------------------------------------------------- 自家發-自動產生群組項目 基本資訊

        'LANDIC
        JobMaker_Form.Spec_Landic_ComboBox.Text =
           read_DbmsData(SPEC_LANDIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '基準階賦歸
        JobMaker_Form.Spec_MFLReturn_ComboBox.Text =
           read_DbmsData(SPEC_MFL_RETURN,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '基準階賦歸-基準階
        JobMaker_Form.Spec_MFLReturn_FL_TextBox.Text =
           read_DbmsData(SPEC_MFL_RETURN_FL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '語音撥放器VONIC
        JobMaker_Form.Spec_Vonic_ComboBox.Text =
           read_DbmsData(SPEC_VONIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '語音撥放器VONIC-標準
        JobMaker_Form.Spec_Vonic_standard_ComboBox.Text =
           read_DbmsData(SPEC_VONIC_STANDARD,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'ELVIC
        JobMaker_Form.Spec_Elvic_ComboBox.Text =
           read_DbmsData(SPEC_ELVIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'ELVIC-PARKING OPE
        Dim temp_spec_elvic_1_parking As String
        temp_spec_elvic_1_parking =
            read_DbmsData(SPEC_ELVIC_1_PARKING,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_1_parking <> "" Then
            JobMaker_Form.Spec_Elvic_Parking_CheckBox.Checked = temp_spec_elvic_1_parking
        End If
        'ELIVC-FLOOR LOCK OUT
        Dim temp_spec_elvic_1_fl_lockout As String
        temp_spec_elvic_1_fl_lockout =
            read_DbmsData(SPEC_ELVIC_1_FL_LOCKOUT,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_1_fl_lockout <> "" Then
            JobMaker_Form.Spec_Elvic_FloorLockOut_CheckBox.Checked = temp_spec_elvic_1_fl_lockout
        End If
        'ELVIC-VIP OPE
        Dim temp_spec_elvic_1_vip As String
        temp_spec_elvic_1_vip =
            read_DbmsData(SPEC_ELVIC_1_VIP,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_1_vip <> "" Then
            JobMaker_Form.Spec_Elvic_VIP_CheckBox.Checked = temp_spec_elvic_1_vip
        End If
        'ELVIC-INDEPENDENT OPE
        Dim temp_spec_elvic_1_indep As String
        temp_spec_elvic_1_indep =
            read_DbmsData(SPEC_ELVIC_1_INDEP,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_1_indep <> "" Then
            JobMaker_Form.Spec_Elvic_Indep_CheckBox.Checked = temp_spec_elvic_1_indep
        End If
        'ELVIC-RETURN TO DESIGNATED FLOOR
        Dim temp_spec_elvic_1_return As String
        temp_spec_elvic_1_return =
            read_DbmsData(SPEC_ELVIC_1_RETURN,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_1_return <> "" Then
            JobMaker_Form.Spec_Elvic_ReturnFL_CheckBox.Checked = temp_spec_elvic_1_return
        End If
        'ELVIC-CHANGE TRAFFIC PATTERN
        Dim temp_spec_elvic_2_traffic As String
        temp_spec_elvic_2_traffic =
            read_DbmsData(SPEC_ELVIC_2_TRAFFIC,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_2_traffic <> "" Then
            JobMaker_Form.Spec_Elvic_Traffic_Peak_CheckBox.Checked = temp_spec_elvic_2_traffic
        End If
        'ELVIC-UP PEAK
        Dim temp_spec_elvic_2_traffic_upPeak As String
        temp_spec_elvic_2_traffic_upPeak =
            read_DbmsData(SPEC_ELVIC_2_TRAFFIC_UPPEAK,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_2_traffic_upPeak <> "" Then
            JobMaker_Form.Spec_Elvic_Traffic_UpPeak_CheckBox.Checked = temp_spec_elvic_2_traffic_upPeak
        End If
        'ELVIC-DOWN PEAK
        Dim temp_spec_elvic_2_traffic_dnPeak As String
        temp_spec_elvic_2_traffic_dnPeak =
            read_DbmsData(SPEC_ELVIC_2_TRAFFIC_DNPEAK,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_2_traffic_dnPeak <> "" Then
            JobMaker_Form.Spec_Elvic_Traffic_DownPeak_CheckBox.Checked = temp_spec_elvic_2_traffic_dnPeak
        End If
        'ELVIC-LUNCH TIME 
        Dim temp_spec_elvic_2_traffic_lunch As String
        temp_spec_elvic_2_traffic_lunch =
            read_DbmsData(SPEC_ELVIC_2_TRAFFIC_LUNCH,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_2_traffic_lunch <> "" Then
            JobMaker_Form.Spec_Elvic_Traffic_Lunch_CheckBox.Checked = temp_spec_elvic_2_traffic_lunch
        End If
        'ELVIC-CHANGE MAIN FLOOR
        Dim temp_spec_elvic_2_mfl As String
        temp_spec_elvic_2_mfl =
            read_DbmsData(SPEC_ELVIC_2_MFL,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_2_mfl <> "" Then
            JobMaker_Form.Spec_Elvic_MainFL_CheckBox.Checked = temp_spec_elvic_2_mfl
        End If

        'ELVIC-ZONING FOR EXPRESS OPE
        Dim temp_spec_elvic_zoning_express As String
        temp_spec_elvic_zoning_express =
            read_DbmsData(SPEC_ELVIC_2_ZONING_EXPRESS,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_zoning_express <> "" Then
            JobMaker_Form.Spec_Elvic_Zoning_CheckBox.Checked = temp_spec_elvic_zoning_express
        End If

        'ELVIC-FLOOR LOCK OUT
        Dim temp_spec_elvic_2_fl_lockout As String
        temp_spec_elvic_2_fl_lockout =
            read_DbmsData(SPEC_ELVIC_2_FL_LOCKOUT,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_2_fl_lockout <> "" Then
            JobMaker_Form.Spec_Elvic_FloorLockOut_GR_CheckBox.Checked = temp_spec_elvic_2_fl_lockout
        End If

        'ELVIC-CAR CALL DISCONNECT
        Dim temp_spec_elvic_2_carcall As String
        temp_spec_elvic_2_carcall =
            read_DbmsData(SPEC_ELVIC_2_CARCALL,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_2_carcall <> "" Then
            JobMaker_Form.Spec_Elvic_CarCall_CheckBox.Checked = temp_spec_elvic_2_carcall
        End If

        'ELVIC-FIRE OPE. COMMAND
        Dim temp_spec_elvic_fire As String
        temp_spec_elvic_fire =
            read_DbmsData(SPEC_ELVIC_3_FIRE,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_fire <> "" Then
            JobMaker_Form.Spec_Elvic_Fire_CheckBox.Checked = temp_spec_elvic_fire
        End If

        'ELVIC-WAVIC OPE. COMMAND
        Dim temp_spec_elvic_3_wavic As String
        temp_spec_elvic_3_wavic =
            read_DbmsData(SPEC_ELVIC_3_WAVIC,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_3_wavic <> "" Then
            JobMaker_Form.Spec_Elvic_Wavic_CheckBox.Checked = temp_spec_elvic_3_wavic
        End If

        'ELVIC-CARE READER COMMAND
        Dim temp_spec_elvic_3_card As String
        temp_spec_elvic_3_card =
            read_DbmsData(SPEC_ELVIC_3_CARD,
                          SQLite_tableName_SpecTW,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If temp_spec_elvic_3_card <> "" Then
            JobMaker_Form.Spec_Elvic_CRD_CheckBox.Checked = temp_spec_elvic_3_card
        End If

        '乘場廳燈
        JobMaker_Form.Spec_HLL_ComboBox.Text =
           read_DbmsData(SPEC_HLL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '殘障仕樣
        JobMaker_Form.Spec_WCOB_ComboBox.Text =
           read_DbmsData(SPEC_WCOB,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '殘障仕樣-Only CHECKBOX
        JobMaker_Form.Spec_WCOB_only_CheckBox.Checked =
           read_DbmsData(SPEC_WCOB_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '殘障仕樣-Only TEXTBOX
        JobMaker_Form.Spec_WCOB_only_TextBox.Text =
           read_DbmsData(SPEC_WCOB_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '殘障仕樣-SCOB
        JobMaker_Form.Spec_WSCOB_ComboBox.Text =
           read_DbmsData(SPEC_WSCOB,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '殘障仕樣-SCOB ONLY CHECKBOX
        JobMaker_Form.Spec_WSCOB_only_CheckBox.Text =
           read_DbmsData(SPEC_WSCOB_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '殘障仕樣-SCOB ONLY TEXTBOX
        JobMaker_Form.Spec_WSCOB_only_TextBox.Text =
           read_DbmsData(SPEC_WSCOB_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '殘障仕樣-鳴動
        JobMaker_Form.Spec_WCOB_Ring_ComboBox.Text =
           read_DbmsData(SPEC_WCOB_RING,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '運轉手盤運轉
        JobMaker_Form.Spec_ATT_ComboBox.Text =
           read_DbmsData(SPEC_ATT,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '浸水管制運轉
        JobMaker_Form.Spec_Flood_ComboBox.Text =
           read_DbmsData(SPEC_FLOOD,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '浸水管制運轉-停止階
        JobMaker_Form.Spec_Flood_FL_TextBox.Text =
           read_DbmsData(SPEC_FLOOD_FL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'LS1M頂部緊急停止開關
        JobMaker_Form.Spec_LS1M_ComboBox.Text =
           read_DbmsData(SPEC_LS1M,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '電力回升
        JobMaker_Form.Spec_PRU_ComboBox.Text =
           read_DbmsData(SPEC_PRU,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell
        JobMaker_Form.Spec_LoadCell_ComboBox.Text =
           read_DbmsData(SPEC_LOAD_CELL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell-裝置在
        JobMaker_Form.Spec_LoadCellPos_ComboBox.Text =
           read_DbmsData(SPEC_LOAD_CELL_POSITION,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB
        JobMaker_Form.Spec_WTB_ComboBox.Text =
           read_DbmsData(SPEC_WTB,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-故障燈
        JobMaker_Form.Spec_WTB_Error_ComboBox.Text =
           read_DbmsData(SPEC_WTB_ERROR,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-休止燈
        JobMaker_Form.Spec_WTB_Stop_ComboBox.Text =
           read_DbmsData(SPEC_WTB_STOP,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-消防燈
        JobMaker_Form.Spec_WTB_FM_ComboBox.Text =
           read_DbmsData(SPEC_WTB_FIREMAN,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-正常燈
        JobMaker_Form.Spec_WTB_Normal_ComboBox.Text =
           read_DbmsData(SPEC_WTB_NORMAL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-緊急電源燈
        JobMaker_Form.Spec_WTB_Urgent_ComboBox.Text =
           read_DbmsData(SPEC_WTB_URGENT,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-火災燈
        JobMaker_Form.Spec_WTB_FO_ComboBox.Text =
           read_DbmsData(SPEC_WTB_FO,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-自家發燈
        JobMaker_Form.Spec_WTB_EmerPow_ComboBox.Text =
           read_DbmsData(SPEC_WTB_EMER,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-警示燈
        JobMaker_Form.Spec_WTB_Alart_ComboBox.Text =
           read_DbmsData(SPEC_WTB_ALART,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-地震燈
        JobMaker_Form.Spec_WTB_EQ_ComboBox.Text =
           read_DbmsData(SPEC_WTB_EQ,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-專用燈
        JobMaker_Form.Spec_WTB_Indep_ComboBox.Text =
           read_DbmsData(SPEC_WTB_INDEP,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-地震開關
        JobMaker_Form.Spec_WTB_EQSW_ComboBox.Text =
           read_DbmsData(SPEC_WTB_EQSW,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-bz解除開關
        JobMaker_Form.Spec_WTB_BZSW_ComboBox.Text =
           read_DbmsData(SPEC_WTB_BZSW,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-check開關
        JobMaker_Form.Spec_WTB_ChkSW_ComboBox.Text =
           read_DbmsData(SPEC_WTB_CHKSW,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-停車開關
        JobMaker_Form.Spec_WTB_PKSW_ComboBox.Text =
           read_DbmsData(SPEC_WTB_PKSW,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-地震指示器
        JobMaker_Form.Spec_WTB_EQIND_ComboBox.Text =
           read_DbmsData(SPEC_WTB_EQIND,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WTB-地震強度
        JobMaker_Form.Spec_WTB_EQMac_ComboBox.Text =
           read_DbmsData(SPEC_WTB_EQMAC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '正背門
        JobMaker_Form.Spec_FrontRearDr_ComboBox.Text =
           read_DbmsData(SPEC_FRONT_REAR_DR,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '各停開關
        JobMaker_Form.Spec_EachStop_ComboBox.Text =
           read_DbmsData(SPEC_EACH_STOP,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '拒付運轉
        JobMaker_Form.Spec_install_ope_ComboBox.Text =
           read_DbmsData(SPEC_INSTALL_OPE,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'vonic蜂鳴器
        JobMaker_Form.Spec_VonicBz_ComboBox.Text =
           read_DbmsData(SPEC_VONICBZ,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '單群控切換
        JobMaker_Form.Spec_OpeSw_ComboBox.Text =
           read_DbmsData(SPEC_OPE_SW,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '單群控切換-入力點Position
        JobMaker_Form.Spec_OpeSw_InputPos_ComboBox.Text =
           read_DbmsData(SPEC_OPE_SW_POS,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '單群控切換-入力點Address
        JobMaker_Form.Spec_OpeSw_InputAddress_TextBox.Text =
           read_DbmsData(SPEC_OPE_SW_ADDRESS,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
    End Sub
    Private Sub Important_Load()
        'IDU CheckBox
        JobMaker_Form.Use_Imp_CheckBox.Checked =
           read_DbmsData(IMPORTANT_Use_ChkBox,
                         SQLite_tableName_Important,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '風扇連動
        'JobMaker_Form.Imp_FAN_ComboBox.Text =
        '   read_DbmsData(IMPORTANT_FAN,
        '                 SQLite_tableName_Important,
        '                 SQLite_connectionPath_Job,
        '                 SQLite_JobDBMS_Name)

        'OVER BALANCE
        JobMaker_Form.Imp_OverBalance_ComboBox.Text =
           read_DbmsData(IMPORTANT_BALANCE,
                         SQLite_tableName_Important,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'WCOB
        JobMaker_Form.Imp_WHB_ComboBox.Text =
           read_DbmsData(IMPORTANT_WCOB,
                         SQLite_tableName_Important,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'DOOR TYPE
        JobMaker_Form.Imp_DoorType_TextBox.Text =
           read_DbmsData(IMPORTANT_DOOR,
                         SQLite_tableName_Important,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'Hall Indicator中的號機基本資訊 -------------------------------------------------------------------
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_HINInfo()

        If JobMaker_Form.Spec_LiftNum_NumericUpDown.Value <> 0 Then
            For lift_i As Integer = 1 To CInt(JobMaker_Form.Spec_LiftNum_NumericUpDown.Value)
                If coverFile_bool = False Then
                    If lift_i < JobMaker_Form.Spec_LiftNum_NumericUpDown.Value Then
                        Insert_DbmsData(dyCtrlName.JobMaker_HINInfoName_Array(0),
                                        SQLite_tableName_Important,
                                        SQLite_connectionPath_Job,
                                        SQLite_JobDBMS_Name)
                    End If
                    For Each mFlowLayoutPanel As Control In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls
                        For ctrl_j As Integer = 1 To dyCtrlName.JobMaker_HINInfoName_Array.Count
                            If mFlowLayoutPanel.Name = $"{dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1)}_{lift_i}" Then
                                update_DbmsData(dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1),
                                                mFlowLayoutPanel.Text,
                                                SQLite_tableName_Important,
                                                SQLite_connectionPath_Job,
                                                SQLite_JobDBMS_Name,
                                                lift_i)
                            ElseIf dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1) = dyCtrlName.JobMaker_HIN_FL_ChkB Or
                                   dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1) = dyCtrlName.JobMaker_HIN_FL_CmbB Then

                                For stopFL_k As Integer = 1 To JobMaker_Form.arr_liftStopFL(lift_i - 1)
                                    If mFlowLayoutPanel.Name = $"{stopFL_k}{dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1)}_{lift_i}" Then
                                        update_DbmsData(dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1),
                                                        mFlowLayoutPanel.Text,
                                                        SQLite_tableName_Important,
                                                        SQLite_connectionPath_Job,
                                                        SQLite_JobDBMS_Name,
                                                        lift_i)
                                    End If
                                Next
                            End If
                        Next
                    Next
                Else
                    Dim temp_specBasic_liftNumber As String
                    temp_specBasic_liftNumber =
                        read_DbmsData(SpecBasic_LiftNumber,
                                      SQLite_tableName_SpecBasic,
                                      SQLite_connectionPath_Job,
                                      SQLite_JobDBMS_Name)
                    Dim overwrite_liftNumber_bool As Boolean


                    If temp_specBasic_liftNumber <> JobMaker_Form.Spec_LiftNum_NumericUpDown.Value Then
                        '比對電梯總數不相同，需要更改
                        overwrite_liftNumber_bool = True

                        '如果新的電梯數量比舊的多，則要插入新的行在SQLite中 ---------------------------------
                        If JobMaker_Form.Spec_LiftNum_NumericUpDown.Value > temp_specBasic_liftNumber Then
                            Dim tempSub_num As Integer
                            tempSub_num = CInt(JobMaker_Form.Spec_LiftNum_NumericUpDown.Value) - CInt(temp_specBasic_liftNumber)
                            For insertRow_i = 1 To tempSub_num
                                Insert_DbmsData(dyCtrlName.JobMaker_LiftInfoName_Array(0),
                                                SQLite_tableName_SpecBasic,
                                                SQLite_connectionPath_Job,
                                                SQLite_JobDBMS_Name)
                            Next
                        End If
                        '---------------------------------如果新的數量比舊的多，則要插入新的行在SQLite中 
                    Else
                        '數量相同但內容不同，需要更改
                        For Each tempCtrl As Control In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls
                            For hin_j As Integer = 1 To dyCtrlName.JobMaker_HINInfoName_Array.Count
                                If tempCtrl.Name = $"{dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1)}_{lift_i}" Then
                                    If tempCtrl.Text <> read_DbmsData_RowID(dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1),
                                                                            SQLite_tableName_Important,
                                                                            SQLite_connectionPath_Job,
                                                                            SQLite_JobDBMS_Name,
                                                                            lift_i) Then
                                        overwrite_liftNumber_bool = True
                                        Exit For
                                    Else
                                        overwrite_liftNumber_bool = False
                                    End If
                                ElseIf dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1) = dyCtrlName.JobMaker_HIN_FL_ChkB Or
                                       dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1) = dyCtrlName.JobMaker_HIN_FL_CmbB Then
                                    If tempCtrl.Text <> read_DbmsData_RowID(dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1),
                                                                            SQLite_tableName_Important,
                                                                            SQLite_connectionPath_Job,
                                                                            SQLite_JobDBMS_Name,
                                                                            lift_i) Then
                                        overwrite_liftNumber_bool = True
                                        Exit For
                                    Else
                                        overwrite_liftNumber_bool = False
                                    End If
                                End If
                            Next
                            If overwrite_liftNumber_bool = True Then
                                Exit For
                            End If
                        Next
                    End If

                    If overwrite_liftNumber_bool Then
                        '當下更新的電梯內容與紀錄中的比較，如果有一處不同就全數刪除設="" -------
                        If lift_i <= 1 Then
                            For hin_j As Integer = 1 To dyCtrlName.JobMaker_HINInfoName_Array.Count
                                update_DbmsData(dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1),
                                                "",
                                                SQLite_tableName_Important,
                                                SQLite_connectionPath_Job,
                                                SQLite_JobDBMS_Name)
                            Next
                        End If
                        '-------當下更新的電梯內容與紀錄中的比較，如果有一處不同就全數刪除設=""

                        '更新新的CheckListBox ----------------------------------------------------------------
                        For Each tempCtrl As Control In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls

                            For hin_j As Integer = 1 To dyCtrlName.JobMaker_HINInfoName_Array.Count
                                If tempCtrl.Name = $"{dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1)}_{lift_i}" Then
                                    update_DbmsData(dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1),
                                                    tempCtrl.Text,
                                                    SQLite_tableName_Important,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    lift_i)
                                ElseIf dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1) = dyCtrlName.JobMaker_HIN_FL_ChkB Or
                                       dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1) = dyCtrlName.JobMaker_HIN_FL_CmbB Then
                                    update_DbmsData(dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1),
                                                    tempCtrl.Text,
                                                    SQLite_tableName_Important,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    lift_i)
                                End If
                            Next
                        Next
                        '---------------------------------------------------------------- 更新新的CheckListBox 
                    End If
                End If
            Next
        End If
        '------------------------------------------------------------------- Hall Indicator中的號機基本資訊
    End Sub
    Private Sub MMIC_Load()
        'MMIC CheckBox
        JobMaker_Form.Use_mmic_CheckBox.Checked =
           read_DbmsData(MMIC_Use_ChkBox,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'MMIC MR BASE
        JobMaker_Form.MMIC_MR_Base_TextBox.Text =
           read_DbmsData(MMIC_MR_BASE,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC MR CP43
        JobMaker_Form.MMIC_MR_CP43x_ComboBox.Text =
           read_DbmsData(MMIC_MR_CP43x,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'MMIC MR EEPROM BASE
        JobMaker_Form.MMIC_MR_EBase_ComboBox.Text =
           read_DbmsData(MMIC_MR_EBase,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC MR EEPROM Car Obj
        JobMaker_Form.MMIC_MR_ECarObj_ComboBox.Text =
           read_DbmsData(MMIC_MR_ECarObj,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC SV TYPE
        JobMaker_Form.MMIC_SV_Type_ComboBox.Text =
           read_DbmsData(MMIC_SV_TYPE,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC SV BASE
        JobMaker_Form.MMIC_SV_Base_TextBox.Text =
           read_DbmsData(MMIC_SV_BASE,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC SV EEPROM BASE
        JobMaker_Form.MMIC_SV_EBase_ComboBox.Text =
           read_DbmsData(MMIC_SV_EBase,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC SV EEPROM Car Obj
        JobMaker_Form.MMIC_SV_ECarObj_ComboBox.Text =
           read_DbmsData(MMIC_SV_ECarObj,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC VD10 ROM DEVICE
        JobMaker_Form.MMIC_VD10_ROM_ComboBox.Text =
           read_DbmsData(MMIC_VD10_ROM,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC VD10 Quantity
        JobMaker_Form.MMIC_VD10_Quantity_ComboBox.Text =
           read_DbmsData(MMIC_VD10_Quantity,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC VD10 Type
        JobMaker_Form.MMIC_VD10_Type_ComboBox.Text =
           read_DbmsData(MMIC_VD10_TYPE,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MMIC VD10 BASE
        JobMaker_Form.MMIC_VD10_Base_TextBox.Text =
           read_DbmsData(MMIC_VD10_BASE,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'MR Number
        JobMaker_Form.MMIC_MR_NumericUpDown.Value =
           read_DbmsData(MMIC_MR_Number,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'MR EEPROM Number
        JobMaker_Form.MMIC_MR_E_NumericUpDown.Value =
           read_DbmsData(MMIC_MR_ENumber,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'SV Number 
        JobMaker_Form.MMIC_SV_NumericUpDown.Value =
           read_DbmsData(MMIC_SV_Number,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'SV EEPROM Number
        JobMaker_Form.MMIC_SV_NumericUpDown.Value =
           read_DbmsData(MMIC_SV_ENumber,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'VD10 Number 
        JobMaker_Form.MMIC_VD10_NumericUpDown.Value =
           read_DbmsData(MMIC_VD10_Number,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'MMIC 各panel中的號機基本資訊 -------------------------------------------------------------------
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_MMICInfo()

        'MR
        dynamicPanel_ReadFromDbms(JobMaker_Form.MMIC_MR_NumericUpDown,
                                  JobMaker_Form.MMIC_MR_Panel,
                                  dyCtrlName.JobMaker_MMIC_MrBase_InfoName_Array.Count,
                                  dyCtrlName.JobMaker_MMIC_MrBase_InfoName_Array,
                                  SQLite_tableName_MMIC)
        'MR-EEPROM DATA
        dynamicPanel_ReadFromDbms(JobMaker_Form.MMIC_MR_E_NumericUpDown,
                                  JobMaker_Form.MMIC_MR_E_Panel,
                                  dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array.Count,
                                  dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array,
                                  SQLite_tableName_MMIC)

        'SV
        dynamicPanel_ReadFromDbms(JobMaker_Form.MMIC_SV_NumericUpDown,
                                  JobMaker_Form.MMIC_SV_Panel,
                                  dyCtrlName.JobMaker_MMIC_SvBase_InfoName_Array.Count,
                                  dyCtrlName.JobMaker_MMIC_SvBase_InfoName_Array,
                                  SQLite_tableName_MMIC)

        'SV-EEPROM DATA
        dynamicPanel_ReadFromDbms(JobMaker_Form.MMIC_SV_E_NumericUpDown,
                                  JobMaker_Form.MMIC_SV_E_Panel,
                                  dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array.Count,
                                  dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array,
                                  SQLite_tableName_MMIC)

        'VD10
        dynamicPanel_ReadFromDbms(JobMaker_Form.MMIC_VD10_NumericUpDown,
                                  JobMaker_Form.MMIC_VD10_Panel,
                                  dyCtrlName.JobMaker_MMIC_VD10Base_InfoName_Array.Count,
                                  dyCtrlName.JobMaker_MMIC_VD10Base_InfoName_Array,
                                  SQLite_tableName_MMIC)

        '------------------------------------------------------------------- panel中的號機基本資訊 
    End Sub
    '------------------------------------------ 載入SQLite 

    ''' <summary>
    ''' [自動產生Panel中，僅讀取單層控制項或是雙層控制項]
    ''' </summary>
    Enum LoadStored_PanelType
        SingleLayer_Panel
        DoubleLayer_Panel
    End Enum

    ''' <summary>
    ''' [儲存自動產生控制項的值]
    ''' </summary>
    ''' <param name="mPanelType">判斷單層或雙層結構</param>
    ''' <param name="mNumericUpDown">NumericUpDown控制項</param>
    ''' <param name="dyCtrl_ArrayCount">自動生成的控制項名稱總數</param>
    ''' <param name="dyCtrl_Array">自動生成的控制項名稱陣列</param>
    ''' <param name="mPanel">Panel控制項</param>
    ''' <param name="sqlite_selectName">儲存NumericUpDown總數量的SQLite名稱</param>
    ''' <param name="sqlite_tableName"></param>
    Overloads Sub dynamicPanel_StoredIntoDbms(mPanelType As LoadStored_PanelType,
                                              mNumericUpDown As NumericUpDown,
                                              dyCtrl_ArrayCount As Integer, dyCtrl_Array As Array,
                                              mPanel As Control,
                                              sqlite_selectName_Number As String, sqlite_tableName As String)

        'panel中的號機基本資訊 -------------------------------------------------------------------
        If mNumericUpDown.Value <> 0 Then
            For lift_i As Integer = 1 To CInt(mNumericUpDown.Value)
                '判斷是否為覆蓋檔案
                If coverFile_bool = False Then
                    If lift_i < mNumericUpDown.Value Then
                        '先插入空值
                        Insert_DbmsData(dyCtrl_Array(0),
                                        sqlite_tableName,
                                        SQLite_connectionPath_Job,
                                        SQLite_JobDBMS_Name)
                    End If
                    For Each tempCtrl_Panel As Control In mPanel.Controls
                        '判斷Panel為單層或雙層
                        If mPanelType = LoadStored_PanelType.SingleLayer_Panel Then
                            For lift_j As Integer = 1 To dyCtrl_ArrayCount
                                If tempCtrl_Panel.Name = $"{dyCtrl_Array(lift_j - 1)}_{lift_i}" Then
                                    update_DbmsData(dyCtrl_Array(lift_j - 1),
                                                    tempCtrl_Panel.Text,
                                                    sqlite_tableName,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    lift_i)
                                End If
                            Next
                        ElseIf mPanelType = LoadStored_PanelType.DoubleLayer_Panel Then
                            '雙層差別在這一層 -------------------------------------------------
                            For Each tempCtrl_DoublePanel As Control In tempCtrl_Panel.Controls
                                '----------------------------------------------雙層差別在這一層 
                                For lift_j As Integer = 1 To dyCtrl_ArrayCount
                                    If tempCtrl_DoublePanel.Name = $"{dyCtrl_Array(lift_j - 1)}_{lift_i}" Then
                                        update_DbmsData(dyCtrl_Array(lift_j - 1),
                                                        tempCtrl_DoublePanel.Text,
                                                        sqlite_tableName,
                                                        SQLite_connectionPath_Job,
                                                        SQLite_JobDBMS_Name,
                                                        lift_i)
                                    End If
                                Next
                            Next
                        End If
                    Next
                Else
                    Dim temp_Number As String
                    temp_Number =
                        read_DbmsData(sqlite_selectName_Number,
                                      sqlite_tableName,
                                      SQLite_connectionPath_Job,
                                      SQLite_JobDBMS_Name)
                    Dim overwrite_liftNumber_bool As Boolean


                    '比對<原本Sqlite內數量>與<現在數量> 不相同時需要更改
                    If temp_Number <> mNumericUpDown.Value Then
                        overwrite_liftNumber_bool = True

                        '如果新的數量比舊的多，則要插入新的行在SQLite中 ---------------------------------
                        If mNumericUpDown.Value > temp_Number Then
                            Dim tempSub_Number As Integer
                            tempSub_Number = CInt(mNumericUpDown.Value) - CInt(temp_Number)
                            For insertRow_i = 1 To tempSub_Number
                                Insert_DbmsData(dyCtrl_Array(0),
                                                sqlite_tableName,
                                                SQLite_connectionPath_Job,
                                                SQLite_JobDBMS_Name)
                            Next
                        End If
                        '---------------------------------如果新的數量比舊的多，則要插入新的行在SQLite中 
                    Else
                        '數量相同但內容不同，需要更改 '---------------------------------
                        For Each tempCtrl_Panel As Control In mPanel.Controls
                            If mPanelType = LoadStored_PanelType.SingleLayer_Panel Then
                                For liftNum_j As Integer = 1 To dyCtrl_ArrayCount
                                    If tempCtrl_Panel.Name = $"{dyCtrl_Array(liftNum_j - 1)}_{lift_i}" Then
                                        If tempCtrl_Panel.Text <> read_DbmsData_RowID(dyCtrl_Array(liftNum_j - 1),
                                                                                      sqlite_tableName,
                                                                                      SQLite_connectionPath_Job,
                                                                                      SQLite_JobDBMS_Name,
                                                                                      lift_i) Then
                                            '新資料 與 舊資料 不同時 更新
                                            overwrite_liftNumber_bool = True
                                            Exit For
                                        Else
                                            '新資料 與 舊資料 相同時 不更新
                                            overwrite_liftNumber_bool = False
                                        End If
                                    End If
                                Next
                                If overwrite_liftNumber_bool = True Then
                                    Exit For
                                End If

                            ElseIf mPanelType = LoadStored_PanelType.DoubleLayer_Panel Then
                                For Each tempCtrl_DoublePanel As Control In tempCtrl_Panel.Controls
                                    For liftNum_j As Integer = 1 To dyCtrl_ArrayCount
                                        If tempCtrl_DoublePanel.Name = $"{dyCtrl_Array(liftNum_j - 1)}_{lift_i}" Then
                                            If tempCtrl_DoublePanel.Text <> read_DbmsData_RowID(dyCtrl_Array(liftNum_j - 1),
                                                                                                sqlite_tableName,
                                                                                                SQLite_connectionPath_Job,
                                                                                                SQLite_JobDBMS_Name,
                                                                                                lift_i) Then
                                                '新資料 與 舊資料 不同時 更新
                                                overwrite_liftNumber_bool = True
                                                Exit For
                                            Else
                                                '新資料 與 舊資料 相同時 不更新
                                                overwrite_liftNumber_bool = False
                                            End If
                                        End If
                                    Next
                                    If overwrite_liftNumber_bool = True Then
                                        Exit For
                                    End If
                                Next

                            End If
                        Next
                        '--------------------------------- 數量相同但內容不同，需要更改 
                    End If

                    If overwrite_liftNumber_bool Then
                        '當下更新的電梯內容與紀錄中的比較，如果有一處不同就全數刪除 設="" -------
                        If lift_i <= 1 Then
                            For liftNum_j As Integer = 1 To dyCtrl_ArrayCount
                                update_DbmsData(dyCtrl_Array(liftNum_j - 1),
                                            "",
                                            sqlite_tableName,
                                            SQLite_connectionPath_Job,
                                            SQLite_JobDBMS_Name)
                            Next
                        End If
                        '-------當下更新的電梯內容與紀錄中的比較，如果有一處不同就全數刪除 設=""

                        '更新新的CheckListBox ----------------------------------------------------------------
                        For Each tempCtrl_Panel As Control In mPanel.Controls
                            If mPanelType = LoadStored_PanelType.SingleLayer_Panel Then
                                '八組自動生成TextBox
                                For lift_j As Integer = 1 To dyCtrl_ArrayCount
                                    If tempCtrl_Panel.Name = $"{dyCtrl_Array(lift_j - 1)}_{lift_i}" Then
                                        update_DbmsData(dyCtrl_Array(lift_j - 1),
                                                        tempCtrl_Panel.Text,
                                                        sqlite_tableName,
                                                        SQLite_connectionPath_Job,
                                                        SQLite_JobDBMS_Name,
                                                        lift_i)
                                    End If
                                Next
                            ElseIf mPanelType = LoadStored_PanelType.DoubleLayer_Panel Then
                                For Each tempCtrl_DoublePanel As Control In tempCtrl_Panel.Controls
                                    For lift_j As Integer = 1 To dyCtrl_ArrayCount
                                        If tempCtrl_DoublePanel.Name = $"{dyCtrl_Array(lift_j - 1)}_{lift_i}" Then
                                            update_DbmsData(dyCtrl_Array(lift_j - 1),
                                                            tempCtrl_DoublePanel.Text,
                                                            sqlite_tableName,
                                                            SQLite_connectionPath_Job,
                                                            SQLite_JobDBMS_Name,
                                                            lift_i)
                                        End If
                                    Next
                                Next
                            End If
                        Next
                        '---------------------------------------------------------------- 更新新的CheckListBox 
                    End If
                End If
            Next
        End If
    End Sub


    ''' <summary>
    ''' [寫入自動生成Panel內的控制項]
    ''' </summary>
    ''' <param name="mNumericUpDown"></param>
    ''' <param name="mPanel"></param>
    ''' <param name="dyCtrl_ArrayCount"></param>
    ''' <param name="dyCtrl_Array"></param>
    ''' <param name="tableName"></param>
    Overloads Sub dynamicPanel_ReadFromDbms(mNumericUpDown As NumericUpDown, mPanel As Panel, dyCtrl_ArrayCount As Integer,
                                            dyCtrl_Array As Array, tableName As String)
        If mNumericUpDown.Value <> 0 Then
            For lift_i As Integer = 1 To CInt(mNumericUpDown.Value)
                For Each tempCtrl As Control In mPanel.Controls
                    For obj_j As Integer = 1 To dyCtrl_ArrayCount
                        If tempCtrl.Name = $"{dyCtrl_Array(obj_j - 1)}_{lift_i}" Then
                            tempCtrl.Text =
                                read_DbmsData_RowID(dyCtrl_Array(obj_j - 1),
                                                    tableName,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    lift_i)
                        End If
                    Next
                Next
            Next
        End If
    End Sub

    ''' <summary>
    ''' [寫入自動生成TabContol內的控制項]
    ''' </summary>
    ''' <param name="mNumericUpDown"></param>
    ''' <param name="mTabControl"></param>
    ''' <param name="dyCtrl_ArrayCount"></param>
    ''' <param name="dyCtrl_Array"></param>
    ''' <param name="tableName"></param>
    Overloads Sub dynamicPanel_ReadFromDbms(mNumericUpDown As NumericUpDown, mTabControl As TabControl, dyCtrl_ArrayCount As Integer,
                                            dyCtrl_Array As Array, tableName As String)
        If mNumericUpDown.Value <> 0 Then
            For lift_i As Integer = 1 To CInt(mNumericUpDown.Value)
                For Each mCtrl_TabControl As Control In mTabControl.Controls
                    For Each mCtrl_TabPage As Control In mCtrl_TabControl.Controls
                        For obj_j As Integer = 1 To dyCtrl_ArrayCount
                            If mCtrl_TabPage.Name = $"{dyCtrl_Array(obj_j - 1)}_{lift_i}" Then
                                mCtrl_TabPage.Text =
                                    read_DbmsData_RowID(dyCtrl_Array(obj_j - 1),
                                                        tableName,
                                                        SQLite_connectionPath_Job,
                                                        SQLite_JobDBMS_Name,
                                                        lift_i)
                            End If
                        Next
                    Next
                Next
            Next
        End If
    End Sub



    ''' <summary>
    ''' 更新指定的SQLite表格
    ''' </summary>
    ''' <param name="SQLite_CellName">欄位</param>
    ''' <param name="SQLite_CellName_value">欄位數值填入</param>
    ''' <param name="SQLite_tableName">表格</param>
    ''' <param name="SQLite_path">路徑</param>
    ''' <param name="SQLite_FileName">檔案名稱</param>
    ''' <returns></returns>
    Overloads Function update_DbmsData(SQLite_CellName As String, SQLite_CellName_value As String, SQLite_tableName As String,
                                       SQLite_path As String, SQLite_FileName As String)
        '----------------------- SQLite Reading -----------------------------
        SQLite_storedGrammer = $"UPDATE {SQLite_tableName} SET {SQLite_CellName} = ""{SQLite_CellName_value}"";"
        Try
            Using msqlite_connect As New SQLiteConnection("Data Source=" & SQLite_path & SQLite_FileName)
                msqlite_connect.Open()
                Using msqlite_command = New SQLiteCommand(SQLite_storedGrammer, msqlite_connect)
                    msqlite_command.ExecuteNonQuery()
                    msqlite_command.Dispose()
                    msqlite_connect.Close()
                End Using
            End Using

            JobMaker_Form.ResultOutput_TextBox.Text += $"{SQLite_CellName} : {SQLite_CellName_value}成功更新{vbCrLf}"
        Catch e As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{SQLite_CellName} : {SQLite_CellName_value}失敗更新{vbCrLf}"
        End Try
        '----------------------- SQLite Reading -----------------------------
    End Function
    ''' <summary>
    ''' 更新特定欄位SQLite的表格
    ''' </summary>
    ''' <param name="SQLite_CellName">欄位</param>
    ''' <param name="SQLite_CellName_value">欄位數值填入</param>
    ''' <param name="SQLite_tableName">表格</param>
    ''' <param name="SQLite_path">路徑</param>
    ''' <param name="SQLite_FileName">檔案名稱</param>
    ''' <param name="rowID">特定欄位</param>
    ''' <returns></returns>
    Overloads Function update_DbmsData(SQLite_CellName As String, SQLite_CellName_value As String, SQLite_tableName As String,
                                       SQLite_path As String, SQLite_FileName As String, rowID As Integer)
        '----------------------- SQLite Reading -----------------------------

        SQLite_storedGrammer = $"UPDATE {SQLite_tableName} SET {SQLite_CellName} = ""{SQLite_CellName_value}"" WHERE ROWID = {rowID};"

        Try
            Using msqlite_connect As New SQLiteConnection("Data Source=" & SQLite_path & SQLite_FileName)
                msqlite_connect.Open()
                Using msqlite_command = New SQLiteCommand(SQLite_storedGrammer, msqlite_connect)
                    msqlite_command.ExecuteNonQuery()
                    msqlite_command.Dispose()
                    msqlite_connect.Close()
                End Using
            End Using
            JobMaker_Form.ResultOutput_TextBox.Text += $"{SQLite_CellName} : {SQLite_CellName_value}成功更新{vbCrLf}"
        Catch e As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{SQLite_CellName} : {SQLite_CellName_value}失敗更新{vbCrLf}"
        End Try
        '----------------------- SQLite Reading -----------------------------
    End Function

    'Overloads Function clearAll_DbmsData(SQLite_CellName As String, SQLite_CellName_value As String, SQLite_tableName As String,
    '                                   SQLite_path As String, SQLite_FileName As String, rowID As Integer)
    '    '----------------------- SQLite Reading -----------------------------
    '    SQLite_storedGrammer = $"UPDATE {SQLite_tableName} SET {SQLite_CellName} = ""{SQLite_CellName_value}"" WHERE ROWID = {rowID};"

    '    Try
    '        Using msqlite_connect As New SQLiteConnection("Data Source=" & SQLite_path & SQLite_FileName)
    '            msqlite_connect.Open()
    '            Using msqlite_command = New SQLiteCommand(SQLite_storedGrammer, msqlite_connect)
    '                msqlite_command.ExecuteNonQuery()
    '                msqlite_command.Dispose()
    '                msqlite_connect.Close()
    '            End Using
    '        End Using
    '        JobMaker_Form.ResultOutput_TextBox.Text += $"{SQLite_CellName} : {SQLite_CellName_value}成功更新{vbCrLf}"
    '    Catch e As Exception
    '        JobMaker_Form.ResultFailOutput_TextBox.Text += $"{SQLite_CellName} : {SQLite_CellName_value}失敗更新{vbCrLf}"
    '    End Try

    '----------------------- SQLite Reading -----------------------------
    'End Function


    ''' <summary>
    ''' SQLite表格中插入新的行
    ''' </summary>
    ''' <param name="SQLite_CellName">欄位</param>
    ''' <param name="SQLite_tableName">表格</param>
    ''' <param name="SQLite_Path">路徑</param>
    ''' <param name="SQLite_FileName">檔案名(包含附檔名)</param>
    ''' <returns></returns>
    Overloads Function Insert_DbmsData(SQLite_CellName As String, SQLite_tableName As String,
                                       SQLite_Path As String, SQLite_FileName As String)
        '----------------------- SQLite Reading -----------------------------
        SQLite_storedGrammer = $"INSERT INTO {SQLite_tableName} ({SQLite_CellName}) VALUES ("""");"
        Try
            Using msqlite_connect As New SQLiteConnection("Data Source=" & SQLite_Path & SQLite_FileName)
                msqlite_connect.Open()
                Using msqlite_command = New SQLiteCommand(SQLite_storedGrammer, msqlite_connect)
                    msqlite_command.ExecuteNonQuery()
                    msqlite_command.Dispose()
                    msqlite_connect.Close()
                End Using
            End Using
            'JobMaker_Form.ResultOutput_TextBox.Text += $"{SQLite_CellName} : {SQLite_CellName_value}成功更新{vbCrLf}"
        Catch e As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{SQLite_CellName} : 插入空值 失敗更新{vbCrLf}"
        End Try
        '----------------------- SQLite Reading -----------------------------
    End Function

    ''' <summary>
    ''' 讀取SQLite單一欄位資料
    ''' </summary>
    ''' <param name="selectName">讀取的欄位</param>
    ''' <param name="tableName">表格</param>
    ''' <param name="SQLite_Path">路徑</param>
    ''' <param name="SQLite_FileName">檔案名(包含附檔名)</param>
    ''' <returns></returns>
    Overloads Function read_DbmsData(selectName As String, tableName As String, SQLite_Path As String, SQLite_FileName As String) As String
        '----------------------- SQLite Reading -----------------------------
        Dim read_string As String
        Try
            Using msqlite_connect As New SQLiteConnection("Data Source=" & SQLite_Path & SQLite_FileName)
                msqlite_connect.Open()
                Using msqlite_command = New SQLiteCommand("SELECT * FROM " & tableName, msqlite_connect)
                    Using msqlite_dataReader As SQLiteDataReader = msqlite_command.ExecuteReader
                        'sqlite_dataReader = msqlite_command.ExecuteReader()
                        While msqlite_dataReader.Read
                            read_string = msqlite_dataReader(selectName).ToString()
                            If read_string <> "" Then
                                Return read_string
                            End If
                        End While
                        msqlite_dataReader.Close()
                        msqlite_command.Dispose()
                        msqlite_connect.Close()
                    End Using
                End Using
            End Using
            JobMaker_Form.ResultOutput_TextBox.Text += $"{tableName} 的 {selectName} 成功讀取{vbCrLf}"
        Catch e As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{tableName} 的 {selectName} 失敗讀取{vbCrLf}"
        End Try
        '----------------------- SQLite Reading -----------------------------
    End Function

    ''' <summary>
    ''' [送狀 > Page CheckListBox > 讀取後計算有無增加或減少]
    ''' </summary>
    ''' <param name="selectName"></param>
    ''' <param name="tableName"></param>
    ''' <param name="sqlite_path"></param>
    ''' <param name="SQLite_FileName"></param>
    ''' <param name="rowid"></param>
    ''' <returns></returns>
    Overloads Function read_DbmsData_RowID(selectName As String, tableName As String,
                                           sqlite_path As String, SQLite_FileName As String,
                                           rowid As Integer)
        '----------------------- SQLite Reading -----------------------------
        Dim read_string As String
        Try
            Using msqlite_connect As New SQLiteConnection("Data Source=" & sqlite_path & SQLite_FileName)
                msqlite_connect.Open()
                Using msqlite_command = New SQLiteCommand($"SELECT * FROM {tableName} WHERE ROWID = {rowid}", msqlite_connect)
                    Using msqlite_dataReader As SQLiteDataReader = msqlite_command.ExecuteReader
                        'msqlite_dataReader = msqlite_command.ExecuteReader()

                        While msqlite_dataReader.Read
                            read_string = msqlite_dataReader(selectName).ToString()
                            If read_string <> "" Then
                                Return read_string
                            End If
                        End While
                        msqlite_dataReader.Close()
                        msqlite_command.Dispose()
                        msqlite_connect.Close()
                    End Using
                End Using
            End Using
            JobMaker_Form.ResultOutput_TextBox.Text += $"{tableName} 的 {selectName} 成功讀取{vbCrLf}"
        Catch e As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{tableName} 的 {selectName} 失敗讀取{vbCrLf}"
        End Try
        '----------------------- SQLite Reading -----------------------------
    End Function
    Overloads Function read_DbmsData_CountRow(selectName As String, tableName As String,
                                              sqlite_path As String, SQLite_FileName As String)
        '----------------------- SQLite Reading -----------------------------
        Dim RowCount As Integer
        Try
            Using msqlite_connect As New SQLiteConnection("Data Source=" & sqlite_path & SQLite_FileName)
                msqlite_connect.Open()
                Using msqlite_command = New SQLiteCommand($"SELECT COUNT({selectName}) FROM {tableName} WHERE {selectName} <> """" ", msqlite_connect)
                    Using msqlite_dataReader As SQLiteDataReader = msqlite_command.ExecuteReader

                        msqlite_command.CommandType = CommandType.Text
                        'sqlite_dataReader = msqlite_command.ExecuteReader()

                        RowCount = (Convert.ToInt64(msqlite_command.ExecuteScalar()))
                        Return RowCount
                        msqlite_dataReader.Close()
                        msqlite_command.Dispose()
                        msqlite_connect.Close()
                    End Using
                End Using
            End Using
            JobMaker_Form.ResultOutput_TextBox.Text += $"{tableName} 的 {selectName} 成功讀取{vbCrLf}"
        Catch e As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{tableName} 的 {selectName} 失敗讀取{vbCrLf}"
        End Try
        '----------------------- SQLite Reading -----------------------------
    End Function

    ''' <summary>
    ''' 送狀 > Page CheckListBox > 讀取後新增]
    ''' </summary>
    ''' <param name="selectName"></param>
    ''' <param name="tableName"></param>
    ''' <param name="wChkListBox"></param>
    ''' <param name="sqlite_path"></param>
    ''' <param name="SQLite_FileName"></param>
    Overloads Sub read_DbmsData_catalogPage(selectName As String, tableName As String, wChkListBox As CheckedListBox,
                                            sqlite_path As String, SQLite_FileName As String)
        '----------------------- SQLite Reading -----------------------------
        Dim read_string As String
        Try
            Using msqlite_connect As New SQLiteConnection("Data Source=" & sqlite_path & SQLite_FileName)
                msqlite_connect.Open()
                Using msqlite_command = New SQLiteCommand($"SELECT * FROM " & tableName, msqlite_connect)
                    Using msqlite_dataReader As SQLiteDataReader = msqlite_command.ExecuteReader
                        'sqlite_dataReader = msqlite_command.ExecuteReader()

                        If wChkListBox.Items.Count = 0 Then
                            While msqlite_dataReader.Read
                                read_string = msqlite_dataReader(selectName).ToString()
                                If read_string <> "" Then
                                    wChkListBox.Items.Add(read_string)
                                End If
                            End While
                        End If
                        msqlite_dataReader.Close()
                        msqlite_command.Dispose()
                        msqlite_connect.Close()
                    End Using
                End Using
            End Using
            JobMaker_Form.ResultOutput_TextBox.Text += $"{tableName} 的 {selectName} 成功讀取{vbCrLf}"
        Catch e As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{tableName} 的 {selectName} 失敗讀取{vbCrLf}"
        End Try

        '----------------------- SQLite Reading -----------------------------
    End Sub

End Class
