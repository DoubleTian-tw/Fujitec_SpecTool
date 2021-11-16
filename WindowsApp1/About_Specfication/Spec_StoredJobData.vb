Imports System.Data.SQLite

Public Class Spec_StoredJobData
    Dim sqlite_connect As SQLiteConnection
    Dim sqlite_cmd As SQLiteCommand

    Dim sqlite_Transaction As SQLiteTransaction
    Dim sqlite_dataReader As SQLiteDataReader


    '[JobMaker > Load ] --------------------------------------------
    Public Load_Job_JobSelect_RadioButton As String = "Load_Job_JobSelect_RadioButton"
    Public Load_Job_ChkListSelect_RadioButton As String = "Load_Job_ChkListSelect_RadioButton"
    Public Load_Job_JobSearch_TextBox As String = "Load_Job_JobSearch_TextBox"
    Public Load_Job_OutputPath_TextBox As String = "Load_Job_OutputPath_TextBox"
    Public Load_Job_BasePath_ComboBox As String = "Load_Job_BasePath_ComboBox"
    Public Load_AutoLoad_Loading_CheckBox As String = "Load_AutoLoad_Loading_CheckBox"
    '-------------------------------------------- [JobMaker > Load ] 

    '[JobMaker > 基本 ] --------------------------------------------
    Public Basic_Use_ChkBox As String = "Basic_Use_ChkBox"
    Public Basic_Local As String = "Basic_Local"
    Public Basic_JobNo_Old As String = "Basic_JobNo_Old"
    Public Basic_JobNo_Mod As String = "Basic_JobNo_Mod"
    Public Basic_JobName As String = "Basic_JobName"
    Public Basic_JobNo_New As String = "Basic_JobNo_New"
    Public Basic_DesignerChinese As String = "Basic_DesignerChinese"
    Public Basic_DesignerEnglish As String = "Basic_DesignerEnglish"
    Public Basic_CheckerChinese As String = "Basic_CheckerChinese"
    Public Basic_CheckerEnglish As String = "Basic_CheckerEnglish"
    Public Basic_ApproverChinese As String = "Basic_ApproverChinese"
    Public Basic_ApproverEnglish As String = "Basic_ApproverEnglish"
    Public Basic_DateTimePicker As String = "Basic_DateTimePicker"
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
    Public SpecBasic_FLEX_Number As String = "SpecBasic_FLEX_Number"
    '------------------------------------------- [ JobMaker > SPEC仕樣 > Basic ] 

    '[ JobMaker > SPEC仕樣 > TW台灣]-------------------------------------------------
    Public SPEC_TW_IDU_CHKBOX As String = "SPEC_TW_IDU_CHKBOX"
    Public SPEC_TW_FP17_CHKBOX As String = "SPEC_TW_FP17_CHKBOX"
    Public SPEC_MACHINE_TYPE As String = "SPEC_MACHINE_TYPE"
    Public SPEC_AUTO_DR As String = "SPEC_AUTO_DR"
    Public SPEC_AUTO_DR_PHOTOEYE As String = "SPEC_AUTO_DR_PHOTOEYE"
    Public SPEC_AUTO_DR_PHOTOEYE_ONLY_TEXTBOX As String = "SPEC_AUTO_DR_PHOTOEYE_ONLY_TEXTBOX"
    Public SPEC_AUTO_DR_PHOTOEYE_ONLY_CHECKBOX As String = "SPEC_AUTO_DR_PHOTOEYE_ONLY_CHECKBOX"
    Public SPEC_AUTO_DR_SAFETY As String = "SPEC_AUTO_DR_SAFETY"
    Public SPEC_AUTO_DR_SAFETY_ONLY_TEXTBOX As String = "SPEC_AUTO_DR_SAFETY_ONLY_TEXTBOX"
    Public SPEC_AUTO_DR_SAFETY_ONLY_CHECKBOX As String = "SPEC_AUTO_DR_SAFETY_ONLY_CHECKBOX"
    Public SPEC_CANCELL_CALL As String = "SPEC_CANCELL_CALL"
    Public SPEC_CANCELL_CALL_SCOB As String = "SPEC_CANCELL_CALL_SCOB"
    Public SPEC_CANCELL_CALL_SCOB_ONLY_TEXTBOX As String = "SPEC_CANCELL_CALL_SCOB_ONLY_TEXTBOX"
    Public SPEC_CANCELL_CALL_SCOB_ONLY_CHECKBOX As String = "SPEC_CANCELL_CALL_SCOB_ONLY_CHECKBOX"
    Public SPEC_CANCELL_BEHIND As String = "SPEC_CANCELL_BEHIND"
    Public SPEC_LAMP_CHK As String = "SPEC_LAMP_CHK"
    'Public SPEC_EC_BOOK As String = "SPEC_EC_BOOK"
    'Public SPEC_INSTALL_BOOK As String = "SPEC_INSTALL_BOOK"
    Public SPEC_AUTO_FAN As String = "SPEC_AUTO_FAN"
    Public SPEC_AUTO_FAN_ION_WITHOUT As String = "SPEC_AUTO_FAN_ION_WITHOUT"
    Public SPEC_AUTO_FAN_ION_ONLY_TEXTBOX As String = "SPEC_AUTO_FAN_ION_ONLY_TEXTBOX"
    Public SPEC_AUTO_FAN_ION_ONLY_CHECKBOX As String = "SPEC_AUTO_FAN_ION_ONLY_CHECKBOX"
    'Public SPEC_AUTO_LIGHT As String = "SPEC_AUTO_LIGHT"
    'Public SPEC_RUN_OPEN As String = "SPEC_RUN_OPEN"
    Public SPEC_CC_CANCEL As String = "SPEC_CC_CANCEL"
    Public SPEC_AUTO_PASS As String = "SPEC_AUTO_PASS"
    Public SPEC_AUTO_PASS_ONLY_TEXTBOX As String = "SPEC_AUTO_PASS_ONLY_TEXTBOX"
    Public SPEC_AUTO_PASS_ONLY_CHECKBOX As String = "SPEC_AUTO_PASS_ONLY_CHECKBOX"
    'Public SPEC_AUTO_LEVEL As String = "SPEC_AUTO_LEVEL"
    Public SPEC_OPERATION As String = "SPEC_OPERATION"
    Public SPEC_INDEP_OPE As String = "SPEC_INDEP_OPE"
    Public SPEC_INDEP_OPE_ONLY_TEXTBOX As String = "SPEC_INDEP_OPE_ONLY_TEXTBOX"
    Public SPEC_INDEP_OPE_ONLY_CHECKBOX As String = "SPEC_INDEP_OPE_ONLY_CHECKBOX"
    Public SPEC_INDEP_OPE_CMD As String = "SPEC_INDEP_OPE_CMD"
    Public SPEC_UCMP As String = "SPEC_UCMP"
    Public SPEC_HIN_CPI As String = "SPEC_HIN_CPI"
    Public SPEC_HIN_CPI_ONLY_TEXTBOX As String = "SPEC_HIN_CPI_ONLY_TEXTBOX"
    Public SPEC_HIN_CPI_ONLY_CHECKBOX As String = "SPEC_HIN_CPI_ONLY_CHECKBOX"
    Public SPEC_FIRE_OPE As String = "SPEC_FIRE_OPE"
    Public SPEC_FIRE_OPE_SIGNAL As String = "SPEC_FIRE_OPE_SIGNAL"
    Public SPEC_FIRE_ONLY_CHECKBOX As String = "SPEC_FIRE_ONLY_CHECKBOX"
    Public SPEC_FIRE_ONLY_TEXTBOX As String = "SPEC_FIRE_ONLY_TEXTBOX"
    Public SPEC_FIREMAN As String = "SPEC_FIREMAN"
    Public SPEC_FIREMAN_ESCAPE_FL As String = "SPEC_FIREMAN_ESCAPE_FL"
    Public SPEC_FIREMAN_ONLY_CHECKBOX As String = "SPEC_FIREMAN_ONLY_CHECKBOX"
    Public SPEC_FIREMAN_ONLY_TEXTBOX As String = "SPEC_FIREMAN_ONLY_TEXTBOX"

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
    Public SPEC_CPI_FM_ONLY_CHECKBOX As String = "SPEC_CPI_FM_ONLY_CHECKBOX"
    Public SPEC_CPI_FM_ONLY_TEXTBOX As String = "SPEC_CPI_FM_ONLY_TEXTBOX"
    Public SPEC_CPI_OLT As String = "SPEC_CPI_OLT"
    Public SPEC_CPI_OLT_ONLY_CHECKBOX As String = "SPEC_CPI_OLT_ONLY_CHECKBOX"
    Public SPEC_CPI_OLT_ONLY_TEXTBOX As String = "SPEC_CPI_OLT_ONLY_TEXTBOX"
    Public SPEC_HALL_GONG As String = "SPEC_HALL_GONG"
    Public SPEC_HALL_GONG_ONLY_CHECKBOX As String = "SPEC_HALL_GONG_ONLY_CHECKBOX"
    Public SPEC_HALL_GONG_ONLY_TEXTBOX As String = "SPEC_HALL_GONG_ONLY_TEXTBOX"

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

    Public SPEC_HPI As String = "SPEC_HPI"
    Public SPEC_HPI_OLT As String = "SPEC_HPI_OLT"
    Public SPEC_HPI_MAIN As String = "SPEC_HPI_MAIN"
    Public SPEC_HPI_INDEP As String = "SPEC_HPI_INDEP"
    Public SPEC_HPI_EMER As String = "SPEC_HPI_EMER"
    Public SPEC_HPI_EMER_ONLY_CHECKBOX As String = "SPEC_HPI_EMER_ONLY_CHECKBOX"
    Public SPEC_HPI_EMER_ONLY_TEXTBOX As String = "SPEC_HPI_EMER_ONLY_TEXTBOX"

    Public SPEC_DR_HOLD As String = "SPEC_DR_HOLD"
    Public SPEC_DR_HOLD_ONLY_CHECKBOX As String = "SPEC_DR_HOLD_ONLY_CHECKBOX"
    Public SPEC_DR_HOLD_ONLY_TEXTBOX As String = "SPEC_DR_HOLD_ONLY_TEXTBOX"
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
    Public SPEC_LANDIC_ONLY_CHECKBOX As String = "SPEC_LANDIC_ONLY_CHECKBOX"
    Public SPEC_LANDIC_ONLY_TEXTBOX As String = "SPEC_LANDIC_ONLY_TEXTBOX"
    Public SPEC_MFL_RETURN As String = "SPEC_MFL_RETURN"
    Public SPEC_MFL_RETURN_ONLY_CHECKBOX As String = "SPEC_MFL_RETURN_ONLY_CHECKBOX"
    Public SPEC_MFL_RETURN_ONLY_TEXTBOX As String = "SPEC_MFL_RETURN_ONLY_TEXTBOX"
    Public SPEC_MFL_RETURN_FL As String = "SPEC_MFL_RETURN_FL"
    Public SPEC_MFL_RETURN_FL_ONLY_CHECKBOX As String = "SPEC_MFL_RETURN_FL_ONLY_CHECKBOX"
    Public SPEC_MFL_RETURN_FL_ONLY_TEXTBOX As String = "SPEC_MFL_RETURN_FL_ONLY_TEXTBOX"
    Public SPEC_VONIC As String = "SPEC_VONIC"
    Public SPEC_VONIC_ONLY_CHECKBOX As String = "SPEC_VONIC_ONLY_CHECKBOX"
    Public SPEC_VONIC_ONLY_TEXTBOX As String = "SPEC_VONIC_ONLY_TEXTBOX"
    Public SPEC_VONIC_STANDARD As String = "SPEC_VONIC_STANDARD"
    Public SPEC_ELVIC As String = "SPEC_ELVIC"
    Public SPEC_ELVIC_ONLY_CHECKBOX As String = "SPEC_ELVIC_ONLY_CHECKBOX"
    Public SPEC_ELVIC_ONLY_TEXTBOX As String = "SPEC_ELVIC_ONLY_TEXTBOX"
    Public SPEC_ELVIC_1_PARKING As String = "SPEC_ELVIC_1_PARKING"
    Public SPEC_ELVIC_1_PARKING_FL_TEXTBOX As String = "SPEC_ELVIC_1_PARKING_FL_TEXTBOX"
    Public SPEC_ELVIC_1_PARKING_FL_ONLY_CHECKBOX As String = "SPEC_ELVIC_1_PARKING_FL_ONLY_CHECKBOX"
    Public SPEC_ELVIC_1_PARKING_FL_ONLY_TEXTBOX As String = "SPEC_ELVIC_1_PARKING_FL_ONLY_TEXTBOX"
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

    Public SPEC_HLL As String = "SPEC_HLL"
    Public SPEC_HLL_ONLY_CHECKBOX As String = "SPEC_HLL_ONLY_CHECKBOX"
    Public SPEC_HLL_ONLY_TEXTBOX As String = "SPEC_HLL_ONLY_TEXTBOX"
    Public SPEC_ATT As String = "SPEC_ATT"
    Public SPEC_ATT_ONLY_CHECKBOX As String = "SPEC_ATT_ONLY_CHECKBOX"
    Public SPEC_ATT_ONLY_TEXTBOX As String = "SPEC_ATT_ONLY_TEXTBOX"
    Public SPEC_FLOOD As String = "SPEC_FLOOD"
    Public SPEC_FLOOD_FL As String = "SPEC_FLOOD_FL"
    Public SPEC_LS1M As String = "SPEC_LS1M"
    Public SPEC_LS1M_ONLY_CHECKBOX As String = "SPEC_LS1M_ONLY_CHECKBOX"
    Public SPEC_LS1M_ONLY_TEXTBOX As String = "SPEC_LS1M_ONLY_TEXTBOX"
    Public SPEC_PRU As String = "SPEC_PRU"
    Public SPEC_PRU_ONLY_CHECKBOX As String = "SPEC_PRU_ONLY_CHECKBOX"
    Public SPEC_PRU_ONLY_TEXTBOX As String = "SPEC_PRU_ONLY_TEXTBOX"
    Public SPEC_LOAD_CELL As String = "SPEC_LOAD_CELL"
    Public SPEC_LOAD_CELL_POSITION As String = "SPEC_LOAD_CELL_POSITION"
    Public SPEC_LOAD_CELL_CAR_BTM_POS_CHECKBOX As String = "SPEC_LOAD_CELL_CAR_BTM_POS_CHECKBOX"
    Public SPEC_LOAD_CELL_CAR_BTM_POS_ONLY_CHECKBOX As String = "SPEC_LOAD_CELL_CAR_BTM_POS_ONLY_CHECKBOX"
    Public SPEC_LOAD_CELL_CAR_BTM_POS_ONLY_TEXTBOX As String = "SPEC_LOAD_CELL_CAR_BTM_POS_ONLY_TEXTBOX"
    Public SPEC_LOAD_CELL_MR_POS_CHECKBOX As String = "SPEC_LOAD_CELL_MR_POS_CHECKBOX"
    Public SPEC_LOAD_CELL_MR_POS_TEXTBOX As String = "SPEC_LOAD_CELL_MR_POS_TEXTBOX"
    Public SPEC_LOAD_CELL_MR_POS_ONLY_CHECKBOX As String = "SPEC_LOAD_CELL_MR_POS_ONLY_CHECKBOX"
    Public SPEC_LOAD_CELL_MR_POS_ONLY_TEXTBOX As String = "SPEC_LOAD_CELL_MR_POS_ONLY_TEXTBOX"
    Public SPEC_WTB As String = "SPEC_WTB"
    Public SPEC_WTB_ERROR As String = "SPEC_WTB_ERROR"
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
    Public SPEC_FRONT_REAR_DR_ONLY_CHECKBOX As String = "SPEC_FRONT_REAR_DR_ONLY_CHECKBOX"
    Public SPEC_FRONT_REAR_DR_ONLY_TEXTBOX As String = "SPEC_FRONT_REAR_DR_ONLY_TEXTBOX"
    Public SPEC_EACH_STOP As String = "SPEC_EACH_STOP"
    Public SPEC_INSTALL_OPE As String = "SPEC_INSTALL_OPE"
    Public SPEC_VONICBZ As String = "SPEC_VONICBZ"
    Public SPEC_VONICBZ_ONLY_CHECKBOX As String = "SPEC_VONICBZ_ONLY_CHECKBOX"
    Public SPEC_VONICBZ_ONLY_TEXTBOX As String = "SPEC_VONICBZ_ONLY_TEXTBOX"
    'Public SPEC_FORCE_CLOSE As String = "SPEC_FORCE_CLOSE"
    Public SPEC_OPE_SW As String = "SPEC_OPE_SW"
    Public SPEC_OPE_SW_ONLY_CHECKBOX As String = "SPEC_OPE_SW_ONLY_CHECKBOX"
    Public SPEC_OPE_SW_ONLY_TEXTBOX As String = "SPEC_OPE_SW_ONLY_TEXTBOX"
    Public SPEC_OPE_SW_POS As String = "SPEC_OPE_SW_POS"
    Public SPEC_OPE_SW_INPUT As String = "SPEC_OPE_SW_INPUT"
    Public SPEC_OPE_SW_ADDRESS As String = "SPEC_OPE_SW_ADDRESS"
    '-------------------------------------------------[ JobMaker > SPEC仕樣 ]

    '[ JobMaker > 重要設定 ] -------------------------------------------------
    Public IMPORTANT_Use_ChkBox As String = "IMPORTANT_Use_ChkBox"
    Public IMPORTANT_FAN As String = "IMPORTANT_FAN"
    Public IMPORTANT_BALANCE As String = "IMPORTANT_BALANCE"
    Public IMPORTANT_WCOB As String = "IMPORTANT_WCOB"
    Public IMPORTANT_DOOR_ChkBox As String = "IMPORTANT_DOOR_ChkBox"
    Public IMPORTANT_DOOR As String = "IMPORTANT_DOOR"
    Public IMPORTANT_HIN_ALLFL_CHECKBOX As String = "IMPORTANT_HIN_ALLFL_CHECKBOX" 'not yet
    Public IMPORTANT_HIN_AUTO_CHECKBOX As String = "IMPORTANT_HIN_AUTO_CHECKBOX" 'not yet
    Public IMPORTANT_HIN_AUTO_COMBOBOX As String = "IMPORTANT_HIN_AUTO_COMBOBOX" 'not yet
    Public IMPORTANT_HIN_FL_CHECKBOX As String = "IMPORTANT_HIN_FL_CHECKBOX" 'not yet
    Public IMPORTANT_HIN_FL_COMBOBOX As String = "IMPORTANT_HIN_FL_COMBOBOX" 'not yet
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
    Public SQLite_tableName_Load As String = "LoadSetting"
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

    ''' <summary>
    ''' 進度條的資料量
    ''' </summary>
    Public loadStored_totalValue As Integer

    '更新或新建SQLite ------------------------------------------
    ''' <summary>
    ''' 新建或覆蓋檔案內容
    ''' </summary>
    ''' <param name="job_dbms">Job檔案名稱</param>
    ''' <param name="coverFile">是否覆蓋? True:是/False:否</param>
    Public Sub SQLiteUpdate_Stored(job_dbms As String, Optional coverFile As Boolean = False)
        Try
            SQLite_JobDBMS_Name = job_dbms
            coverFile_bool = coverFile

            With JobMaker_Form
                .ResultOutput_TextBox.Text = ""
            End With

            If JobMaker_Form.Load_Job_JobSelect_RadioButton.Checked Or JobMaker_Form.Load_Job_ChkListSelect_RadioButton.Checked Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"更新 「Load」 開始 ======================= {vbCrLf}{vbCrLf}")
                Load_TabPage_Stored()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================更新 「Load」 結束 {vbCrLf}{vbCrLf}")
            End If
            If JobMaker_Form.Use_Basic_CheckBox.Checked Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"更新 「基本」 開始 ======================= {vbCrLf}{vbCrLf}")
                Basic_TabPage_Stored()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================更新 「基本」 結束 {vbCrLf}{vbCrLf}")
            End If
            If JobMaker_Form.Use_ChkList_CheckBox.Checked Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"更新 「CheckList」 開始 ======================= {vbCrLf}{vbCrLf}")
                CheckList_TabPage_Stored()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================更新 「CheckList」 結束 {vbCrLf}{vbCrLf}")
            End If
            If JobMaker_Form.Use_Program_CheckBox.Checked Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"更新 「程式變更」 開始 ======================= {vbCrLf}{vbCrLf}")
                ProgramChange_TabPage_Stored()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================更新 「程式變更」 結束 {vbCrLf}{vbCrLf}")
            End If
            If JobMaker_Form.Use_SpecBasic_CheckBox.Checked Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"更新 「仕樣」 開始 ======================= {vbCrLf}{vbCrLf}")
                SpecBasic_TabPage_Stored()
                SpecTW_TabPage_Stored()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================更新 「仕樣」 結束 {vbCrLf}{vbCrLf}")
            End If
            If JobMaker_Form.Use_Imp_CheckBox.Checked Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"更新 「重要設定」 開始 ======================= {vbCrLf}{vbCrLf}")
                Important_TabPage_Stored()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"更新 「重要設定」 開始 ======================= {vbCrLf}{vbCrLf}")
                JobMaker_Form.ResultOutput_TextBox.Text += $"=======================更新 「重要設定」 結束 {vbCrLf}"
            End If
            If JobMaker_Form.Use_mmic_CheckBox.Checked Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"更新 「MMIC」 開始 ======================= {vbCrLf}{vbCrLf}")
                MMIC_TabePage_Stored()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================更新 「MMIC」 結束 {vbCrLf}{vbCrLf}")
            End If

            MsgBox($"寫入成功",, "Fine")
        Catch e As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.SQLiteUpdate_Stored")
            errorInfo.writeInfoError_InfoTxt($"寫入失敗 : {e.Message}{vbCrLf}")
            MsgBox($"寫入失敗 : {e.Message}",, "Fail")
            JobMaker_Form.ResultOutput_TextBox.Text += $"寫入失敗 : {e.Message} {vbCrLf}"
        End Try

    End Sub

    ''' <summary>
    ''' 儲存 Load TabPage 中的資料
    ''' </summary>
    Private Sub Load_TabPage_Stored()
        '仕樣書路徑 > 仕樣書
        update_DbmsData(Load_Job_JobSelect_RadioButton,
                        JobMaker_Form.Load_Job_JobSelect_RadioButton.Checked,
                        SQLite_tableName_Load,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '仕樣書路徑 > CheckList
        update_DbmsData(Load_Job_ChkListSelect_RadioButton,
                        JobMaker_Form.Load_Job_ChkListSelect_RadioButton.Checked,
                        SQLite_tableName_Load,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '仕樣書路徑 > 搜尋工番 TextBox
        update_DbmsData(Load_Job_JobSearch_TextBox,
                        JobMaker_Form.Load_Job_JobSearch_TextBox.Text,
                        SQLite_tableName_Load,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '仕樣書路徑 > 輸出目標路徑 TextBox
        update_DbmsData(Load_Job_OutputPath_TextBox,
                        JobMaker_Form.Load_Job_OutputPath_TextBox.Text,
                        SQLite_tableName_Load,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '仕樣書路徑 > 來源Excel ComboBox
        update_DbmsData(Load_Job_BasePath_ComboBox,
                        JobMaker_Form.Load_Job_BasePath_ComboBox.Text,
                        SQLite_tableName_Load,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
    End Sub

    ''' <summary>
    ''' 儲存 基本TabPage 中的資料
    ''' </summary>
    Private Sub Basic_TabPage_Stored()
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
        'Approver Chinese
        update_DbmsData(Basic_ApproverChinese,
                        JobMaker_Form.Basic_ApproverChinese_ComboBox.Text,
                        SQLite_tableName_Basic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Approver English
        update_DbmsData(Basic_ApproverEnglish,
                        JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text,
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
    Private Sub CheckList_TabPage_Stored()
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
    Private Sub ProgramChange_TabPage_Stored()
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
    Private Sub SpecBasic_TabPage_Stored()
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_LiftInfo()
        'SPEC ---------------------------------------------------
        '是否使用分頁
        update_DbmsData(SpecBasic_Use_ChkBox,
                        JobMaker_Form.Use_SpecBasic_CheckBox.Checked,
                        SQLite_tableName_SpecBasic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '機種 
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.Spec_MachineType_NumericUpDown,
                                    dyCtrlName.JobMaker_MachinTypeInfoName_Array.Count,
                                    dyCtrlName.JobMaker_MachinTypeInfoName_Array,
                                    JobMaker_Form.Spec_MachineType_Panel,
                                    SpecBasic_MachineType_Number,
                                    SQLite_tableName_SpecBasic)
        '控制方式
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.Spec_MachineType_NumericUpDown,
                                    dyCtrlName.JobMaker_ControlWayInfoName_Array.Count,
                                    dyCtrlName.JobMaker_ControlWayInfoName_Array,
                                    JobMaker_Form.Spec_ControlWay_Panel,
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
                                    dyCtrlName.JobMaker_PurposeInfoName_Array.Count,
                                    dyCtrlName.JobMaker_PurposeInfoName_Array,
                                    JobMaker_Form.Spec_Purpose_Panel,
                                    SpecBasic_Purpose_Number,
                                    SQLite_tableName_SpecBasic)
        '用途 數量
        update_DbmsData(SpecBasic_Purpose_Number,
                        JobMaker_Form.Spec_Purpose_NumericUpDown.Value,
                        SQLite_tableName_SpecBasic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'FLEX
        dynamicPanel_StoredIntoDbms(LoadStored_PanelType.SingleLayer_Panel,
                                    JobMaker_Form.Spec_FLEX_N_NumericUpDown,
                                    dyCtrlName.JobMaker_FLEXInfoName_Array.Count,
                                    dyCtrlName.JobMaker_FLEXInfoName_Array,
                                    JobMaker_Form.Spec_FLEX_N_Panel,
                                    SpecBasic_FLEX_Number,
                                    SQLite_tableName_SpecBasic)
        'FLEX 數量
        update_DbmsData(SpecBasic_FLEX_Number,
                        JobMaker_Form.Spec_FLEX_N_NumericUpDown.Value,
                        SQLite_tableName_SpecBasic,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        'panel中的號機基本資訊 -------------------------------------------------------------------
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
    Private Sub SpecTW_TabPage_Stored()
        'SPEC ---------------------------------------------------
        'IDU CheckBox
        update_DbmsData(SPEC_TW_IDU_CHKBOX,
                        JobMaker_Form.Use_SpecTWIDU_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'FP17 CheckBox
        update_DbmsData(SPEC_TW_FP17_CHKBOX,
                        JobMaker_Form.Use_SpecTWFP17_CheckBox.Checked,
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
        '開門時限自動調節-光電裝置-Only checkBox
        update_DbmsData(SPEC_AUTO_DR_PHOTOEYE_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_PhotoEye_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門時限自動調節-光電裝置-Only TextBox
        update_DbmsData(SPEC_AUTO_DR_PHOTOEYE_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_PhotoEye_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門時限自動調節-機械式裝置
        update_DbmsData(SPEC_AUTO_DR_SAFETY,
                        JobMaker_Form.Spec_MechSafety_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門時限自動調節-機械式裝置-Only checkBox
        update_DbmsData(SPEC_AUTO_DR_SAFETY_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_MechSafety_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門時限自動調節-機械式裝置-Only TextBox
        update_DbmsData(SPEC_AUTO_DR_SAFETY_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_MechSafety_Only_TextBox.Text,
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
        '取消嬉戲呼叫-副COB--Only checkBox
        update_DbmsData(SPEC_CANCELL_CALL_SCOB_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_SCOB_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '取消嬉戲呼叫-副COB-Only TextBox
        update_DbmsData(SPEC_CANCELL_CALL_SCOB_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_SCOB_Only_TextBox.Text,
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
        '風扇連動-離子除菌-Only checkBox
        update_DbmsData(SPEC_AUTO_FAN_ION_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_ION_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '風扇連動-離子除菌-Only TextBox
        update_DbmsData(SPEC_AUTO_FAN_ION_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_ION_Only_TextBox.Text,
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
        '自動滿員通過-Only checkBox
        update_DbmsData(SPEC_AUTO_PASS_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_AutoPass_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '自動滿員通過-Only TextBox
        update_DbmsData(SPEC_AUTO_PASS_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_AutoPass_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '專用運轉
        update_DbmsData(SPEC_INDEP_OPE,
                        JobMaker_Form.Spec_Indep_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '專用運轉-Only checkBox
        update_DbmsData(SPEC_INDEP_OPE_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Indep_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '專用運轉-Only TextBox
        update_DbmsData(SPEC_INDEP_OPE_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Indep_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '專用運轉-Only checkBox
        update_DbmsData(SPEC_INDEP_OPE_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Indep_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '專用運轉-Only TextBox
        update_DbmsData(SPEC_INDEP_OPE_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Indep_Only_TextBox.Text,
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
        'HIN/CPI-Only checkBox
        update_DbmsData(SPEC_HIN_CPI_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_HinCpi_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'HIN/CPI-Only TextBox
        update_DbmsData(SPEC_HIN_CPI_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_HinCpi_Only_TextBox.Text,
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
        update_DbmsData(SPEC_FIREMAN_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Fireman_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)

        '消防梯運轉-Only n 號機 TextBox
        update_DbmsData(SPEC_FIREMAN_ONLY_TEXTBOX,
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
        '車廂管制運轉燈-緊急 ONLY CHECKBOX
        update_DbmsData(SPEC_CPI_FM_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_CpiFM_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '車廂管制運轉燈-緊急 ONLY TEXTBOX
        update_DbmsData(SPEC_CPI_FM_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_CpiFM_Only_TextBox.Text,
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
        '乘場到著鈴-Only CheckBox
        update_DbmsData(SPEC_HALL_GONG_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_HallGong_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場到著鈴-Only TextBox
        update_DbmsData(SPEC_HALL_GONG_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_HallGong_Only_TextBox.Text,
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
        '乘場信號文字-緊急 Only CheckBox
        update_DbmsData(SPEC_HPI_EMER_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_HpiFM_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場信號文字-緊急 TextBox
        update_DbmsData(SPEC_HPI_EMER_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_HpiFM_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門延長按鈕
        update_DbmsData(SPEC_DR_HOLD,
                        JobMaker_Form.Spec_DrHold_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門延長按鈕-Only CheckBox
        update_DbmsData(SPEC_DR_HOLD_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_DrHold_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '開門延長按鈕-Only TextBox
        update_DbmsData(SPEC_DR_HOLD_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_DrHold_Only_TextBox.Text,
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
        'LANDIC Only CheckBox
        update_DbmsData(SPEC_LANDIC_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Landic_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'LANDIC Only TextBox
        update_DbmsData(SPEC_LANDIC_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Landic_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '基準階賦歸
        update_DbmsData(SPEC_MFL_RETURN,
                        JobMaker_Form.Spec_MFLReturn_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '基準階賦歸-Only CheckBox
        update_DbmsData(SPEC_MFL_RETURN_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_MFLReturn_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '基準階賦歸-Only CheckBox
        update_DbmsData(SPEC_MFL_RETURN_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_MFLReturn_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '基準階賦歸-基準階
        update_DbmsData(SPEC_MFL_RETURN_FL,
                        JobMaker_Form.Spec_MFLReturn_FL_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '基準階賦歸-基準階 Only CheckBox
        update_DbmsData(SPEC_MFL_RETURN_FL_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_MFLReturn_FL_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '基準階賦歸-基準階 Only TextBox
        update_DbmsData(SPEC_MFL_RETURN_FL_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_MFLReturn_FL_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '語音撥放器VONIC
        update_DbmsData(SPEC_VONIC,
                        JobMaker_Form.Spec_Vonic_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '語音撥放器VONIC-Only CheckBox
        update_DbmsData(SPEC_VONIC_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Vonic_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '語音撥放器VONIC-Only TextBox
        update_DbmsData(SPEC_VONIC_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Vonic_Only_TextBox.Text,
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
        'ELVIC Only CheckBox
        update_DbmsData(SPEC_ELVIC_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Elvic_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC Only TextBox
        update_DbmsData(SPEC_ELVIC_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Elvic_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-PARKING OPE
        update_DbmsData(SPEC_ELVIC_1_PARKING,
                        JobMaker_Form.Spec_Elvic_Parking_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-PARKING FL
        update_DbmsData(SPEC_ELVIC_1_PARKING_FL_TEXTBOX,
                        JobMaker_Form.Spec_Elvic_ParkingFL_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-PARKING FL Only CheckBox
        update_DbmsData(SPEC_ELVIC_1_PARKING_FL_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_Elvic_ParkingFL_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'ELVIC-PARKING FL Only TextBox
        update_DbmsData(SPEC_ELVIC_1_PARKING_FL_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_Elvic_ParkingFL_Only_TextBox.Text,
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
        'ELVIC-Express Service
        update_DbmsData(SPEC_ELVIC_1_EXPRESS,
                        JobMaker_Form.Spec_Elvic_Express_CheckBox.Checked,
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
        '乘場廳燈 Only CheckBox
        update_DbmsData(SPEC_HLL_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_HLL_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '乘場廳燈 Only TextBox
        update_DbmsData(SPEC_HLL_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_HLL_Only_TextBox.Text,
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
                        JobMaker_Form.Spec_WCOB_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣-Only TEXTBOX
        update_DbmsData(SPEC_WCOB_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_WCOB_Only_TextBox.Text,
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
                        JobMaker_Form.Spec_WSCOB_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '殘障仕樣-SCOB ONLY TEXTBOX
        update_DbmsData(SPEC_WSCOB_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_WSCOB_Only_TextBox.Text,
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
        '運轉手盤運轉 Only CheckBox
        update_DbmsData(SPEC_ATT_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_ATT_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '運轉手盤運轉 Only TextBox
        update_DbmsData(SPEC_ATT_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_ATT_Only_TextBox.Text,
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
        'LS1M頂部緊急停止開關 Only CheckBox
        update_DbmsData(SPEC_LS1M_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_LS1M_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'LS1M頂部緊急停止開關 Only TextBox
        update_DbmsData(SPEC_LS1M_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_LS1M_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '電力回升
        update_DbmsData(SPEC_PRU,
                        JobMaker_Form.Spec_PRU_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '電力回升 Only CheckBox
        update_DbmsData(SPEC_PRU_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_PRU_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '電力回升 Only TextBox
        update_DbmsData(SPEC_PRU_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_PRU_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell
        update_DbmsData(SPEC_LOAD_CELL,
                        JobMaker_Form.Spec_LoadCell_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell-裝置在車廂下 CheckBox
        update_DbmsData(SPEC_LOAD_CELL_CAR_BTM_POS_CHECKBOX,
                        JobMaker_Form.Spec_LoadCellPos_CarBtm_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell-裝置在車廂下 Only CheckBox
        update_DbmsData(SPEC_LOAD_CELL_CAR_BTM_POS_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_LoadCellPos_CarBtm_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell-裝置在車廂下 Only TextBox
        update_DbmsData(SPEC_LOAD_CELL_CAR_BTM_POS_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_LoadCellPos_CarBtm_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell-裝置在機房 CheckBox
        update_DbmsData(SPEC_LOAD_CELL_MR_POS_CHECKBOX,
                        JobMaker_Form.Spec_LoadCellPos_MR_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell-裝置在機房 TextBox
        update_DbmsData(SPEC_LOAD_CELL_MR_POS_TEXTBOX,
                        JobMaker_Form.Spec_LoadCellPos_MR_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell-裝置在機房 Only CheckBox
        update_DbmsData(SPEC_LOAD_CELL_MR_POS_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_LoadCellPos_MR_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'Load Cell-裝置在機房 Only CheckBox
        update_DbmsData(SPEC_LOAD_CELL_MR_POS_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_LoadCellPos_MR_Only_TextBox.Text,
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
        '正背門 Only CheckBox
        update_DbmsData(SPEC_FRONT_REAR_DR_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_FrontRearDr_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '正背門 Only TextBox
        update_DbmsData(SPEC_FRONT_REAR_DR_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_FrontRearDr_Only_TextBox.Text,
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
        'vonic蜂鳴器-Only CheckBox
        update_DbmsData(SPEC_VONICBZ_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_VonicBz_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'vonic蜂鳴器-Only TextBox
        update_DbmsData(SPEC_VONICBZ_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_VonicBz_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '單群控切換
        update_DbmsData(SPEC_OPE_SW,
                        JobMaker_Form.Spec_OpeSw_ComboBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '單群控切換 Only CheckBox
        update_DbmsData(SPEC_OPE_SW_ONLY_CHECKBOX,
                        JobMaker_Form.Spec_OpeSw_Only_CheckBox.Checked,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '單群控切換 Only TextBox
        update_DbmsData(SPEC_OPE_SW_ONLY_TEXTBOX,
                        JobMaker_Form.Spec_OpeSw_Only_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '單群控切換-裝置在
        update_DbmsData(SPEC_OPE_SW_POS,
                        JobMaker_Form.Spec_OpeSw_DevicePos_TextBox.Text,
                        SQLite_tableName_SpecTW,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        '單群控切換-入力點Position
        update_DbmsData(SPEC_OPE_SW_INPUT,
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
    Private Sub Important_TabPage_Stored()
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
        'DOOR TYPE CheckBox
        update_DbmsData(IMPORTANT_DOOR_ChkBox,
                        JobMaker_Form.Imp_DoorType_CheckBox.Checked,
                        SQLite_tableName_Important,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'DOOR TYPE
        update_DbmsData(IMPORTANT_DOOR,
                        JobMaker_Form.Imp_DoorType_TextBox.Text,
                        SQLite_tableName_Important,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'HIN-制御階 CheckBox ================================================================
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        Dim lift_i As Integer = 0
        Dim chkBox_allFL_arrayList, chkBox_autoInsert_arrayList, cmbBox_autoInsert_arrayList As New ArrayList
        Dim chkBox_eachFL_arrayList, cmbBox_eachFL_arrayList As New ArrayList

        For Each flowPanel As Control In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls
            lift_i += 1
            For Each ctrl As Control In flowPanel.Controls
                If ctrl.GetType.Name = replaceControllerName.ctrlTypeName_CheckBox Then
                    'CheckBox 全樓層都打勾
                    If ctrl.Name = $"{dyCtrlName.JobMaker_HIN_AllFL_ChkB}_{lift_i}" Then
                        chkBox_allFL_arrayList.Add(ctrl)
                    End If
                    'CheckBox 自動填入
                    If ctrl.Name = $"{dyCtrlName.JobMaker_HIN_ChoAuto_ChkB}_{lift_i}" Then
                        chkBox_autoInsert_arrayList.Add(ctrl)
                    End If
                    'CheckBox 各樓層
                    For stopFL As Integer = 1 To CInt(JobMaker_Form.arr_liftStopFL(lift_i - 1))
                        If ctrl.Name = $"{stopFL}{dyCtrlName.JobMaker_HIN_FL_ChkB}_{lift_i}" Then
                            chkBox_eachFL_arrayList.Add(ctrl)
                        End If
                    Next
                ElseIf ctrl.GetType.Name = replaceControllerName.ctrlTypeName_ComboBox Then
                    'ComboBox 自動填入
                    If ctrl.Name = $"{dyCtrlName.JobMaker_HIN_ChoAuto_CmbB}_{lift_i}" Then
                        cmbBox_autoInsert_arrayList.Add(ctrl)
                    End If
                    'ComboBox 各樓層
                    For stopFL As Integer = 1 To CInt(JobMaker_Form.arr_liftStopFL(lift_i - 1))
                        If ctrl.Name = $"{stopFL}{dyCtrlName.JobMaker_HIN_FL_CmbB}_{lift_i}" Then
                            cmbBox_eachFL_arrayList.Add(ctrl)
                        End If
                    Next
                End If
            Next ctrl
        Next flowPanel


        Dim currentFL As Integer = 0
        Dim currentLiftNum As Integer = 0

        If chkBox_eachFL_arrayList.Count >= chkBox_allFL_arrayList.Count Then
            'Each FL CheckBox
            For Each chk_eachFL As CheckBox In chkBox_eachFL_arrayList
                currentFL += 1
                Insert_DbmsData(IMPORTANT_HIN_FL_CHECKBOX,
                                SQLite_tableName_Important,
                                SQLite_connectionPath_Job,
                                SQLite_JobDBMS_Name)

                update_DbmsData(IMPORTANT_HIN_FL_CHECKBOX,
                                chk_eachFL.Checked,
                                SQLite_tableName_Important,
                                SQLite_connectionPath_Job,
                                SQLite_JobDBMS_Name,
                                currentFL)
            Next
        Else
            '全樓層都打勾 CheckBox
            For Each chk_AllFL As CheckBox In chkBox_allFL_arrayList
                currentLiftNum += 1
                Insert_DbmsData(IMPORTANT_HIN_ALLFL_CHECKBOX,
                                SQLite_tableName_Important,
                                SQLite_connectionPath_Job,
                                SQLite_JobDBMS_Name)
                update_DbmsData(IMPORTANT_HIN_ALLFL_CHECKBOX,
                                chk_AllFL.Checked,
                                SQLite_tableName_Important,
                                SQLite_connectionPath_Job,
                                SQLite_JobDBMS_Name,
                                currentLiftNum)
            Next
        End If

        'Each FL ComboBox
        currentFL = 0
        For Each cmb_eachFL As ComboBox In cmbBox_eachFL_arrayList
            currentFL += 1
            update_DbmsData(IMPORTANT_HIN_FL_COMBOBOX,
                            cmb_eachFL.Text,
                            SQLite_tableName_Important,
                            SQLite_connectionPath_Job,
                            SQLite_JobDBMS_Name,
                            currentFL)
        Next

        '全樓層都打勾 CheckBox
        For Each chk_AllFL As CheckBox In chkBox_allFL_arrayList
            currentLiftNum += 1
            update_DbmsData(IMPORTANT_HIN_ALLFL_CHECKBOX,
                            chk_AllFL.Checked,
                            SQLite_tableName_Important,
                            SQLite_connectionPath_Job,
                            SQLite_JobDBMS_Name,
                            currentLiftNum)
        Next
        '自動填入 CheckBox
        currentLiftNum = 0
        For Each chk_auto As CheckBox In chkBox_autoInsert_arrayList
            currentLiftNum += 1
            update_DbmsData(IMPORTANT_HIN_AUTO_CHECKBOX,
                            chk_auto.Checked,
                            SQLite_tableName_Important,
                            SQLite_connectionPath_Job,
                            SQLite_JobDBMS_Name,
                            currentLiftNum)
        Next
        '自動填入 ComboBox
        currentLiftNum = 0
        For Each cmb_auto As ComboBox In cmbBox_autoInsert_arrayList
            currentLiftNum += 1
            update_DbmsData(IMPORTANT_HIN_AUTO_COMBOBOX,
                            cmb_auto.Text,
                            SQLite_tableName_Important,
                            SQLite_connectionPath_Job,
                            SQLite_JobDBMS_Name,
                            currentLiftNum)
        Next
        '================================================================ HIN-制御階 CheckBox 

    End Sub
    Private Sub MMIC_TabePage_Stored()
        'MMIC CheckBox
        update_DbmsData(MMIC_Use_ChkBox,
                        JobMaker_Form.Use_mmic_CheckBox.Checked,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC 機種 ComboBox
        update_DbmsData(MMIC_MachineType,
                        JobMaker_Form.MMIC_MachineType_ComboBox.Text,
                        SQLite_tableName_MMIC,
                        SQLite_connectionPath_Job,
                        SQLite_JobDBMS_Name)
        'MMIC FLEX-N ComboBox
        update_DbmsData(MMIC_FLEX,
                        JobMaker_Form.MMIC_FLEX_N_ComboBox.Text,
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

    ''' <summary>
    ''' 計算需要載入資料的控制項數量
    ''' </summary>
    Public Sub loadStored_controllerCount()

        ctrl_distingusih_type(JobMaker_Form.Load_TabPage)
        '仕樣書路徑
        ctrl_distingusih_type(JobMaker_Form.JobPath_TabPage)
        '載入SQLite
        ctrl_distingusih_type(JobMaker_Form.LoadSQL_TabPage)

        '基本
        ctrl_distingusih_type(JobMaker_Form.Basic_TabPage)

        'CheckList
        ctrl_distingusih_type(JobMaker_Form.CheckList_TabPage)
        'CheckList
        ctrl_distingusih_type(JobMaker_Form.CheckList_FlowLayoutPanel)
        'CheckList
        ctrl_distingusih_type(JobMaker_Form.CheckList2_FlowLayoutPanel)
        'CheckList
        ctrl_distingusih_type(JobMaker_Form.CheckList3_FlowLayoutPanel)

        '程式變更
        ctrl_distingusih_type(JobMaker_Form.ProgramChange_TabPage)
        '程式變更
        ctrl_distingusih_type(JobMaker_Form.ProgramChange_FlowLayoutPanel)
        '程式變更
        ctrl_distingusih_type(JobMaker_Form.use_ProgramChg_Panel4)

        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_BasicAll_TabPage)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.SpecBasic_GroupBox)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.SpecBasic_GroupBox2)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_TW_TabPage)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_TW_FlowLayoutPanel1)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_TW_FlowLayoutPanel2)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_TW_FlowLayoutPanel3)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_TW_FlowLayoutPanel4)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_TW_FlowLayoutPanel5)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_TW_FlowLayoutPanel6)
        '仕樣
        ctrl_distingusih_type(JobMaker_Form.Spec_TW_FlowLayoutPanel7)

        '重要設定
        ctrl_distingusih_type(JobMaker_Form.Important_TabPage)

        'MMIC
        ctrl_distingusih_type(JobMaker_Form.MMIC_TabPage)
        'MMIC
        ctrl_distingusih_type(JobMaker_Form.MMIC_Panel)

        'Loading全部資料的數量 顯示在label上
        Try
            JobMaker_Form.SQLite_TotalDataLoading_Label.Text = $"/ {loadStored_totalValue}"
            LoadStored_ProgressBar_Form.SQLite_TotalDataLoading_Label.Text = $"/ {loadStored_totalValue}"
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StotredJobData.loadStored_controllerCount")
            errorInfo.writeInfoError_InfoTxt($"Loading Label 錯誤 : {ex.Message}")
        End Try
    End Sub

    Private Sub ctrl_distingusih_type(mTabPages As Control)
        Try
            For Each ctrl As Control In mTabPages.Controls
                If ctrl.GetType.Name = "GroupBox" Then
                    For Each ctrl_grp As Control In ctrl.Controls
                        ctrl_count(ctrl_grp)
                    Next
                ElseIf ctrl.GetType.Name = "Panel" Then
                    For Each ctrl_grp As Control In ctrl.Controls
                        ctrl_count(ctrl_grp)
                    Next
                Else
                    ctrl_count(ctrl)
                End If
            Next
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StotredJobData.ctrl_distingusih_type")
            errorInfo.writeInfoError_InfoTxt($"計算控制項數量錯誤 : {ex.Message}")
        End Try
    End Sub

    Private Sub ctrl_count(ctrl As Control)
        If ctrl.GetType.Name = "TextBox" Then
            loadStored_totalValue += 1
        ElseIf ctrl.GetType.Name = "CheckBox" Then
            loadStored_totalValue += 1
        ElseIf ctrl.GetType.Name = "ComboBox" Then
            loadStored_totalValue += 1
        ElseIf ctrl.GetType.Name = "RadioButton" Then
            loadStored_totalValue += 1
        ElseIf ctrl.GetType.Name = "NumericUpDown" Then
            loadStored_totalValue += 1
        End If
    End Sub

    ''' <summary>
    ''' 再次讀取Spec Page2 裡的Panel資料
    ''' </summary>
    ''' <param name="job_dbms"></param>
    Public Sub SQLiteLoading_FixBug_Stored(job_dbms As String)
        SQLite_JobDBMS_Name = job_dbms
        Dim Spec_ChkBox As String =
                read_DbmsData(SpecBasic_Use_ChkBox,
                              SQLite_tableName_SpecBasic,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
        If Spec_ChkBox = "True" Then
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"Fix Bug 讀取 「仕樣」 開始 ============== {vbCrLf}{vbCrLf}")
            SpecBasic_TabPage_Page2_Load()
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"===============Fix Bug 讀取 「仕樣」 結束 {vbCrLf}{vbCrLf}")
        End If
    End Sub
    '載入SQLite ------------------------------------------
    Public Sub SQLiteLoading_Stored(job_dbms As String)
        Try
            'ProgressBar ---------------------------------
            With JobMaker_Form
                .SQLite_Loading_PictureBox.Visible = True
                .SQLite_EachDataLoading_Label.Visible = True
                .SQLite_EachDataLoading_Label.Text = 0
                .SQLite_TotalDataLoading_Label.Visible = True
                .SQLite_LoadingText_Label.Visible = True
                .Refresh()
                .ResultOutput_TextBox.Text = ""
                .ResultFailOutput_TextBox.Text = ""
            End With
            loadStored_totalValue = 0
            loadStored_controllerCount()
            '--------------------------------- ProgressBar

            SQLite_JobDBMS_Name = job_dbms

            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                $"★-☆-★-☆-★-☆-★-☆-★-「{JobMaker_Form.Load_SQLite_JobSearch_ComboBox.Text}」-☆-★-☆-★-☆-★-☆-★-☆ {vbCrLf}{vbCrLf}")

            Dim load_job_radioBtn As String =
                read_DbmsData(Load_Job_JobSelect_RadioButton,
                              SQLite_tableName_Load,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            Dim load_chkList_radioBtn As String =
                read_DbmsData(Load_Job_ChkListSelect_RadioButton,
                              SQLite_tableName_Load,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If load_job_radioBtn = "True" Or load_chkList_radioBtn = "True" Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"讀取 「Load」 開始 ======================= {vbCrLf}{vbCrLf}")
                Load_TabPage_Load()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================讀取 「Load」 結束 {vbCrLf}{vbCrLf}")
            End If
            '---------------------------------------------
            Dim Basic_ChkBox As String =
                read_DbmsData(Basic_Use_ChkBox,
                              SQLite_tableName_Basic,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If Basic_ChkBox = "True" Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"讀取 「基本」 開始 ======================= {vbCrLf}{vbCrLf}")
                Basic_TabPage_Load()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================讀取 「基本」 結束 {vbCrLf}{vbCrLf}")
            End If
            '---------------------------------------------
            Dim CheckList_ChkBox As String =
                read_DbmsData(ChkList_Use_ChkBox,
                              SQLite_tableName_CheckList,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If CheckList_ChkBox = "True" Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"讀取 「CheckList」 開始 ======================= {vbCrLf}{vbCrLf}")
                CheckList_TabPage_Load()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================讀取 「CheckList」 結束 {vbCrLf}{vbCrLf}")
            End If
            '---------------------------------------------
            Dim Program_ChkBox As String =
                read_DbmsData(ChkList_Prgm_Use_ChkBox,
                              SQLite_tableName_Program,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If Program_ChkBox = "True" Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"讀取 「程式變更」 開始 ======================= {vbCrLf}{vbCrLf}")
                ProgramChange_TabPage_Load()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================讀取 「程式變更」 結束 {vbCrLf}{vbCrLf}")
            End If
            '---------------------------------------------

            Dim Spec_ChkBox As String =
                read_DbmsData(SpecBasic_Use_ChkBox,
                              SQLite_tableName_SpecBasic,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If Spec_ChkBox = "True" Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"讀取 「仕樣」 開始 ======================= {vbCrLf}{vbCrLf}")
                SpecBasic_TabPage_Page1_Load()
                'SpecBasic_TabPage_Page2_Load()
                SpecTW_TabPage_Load()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================讀取 「仕樣」 結束 {vbCrLf}{vbCrLf}")
            End If
            '---------------------------------------------
            Dim Imp_ChkBox As String =
                read_DbmsData(IMPORTANT_Use_ChkBox,
                              SQLite_tableName_Important,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If Imp_ChkBox = "True" Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"讀取 「重要設定」 開始 ======================= {vbCrLf}{vbCrLf}")
                Important_TabPage_Load()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================讀取 「重要設定」 結束 {vbCrLf}{vbCrLf}")
            End If
            '---------------------------------------------
            Dim MMIC_ChkBox As String =
                read_DbmsData(MMIC_Use_ChkBox,
                              SQLite_tableName_MMIC,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)
            If MMIC_ChkBox = "True" Then
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"讀取 「MMIC」 開始 ======================= {vbCrLf}{vbCrLf}")
                MMIC_TabPage_Load()
                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                                  $"=======================讀取 「MMIC」 結束 {vbCrLf}{vbCrLf}")
            End If

            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                $"★-☆-★-☆-★-☆-★-☆-★-「{JobMaker_Form.Load_SQLite_JobSearch_ComboBox.Text}」-☆-★-☆-★-☆-★-☆-★-☆ {vbCrLf}{vbCrLf}")

            Dim doneResult = MsgBox("載入成功",, "Fine")
            If doneResult = MsgBoxResult.Ok Then
                'ProgressBar -----------------------
                With JobMaker_Form
                    .SQLite_Loading_PictureBox.Visible = False
                    .SQLite_EachDataLoading_Label.Visible = False
                    .SQLite_TotalDataLoading_Label.Visible = False
                    .SQLite_LoadingText_Label.Visible = False
                End With
                '----------------------- ProgressBar 
            End If
        Catch e As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.SQLiteLoading_Stored")
            errorInfo.writeInfoError_InfoTxt($"載入失敗 : {e.Message}")
            MsgBox($"載入失敗 : {e.Message}",, "Fail")
        End Try
    End Sub

    ''' <summary>
    ''' 載入 Load TabPage 中的資料
    ''' </summary>
    Private Sub Load_TabPage_Load()
        '仕樣書路徑 > 仕樣書 RadioButton
        chkbox_and_radioBtn_checkState_when_load(Load_Job_JobSelect_RadioButton,
                                                 SQLite_tableName_Load,
                                                 JobMaker_Form.Load_Job_JobSelect_RadioButton)
        '仕樣書路徑 > CheckList RadioButton
        chkbox_and_radioBtn_checkState_when_load(Load_Job_ChkListSelect_RadioButton,
                                                 SQLite_tableName_Load,
                                                 JobMaker_Form.Load_Job_ChkListSelect_RadioButton)
        '仕樣書路徑 > 搜尋路徑 TextBox
        JobMaker_Form.Load_Job_JobSearch_TextBox.Text =
           read_DbmsData(Load_Job_JobSearch_TextBox,
                         SQLite_tableName_Load,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '仕樣書路徑 > 最後輸出路徑 TextBox
        JobMaker_Form.Load_Job_OutputPath_TextBox.Text =
           read_DbmsData(Load_Job_OutputPath_TextBox,
                         SQLite_tableName_Load,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '仕樣書路徑 > 來源Excel ComboBox
        JobMaker_Form.Load_Job_BasePath_ComboBox.Text =
           read_DbmsData(Load_Job_BasePath_ComboBox,
                         SQLite_tableName_Load,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
    End Sub
    ''' <summary>
    ''' 載入 基本 TabPage 中的資料
    ''' </summary>
    Private Sub Basic_TabPage_Load()
        'Basic Use CheckBox
        chkbox_and_radioBtn_checkState_when_load(Basic_Use_ChkBox,
                                                 SQLite_tableName_Basic,
                                                 JobMaker_Form.Use_Basic_CheckBox)

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

        'Checker Chinese
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
        'Approver Chinese
        JobMaker_Form.Basic_ApproverChinese_ComboBox.Text =
            read_DbmsData(Basic_ApproverChinese,
                          SQLite_tableName_Basic,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        'Approver English
        JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text =
            read_DbmsData(Basic_ApproverEnglish,
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
    Private Sub CheckList_TabPage_Load()
        'CheckList Use Checkbox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Use_ChkBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.Use_ChkList_CheckBox)
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
        chkbox_and_radioBtn_checkState_when_load(ChkList_PA_ChkBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_PaSheet_CheckBox)
        'OS CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_OS_ChkBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_OS_CheckBox)
        'CFM CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_CFM_ChkBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_Confirm_CheckBox)
        'ELE CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_ELE_ChkBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_Elec_CheckBox)
        'Q1 no RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q1No_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_1_no_RadioButton)
        'Q1 yes RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q1Yes_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_1_yes_RadioButton)
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
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q2No_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_2_no_RadioButton)
        'Q2 yes RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q2Yes_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_2_yes_RadioButton)
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
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q3No_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_3_no_RadioButton)
        'Q3 yes RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q3Yes_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_3_yes_RadioButton)
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
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q5No_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_5_no_RadioButton)
        'Q5 標準 RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q5Std_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_5_std_RadioButton)
        'Q5 工直 RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q5NoStd_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_5_nstd_RadioButton)
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
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q6No_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_6_no_RadioButton)
        'Q6 yes RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q6Yes_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_6_yes_RadioButton)
        'Q6 yes Check RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q6YesChk_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_6_yesChk_RadioButton)
        'Q6 yes Item RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q6YesItem_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_6_yesItem_RadioButton)
        'Q6 yes 檢驗目標 TextBox
        JobMaker_Form.ChkList_6_yes_Content_TextBox.Text =
           read_DbmsData(ChkList_Q6Yes_Content,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q7 no RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q7No_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_7_no_RadioButton)
        'Q7 yes RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q7Yes_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_7_yes_RadioButton)

        'Q7 yes 文書 Textbox
        JobMaker_Form.ChkList_7_yes1_content_TextBox.Text =
           read_DbmsData(ChkList_Q7Yes_Content,
                         SQLite_tableName_CheckList,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Q8 no RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q8No_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_8_no_RadioButton)
        'Q8 yes RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q8Yes_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_8_yes_RadioButton)
        'Q8 item RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q8Item_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_8Item_RadioButton)
        'Q9 no RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q9No_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_9_no_RadioButton)
        'Q9 yes RadioButton
        chkbox_and_radioBtn_checkState_when_load(ChkList_Q9Yes_RadioBox,
                                                 SQLite_tableName_CheckList,
                                                 JobMaker_Form.ChkList_9_yes_RadioButton)
    End Sub
    Private Sub ProgramChange_TabPage_Load()

        'Program Change Use CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_Use_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.Use_Program_CheckBox)
        '1. 變更理由 TextBox
        JobMaker_Form.PrmList_1_reason_TextBox.Text =
           read_DbmsData(ChkList_Prgm_1_reason,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '2.測試裝置CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_2_Test_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_2_test_CheckBox)
        '2.控制盤CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_2_COP_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_2_COP_CheckBox)
        '2.研修測試塔CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_2_Tower_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_2_Tower_CheckBox)
        '2.其他CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_2_Other_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_2_Other_CheckBox)
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
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_3_Debug_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_3_debug_CheckBox)
        '3.Test CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_3_Test_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_3_test_CheckBox)
        '3.Confrim CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_3_CFM_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_3_confirm_CheckBox)
        '3.Execution CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_3_EXE_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_3_excute_CheckBox)
        '3.Other CheckBox
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_3_Other_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_3_other_Checkbox)
        '3.Other TextBox
        JobMaker_Form.PrmList_3_other_TextBox.Text =
           read_DbmsData(ChkList_Prgm_3_OtherContent,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '4.1 Auto Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_1Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes1_RadioButton)
        '4.1 Auto No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_1No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no1_RadioButton)
        '4.2 Input Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_2Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes2_RadioButton)
        '4.2 Input No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_2No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no2_RadioButton)
        '4.3 Ini Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_3Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes3_RadioButton)
        '4.3 Ini No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_3No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no3_RadioButton)
        '4.4 Case Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_4Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes4_RadioButton)
        '4.4 Case No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_4No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no4_RadioButton)
        '4.5 If Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_5Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes5_RadioButton)
        '4.5 If No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_5No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no5_RadioButton)
        '4.6 Loop Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_6Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes6_RadioButton)
        '4.6 Loop No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_6No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no6_RadioButton)
        '4.7 Range Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_7Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes7_RadioButton)
        '4.7 Range No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_7No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no7_RadioButton)
        '4.8 Casting Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_8Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes8_RadioButton)
        '4.8 Casting No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_8No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no8_RadioButton)
        '4.9 0 Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_9Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes9_RadioButton)
        '4.9 0 No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_9No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no9_RadioButton)
        '4.10 Count Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_10Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes10_RadioButton)
        '4.10 Count No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_10No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no10_RadioButton)
        '4.11 Address Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_11Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes11_RadioButton)
        '4.11 Address No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_11No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no11_RadioButton)
        '4.12 Custom Yes RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_12Yes_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_yes12_RadioButton)
        '4.12 Custom No RadioBtn
        chkbox_and_radioBtn_checkState_when_load(ChkList_Prgm_4_12No_ChkBox,
                                                 SQLite_tableName_Program,
                                                 JobMaker_Form.PrmList_4_no12_RadioButton)
        '4 Content RadioBtn
        JobMaker_Form.PrmList_4_content12_TextBox.Text =
           read_DbmsData(ChkList_Prgm_4_TestContent,
                         SQLite_tableName_Program,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
    End Sub



    Private Sub DWG_Load()
        'DWG Use CheckBox
        chkbox_and_radioBtn_checkState_when_load(DWG_Use_ChkBox,
                                                 SQLite_tableName_DWG,
                                                 JobMaker_Form.Use_prk_CheckBox)

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
    Private Sub SpecBasic_TabPage_Page1_Load()

        'Spec Use CheckBox
        chkbox_and_radioBtn_checkState_when_load(SpecBasic_Use_ChkBox,
                                                 SQLite_tableName_SpecBasic,
                                                 JobMaker_Form.Use_SpecBasic_CheckBox)

        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_LiftInfo()
        '電梯總數 Textbox
        JobMaker_Form.Spec_LiftNum_NumericUpDown.Value =
        read_DbmsData(SpecBasic_LiftNumber,
                         SQLite_tableName_SpecBasic,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_LiftNum_NumericUpDown,
                                  JobMaker_Form.SpecBasic_LiftItem_Dynamic_Panel,
                                  dyCtrlName.JobMaker_LiftInfoName_Array.Count,
                                  dyCtrlName.JobMaker_LiftInfoName_Array,
                                  SQLite_tableName_SpecBasic)

    End Sub

    Private Sub SpecBasic_TabPage_Page2_Load()


        '機種 數量
        With JobMaker_Form
            .Spec_MachineType_NumericUpDown.Value = 0
            .Spec_MachineType_Panel.Controls.Clear()
            .Spec_ControlWay_Panel.Controls.Clear()

            set_numericUpDown_value_when_load(SpecBasic_MachineType_Number,
                                              SQLite_tableName_SpecBasic,
                                              .Spec_MachineType_NumericUpDown)
        End With
        '用途 數量
        With JobMaker_Form
            .Spec_Purpose_NumericUpDown.Value = 0
            .Spec_Purpose_Panel.Controls.Clear()

            set_numericUpDown_value_when_load(SpecBasic_Purpose_Number,
                                              SQLite_tableName_SpecBasic,
                                              .Spec_Purpose_NumericUpDown)
        End With
        'FLEX-N 數量
        With JobMaker_Form
            .Spec_FLEX_N_NumericUpDown.Value = 0
            .Spec_FLEX_N_Panel.Controls.Clear()
            set_numericUpDown_value_when_load(SpecBasic_FLEX_Number,
                                          SQLite_tableName_SpecBasic,
                                          .Spec_FLEX_N_NumericUpDown)
        End With

        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_LiftInfo()

        '機種 
        dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_MachineType_NumericUpDown,
                                  JobMaker_Form.Spec_MachineType_Panel,
                                  dyCtrlName.JobMaker_MachinTypeInfoName_Array.Count,
                                  dyCtrlName.JobMaker_MachinTypeInfoName_Array,
                                  SQLite_tableName_SpecBasic)
        '控制方式
        dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_MachineType_NumericUpDown,
                                  JobMaker_Form.Spec_ControlWay_Panel,
                                  dyCtrlName.JobMaker_ControlWayInfoName_Array.Count,
                                  dyCtrlName.JobMaker_ControlWayInfoName_Array,
                                  SQLite_tableName_SpecBasic)

        '用途 Textbox
        dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_Purpose_NumericUpDown,
                                  JobMaker_Form.Spec_Purpose_Panel,
                                  dyCtrlName.JobMaker_PurposeInfoName_Array.Count,
                                  dyCtrlName.JobMaker_PurposeInfoName_Array,
                                  SQLite_tableName_SpecBasic)
        'FLEX-N Textbox
        dynamicPanel_ReadFromDbms(JobMaker_Form.Spec_FLEX_N_NumericUpDown,
                                  JobMaker_Form.Spec_FLEX_N_Panel,
                                  dyCtrlName.JobMaker_FLEXInfoName_Array.Count,
                                  dyCtrlName.JobMaker_FLEXInfoName_Array,
                                  SQLite_tableName_SpecBasic)

    End Sub
    Private Sub SpecTW_TabPage_Load()
        'Spec TW Use CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_TW_IDU_CHKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Use_SpecTWIDU_CheckBox)

        chkbox_and_radioBtn_checkState_when_load(SPEC_TW_FP17_CHKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Use_SpecTWFP17_CheckBox)
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

        '開門時限自動調節-光電裝置Only Checkbox
        chkbox_and_radioBtn_checkState_when_load(SPEC_AUTO_DR_PHOTOEYE_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_PhotoEye_Only_CheckBox)

        '開門時限自動調節-光電裝置Only Textbox
        JobMaker_Form.Spec_PhotoEye_Only_TextBox.Text =
           read_DbmsData(SPEC_AUTO_DR_PHOTOEYE_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '開門時限自動調節-機械式裝置

        JobMaker_Form.Spec_MechSafety_ComboBox.Text =
           read_DbmsData(SPEC_AUTO_DR_SAFETY,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '開門時限自動調節-機械式裝置Only Checkbox
        chkbox_and_radioBtn_checkState_when_load(SPEC_AUTO_DR_SAFETY_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_MechSafety_Only_CheckBox)
        '開門時限自動調節-機械式裝置Only Textbox
        JobMaker_Form.Spec_MechSafety_Only_TextBox.Text =
           read_DbmsData(SPEC_AUTO_DR_SAFETY_ONLY_TEXTBOX,
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
        '取消嬉戲呼叫-副COB Only Checkbox
        chkbox_and_radioBtn_checkState_when_load(SPEC_CANCELL_CALL_SCOB_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_SCOB_Only_CheckBox)
        '取消嬉戲呼叫-副COB Only Textbox
        JobMaker_Form.Spec_SCOB_Only_TextBox.Text =
           read_DbmsData(SPEC_CANCELL_CALL_SCOB_ONLY_TEXTBOX,
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
        '風扇連動-離子除菌 Only Checkbox
        chkbox_and_radioBtn_checkState_when_load(SPEC_AUTO_FAN_ION_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_ION_Only_CheckBox)
        '風扇連動-離子除菌 Only Textbox
        JobMaker_Form.Spec_ION_Only_TextBox.Text =
           read_DbmsData(SPEC_AUTO_FAN_ION_ONLY_TEXTBOX,
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
        '自動滿員通過 Only Checkbox
        chkbox_and_radioBtn_checkState_when_load(SPEC_AUTO_PASS_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_AutoPass_Only_CheckBox)
        '自動滿員通過 Only Textbox
        JobMaker_Form.Spec_AutoPass_Only_TextBox.Text =
           read_DbmsData(SPEC_AUTO_PASS_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '專用運轉
        JobMaker_Form.Spec_Indep_ComboBox.Text =
           read_DbmsData(SPEC_INDEP_OPE,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '專用運轉 Only Checkbox
        chkbox_and_radioBtn_checkState_when_load(SPEC_INDEP_OPE_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Indep_Only_CheckBox)
        '專用運轉 Only Textbox
        JobMaker_Form.Spec_Indep_Only_TextBox.Text =
           read_DbmsData(SPEC_INDEP_OPE_ONLY_TEXTBOX,
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
        'HIN/CPI Only Checkbox
        chkbox_and_radioBtn_checkState_when_load(SPEC_HIN_CPI_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_HinCpi_Only_CheckBox)
        'HIN/CPI Only Textbox
        JobMaker_Form.Spec_HinCpi_Only_TextBox.Text =
           read_DbmsData(SPEC_HIN_CPI_ONLY_TEXTBOX,
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_FIRE_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Fire_Only_CheckBox)
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_FIREMAN_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Fireman_Only_CheckBox)
        '消防梯運轉-Only n 號機
        JobMaker_Form.Spec_Fireman_Only_TextBox.Text =
           read_DbmsData(SPEC_FIREMAN_ONLY_TEXTBOX,
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_PARKING_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Parking_Only_CheckBox)
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_SEISMIC_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Seismic_Only_CheckBox)
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_SEISMIC_CANCEL_SW_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_SeismicSensor_Only_CheckBox)
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_SEISMIC_CANCEL_SW_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_SeismicSW_Only_CheckBox)
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
        '車廂管制運轉燈-緊急 ONLY CHECKBOX
        chkbox_and_radioBtn_checkState_when_load(SPEC_CPI_FM_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CpiFM_Only_CheckBox)
        '車廂管制運轉燈-緊急 ONLY TEXTBOX
        JobMaker_Form.Spec_CpiFM_Only_TextBox.Text =
           read_DbmsData(SPEC_CPI_FM_ONLY_TEXTBOX,
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_CPI_OLT_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CpiOLT_Only_CheckBox)
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_CAR_GONG_CARTOP_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CarGong_Top_CheckBox)
        '車廂上到著鈴-CAR [TOP] ONLY CHECKBOX
        chkbox_and_radioBtn_checkState_when_load(SPEC_CAR_GONG_CARTOP_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CarGong_Top_Only_CheckBox)
        '車廂上到著鈴-CAR [TOP] ONLY TEXTBOX
        JobMaker_Form.Spec_CarGong_Top_Only_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_CARTOP_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)


        '車廂上到著鈴-CAR [TOP BTM] CHECKBOX
        chkbox_and_radioBtn_checkState_when_load(SPEC_CAR_GONG_CARTOPBTM_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CarGong_TopBtm_CheckBox)
        '車廂上到著鈴-CAR [TOP BTM] ONLY CHECKBOX
        chkbox_and_radioBtn_checkState_when_load(SPEC_CAR_GONG_CARTOPBTM_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CarGong_TopBtm_Only_CheckBox)
        '車廂上到著鈴-CAR [TOP BTM] ONLY TEXTBOX
        JobMaker_Form.Spec_CarGong_TopBtm_Only_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_CARTOPBTM_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)


        '車廂上到著鈴-CAR [COB] CHECKBOX
        chkbox_and_radioBtn_checkState_when_load(SPEC_CAR_GONG_COB_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CarGong_COB_CheckBox)
        '車廂上到著鈴-CAR [COB] ONLY CHECKBOX
        chkbox_and_radioBtn_checkState_when_load(SPEC_CAR_GONG_COB_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CarGong_COB_Only_CheckBox)

        '車廂上到著鈴-CAR [COB] ONLY TEXTBOX
        JobMaker_Form.Spec_CarGong_COB_Only_TextBox.Text =
           read_DbmsData(SPEC_CAR_GONG_COB_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '車廂上到著鈴-CAR [VONIC] CHECKBOX
        chkbox_and_radioBtn_checkState_when_load(SPEC_CAR_GONG_VONIC_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CarGong_VONIC_CheckBox)
        '車廂上到著鈴-CAR [VONIC] ONLY CHECKBOX
        chkbox_and_radioBtn_checkState_when_load(SPEC_CAR_GONG_VONIC_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_CarGong_VONIC_Only_CheckBox)
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
        '乘場到著鈴 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_HALL_GONG_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_HallGong_Only_CheckBox)
        '乘場到著鈴 Only TextBox
        JobMaker_Form.Spec_HallGong_Only_TextBox.Text =
           read_DbmsData(SPEC_HALL_GONG_ONLY_TEXTBOX,
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
        '乘場信號文字-緊急 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_HPI_EMER_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_HpiFM_Only_CheckBox)
        '乘場信號文字-緊急 Only TextBox
        JobMaker_Form.Spec_HpiFM_Only_TextBox.Text =
           read_DbmsData(SPEC_HPI_EMER_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        '開門延長按鈕
        JobMaker_Form.Spec_DrHold_ComboBox.Text =
           read_DbmsData(SPEC_DR_HOLD,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '開門延長按鈕 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_DR_HOLD_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_DrHold_Only_CheckBox)
        '開門延長按鈕 Only TextBox
        JobMaker_Form.Spec_DrHold_Only_TextBox.Text =
           read_DbmsData(SPEC_DR_HOLD_ONLY_TEXTBOX,
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
        '------------------------------------------------------------------- 自家發-自動產生群組項目 基本資訊

        'LANDIC
        JobMaker_Form.Spec_Landic_ComboBox.Text =
           read_DbmsData(SPEC_LANDIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'LANDIC Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_LANDIC_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Landic_Only_CheckBox)
        'LANDIC Only TextBox
        JobMaker_Form.Spec_Landic_Only_TextBox.Text =
           read_DbmsData(SPEC_LANDIC_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '基準階賦歸
        JobMaker_Form.Spec_MFLReturn_ComboBox.Text =
           read_DbmsData(SPEC_MFL_RETURN,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '基準階賦歸 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_MFL_RETURN_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_MFLReturn_Only_CheckBox)
        '基準階賦歸 Only TextBox
        JobMaker_Form.Spec_MFLReturn_Only_TextBox.Text =
           read_DbmsData(SPEC_MFL_RETURN_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '基準階賦歸-基準階
        JobMaker_Form.Spec_MFLReturn_FL_TextBox.Text =
           read_DbmsData(SPEC_MFL_RETURN_FL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '基準階賦歸-基準階 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_MFL_RETURN_FL_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_MFLReturn_FL_Only_CheckBox)
        '基準階賦歸-基準階 Only TextBox
        JobMaker_Form.Spec_MFLReturn_FL_Only_TextBox.Text =
           read_DbmsData(SPEC_MFL_RETURN_FL_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '語音撥放器VONIC
        JobMaker_Form.Spec_Vonic_ComboBox.Text =
           read_DbmsData(SPEC_VONIC,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '語音撥放器VONIC Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_VONIC_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Vonic_Only_CheckBox)
        '語音撥放器VONIC Only TextBox
        JobMaker_Form.Spec_Vonic_Only_TextBox.Text =
           read_DbmsData(SPEC_VONIC_ONLY_TEXTBOX,
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
        'ELVIC Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Only_CheckBox)
        'ELVIC Only TextBox
        JobMaker_Form.Spec_Elvic_Only_TextBox.Text =
           read_DbmsData(SPEC_ELVIC_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'ELVIC-PARKING OPE
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_1_PARKING,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Parking_CheckBox)
        'ELVIC-PARKING FL
        JobMaker_Form.Spec_Elvic_ParkingFL_TextBox.Text =
           read_DbmsData(SPEC_ELVIC_1_PARKING_FL_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'ELVIC-PARKING FL Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_1_PARKING_FL_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_ParkingFL_Only_CheckBox)
        'ELVIC-PARKING FL Only TextBox
        JobMaker_Form.Spec_Elvic_ParkingFL_Only_TextBox.Text =
           read_DbmsData(SPEC_ELVIC_1_PARKING_FL_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'ELIVC-FLOOR LOCK OUT
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_1_FL_LOCKOUT,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_FloorLockOut_CheckBox)
        'ELVIC-VIP OPE
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_1_VIP,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_VIP_CheckBox)
        'ELVIC-Express Service
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_1_EXPRESS,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Express_CheckBox)
        'ELVIC-INDEPENDENT OPE
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_1_INDEP,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Indep_CheckBox)
        'ELVIC-RETURN TO DESIGNATED FLOOR
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_1_RETURN,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_ReturnFL_CheckBox)
        'ELVIC-CHANGE TRAFFIC PATTERN
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_2_TRAFFIC,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Traffic_Peak_CheckBox)
        'ELVIC-UP PEAK
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_2_TRAFFIC_UPPEAK,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Traffic_UpPeak_CheckBox)
        'ELVIC-DOWN PEAK
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_2_TRAFFIC_DNPEAK,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Traffic_DownPeak_CheckBox)
        'ELVIC-LUNCH TIME 
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_2_TRAFFIC_LUNCH,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Traffic_Lunch_CheckBox)
        'ELVIC-CHANGE MAIN FLOOR
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_2_MFL,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_MainFL_CheckBox)
        'ELVIC-ZONING FOR EXPRESS OPE
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_2_ZONING_EXPRESS,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Zoning_CheckBox)
        'ELVIC-FLOOR LOCK OUT
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_2_FL_LOCKOUT,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_FloorLockOut_GR_CheckBox)
        'ELVIC-CAR CALL DISCONNECT
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_2_CARCALL,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_CarCall_CheckBox)
        'ELVIC-FIRE OPE. COMMAND
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_3_FIRE,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Fire_CheckBox)
        
        'ELVIC-WAVIC OPE. COMMAND
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_3_WAVIC,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_Wavic_CheckBox)

        'ELVIC-CARE READER COMMAND
        chkbox_and_radioBtn_checkState_when_load(SPEC_ELVIC_3_CARD,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_Elvic_CRD_CheckBox)
        '乘場廳燈
        JobMaker_Form.Spec_HLL_ComboBox.Text =
           read_DbmsData(SPEC_HLL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '乘場廳燈 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_HLL_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_HLL_Only_CheckBox)
        '乘場廳燈 Only TextBox
        JobMaker_Form.Spec_HLL_Only_TextBox.Text =
           read_DbmsData(SPEC_HLL_ONLY_TEXTBOX,
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
        chkbox_and_radioBtn_checkState_when_load(SPEC_WCOB_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_WCOB_Only_CheckBox)
        '殘障仕樣-Only TEXTBOX
        JobMaker_Form.Spec_WCOB_Only_TextBox.Text =
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
        JobMaker_Form.Spec_WSCOB_Only_CheckBox.Checked =
           read_DbmsData(SPEC_WSCOB_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '殘障仕樣-SCOB ONLY TEXTBOX
        JobMaker_Form.Spec_WSCOB_Only_TextBox.Text =
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
        '運轉手盤運轉 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_ATT_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_ATT_Only_CheckBox)
        '運轉手盤運轉 Only TextBox
        JobMaker_Form.Spec_ATT_Only_TextBox.Text =
           read_DbmsData(SPEC_ATT_ONLY_TEXTBOX,
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
        'LS1M頂部緊急停止開關 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_LS1M_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_LS1M_Only_CheckBox)
        'LS1M頂部緊急停止開關 Only TextBox
        JobMaker_Form.Spec_LS1M_Only_TextBox.Text =
           read_DbmsData(SPEC_LS1M_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '電力回升
        JobMaker_Form.Spec_PRU_ComboBox.Text =
           read_DbmsData(SPEC_PRU,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '電力回升 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_PRU_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_PRU_Only_CheckBox)
        '電力回升 Only TextBox
        JobMaker_Form.Spec_PRU_Only_TextBox.Text =
           read_DbmsData(SPEC_PRU_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell
        JobMaker_Form.Spec_LoadCell_ComboBox.Text =
           read_DbmsData(SPEC_LOAD_CELL,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell-車廂下 CheckBox
        JobMaker_Form.Spec_LoadCellPos_CarBtm_CheckBox.Checked =
           read_DbmsData(SPEC_LOAD_CELL_CAR_BTM_POS_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell-車廂下 Only CheckBox
        JobMaker_Form.Spec_LoadCellPos_CarBtm_Only_CheckBox.Checked =
           read_DbmsData(SPEC_LOAD_CELL_CAR_BTM_POS_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell-車廂下 Only TextBox
        JobMaker_Form.Spec_LoadCellPos_CarBtm_Only_TextBox.Text =
           read_DbmsData(SPEC_LOAD_CELL_CAR_BTM_POS_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell-機房 CheckBox
        JobMaker_Form.Spec_LoadCellPos_MR_CheckBox.Checked =
           read_DbmsData(SPEC_LOAD_CELL_MR_POS_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell-機房 TextBox
        JobMaker_Form.Spec_LoadCellPos_MR_TextBox.Text =
           read_DbmsData(SPEC_LOAD_CELL_MR_POS_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell-機房 Only CheckBox
        JobMaker_Form.Spec_LoadCellPos_MR_Only_CheckBox.Checked =
           read_DbmsData(SPEC_LOAD_CELL_MR_POS_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'Load Cell-機房 Only TextBox
        JobMaker_Form.Spec_LoadCellPos_MR_Only_TextBox.Text =
           read_DbmsData(SPEC_LOAD_CELL_MR_POS_ONLY_TEXTBOX,
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
        '正背門 Only CheckBox
        JobMaker_Form.Spec_FrontRearDr_Only_CheckBox.Checked =
           read_DbmsData(SPEC_FRONT_REAR_DR_ONLY_CHECKBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '正背門 Only TextBox
        JobMaker_Form.Spec_FrontRearDr_Only_TextBox.Text =
           read_DbmsData(SPEC_FRONT_REAR_DR_ONLY_TEXTBOX,
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
        'vonic蜂鳴器 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_VONICBZ_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_VonicBz_Only_CheckBox)
        'vonic蜂鳴器 Only TextBox
        JobMaker_Form.Spec_VonicBz_Only_TextBox.Text =
           read_DbmsData(SPEC_VONICBZ_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '單群控切換
        JobMaker_Form.Spec_OpeSw_ComboBox.Text =
           read_DbmsData(SPEC_OPE_SW,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '單群控切換 Only CheckBox
        chkbox_and_radioBtn_checkState_when_load(SPEC_OPE_SW_ONLY_CHECKBOX,
                                                 SQLite_tableName_SpecTW,
                                                 JobMaker_Form.Spec_OpeSw_Only_CheckBox)
        '單群控切換 Only TextBox
        JobMaker_Form.Spec_OpeSw_Only_TextBox.Text =
           read_DbmsData(SPEC_OPE_SW_ONLY_TEXTBOX,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '單群控切換-裝置在
        JobMaker_Form.Spec_OpeSw_DevicePos_TextBox.Text =
           read_DbmsData(SPEC_OPE_SW_POS,
                         SQLite_tableName_SpecTW,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        '單群控切換-入力點Position
        JobMaker_Form.Spec_OpeSw_InputPos_ComboBox.Text =
           read_DbmsData(SPEC_OPE_SW_INPUT,
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
    Private Sub Important_TabPage_Load()
        'IDU CheckBox
        chkbox_and_radioBtn_checkState_when_load(IMPORTANT_Use_ChkBox,
                                                 SQLite_tableName_Important,
                                                 JobMaker_Form.Use_Imp_CheckBox)
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

        'DOOR TYPE CheckBox
        chkbox_and_radioBtn_checkState_when_load(IMPORTANT_DOOR_ChkBox,
                                                 SQLite_tableName_Important,
                                                 JobMaker_Form.Imp_DoorType_CheckBox)
        'DOOR TYPE
        JobMaker_Form.Imp_DoorType_TextBox.Text =
           read_DbmsData(IMPORTANT_DOOR,
                         SQLite_tableName_Important,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)

        'Hall Indicator中的號機基本資訊 -------------------------------------------------------------------
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        Dim lift_i As Integer = 0
        Dim chkBox_allFL_arrayList, chkBox_autoInsert_arrayList, cmbBox_autoInsert_arrayList As New ArrayList
        Dim chkBox_eachFL_arrayList, cmbBox_eachFL_arrayList As New ArrayList

        For Each flowPanel As Control In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls
            lift_i += 1
            For Each ctrl As Control In flowPanel.Controls
                If ctrl.GetType.Name = replaceControllerName.ctrlTypeName_CheckBox Then
                    'CheckBox 全樓層都打勾
                    If ctrl.Name = $"{dyCtrlName.JobMaker_HIN_AllFL_ChkB}_{lift_i}" Then
                        chkBox_allFL_arrayList.Add(ctrl)
                    End If
                    'CheckBox 自動填入
                    If ctrl.Name = $"{dyCtrlName.JobMaker_HIN_ChoAuto_ChkB}_{lift_i}" Then
                        chkBox_autoInsert_arrayList.Add(ctrl)
                    End If
                    'CheckBox 各樓層
                    For stopFL As Integer = 1 To CInt(JobMaker_Form.arr_liftStopFL(lift_i - 1))
                        If ctrl.Name = $"{stopFL}{dyCtrlName.JobMaker_HIN_FL_ChkB}_{lift_i}" Then
                            chkBox_eachFL_arrayList.Add(ctrl)
                        End If
                    Next
                ElseIf ctrl.GetType.Name = replaceControllerName.ctrlTypeName_ComboBox Then
                    'ComboBox 自動填入
                    If ctrl.Name = $"{dyCtrlName.JobMaker_HIN_ChoAuto_CmbB}_{lift_i}" Then
                        cmbBox_autoInsert_arrayList.Add(ctrl)
                    End If
                    'ComboBox 各樓層
                    For stopFL As Integer = 1 To CInt(JobMaker_Form.arr_liftStopFL(lift_i - 1))
                        If ctrl.Name = $"{stopFL}{dyCtrlName.JobMaker_HIN_FL_CmbB}_{lift_i}" Then
                            cmbBox_eachFL_arrayList.Add(ctrl)
                        End If
                    Next
                End If
            Next ctrl
        Next flowPanel


        'CheckBox 全樓層都打勾
        Dim currentLiftNum As Integer = 0
        For Each chk_allFL As CheckBox In chkBox_allFL_arrayList
            currentLiftNum += 1
            chk_allFL.Checked = read_DbmsData_RowID(IMPORTANT_HIN_ALLFL_CHECKBOX,
                                                    SQLite_tableName_Important,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    currentLiftNum)
        Next
        'CheckBox 自動填入
        currentLiftNum = 0
        For Each chk_auto As CheckBox In chkBox_autoInsert_arrayList
            currentLiftNum += 1
            chk_auto.Checked = read_DbmsData_RowID(IMPORTANT_HIN_AUTO_CHECKBOX,
                                                    SQLite_tableName_Important,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    currentLiftNum)
        Next
        'ComboBox 自動填入
        currentLiftNum = 0
        For Each cmb_auto As ComboBox In cmbBox_autoInsert_arrayList
            currentLiftNum += 1
            cmb_auto.Text = read_DbmsData_RowID(IMPORTANT_HIN_AUTO_COMBOBOX,
                                                    SQLite_tableName_Important,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    currentLiftNum)
        Next

        'CheckBox 各樓層
        Dim currentFL As Integer = 0
        For Each chk_eachFL As CheckBox In chkBox_eachFL_arrayList
            currentFL += 1
            chk_eachFL.Checked = read_DbmsData_RowID(IMPORTANT_HIN_FL_CHECKBOX,
                                                    SQLite_tableName_Important,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    currentFL)
        Next
        'ComboBox 各樓層
        currentFL = 0
        For Each chk_eachFL As ComboBox In cmbBox_eachFL_arrayList
            currentFL += 1
            chk_eachFL.Text = read_DbmsData_RowID(IMPORTANT_HIN_FL_COMBOBOX,
                                                    SQLite_tableName_Important,
                                                    SQLite_connectionPath_Job,
                                                    SQLite_JobDBMS_Name,
                                                    currentFL)
        Next

        'All ComboBox
        'Dim dyCtrlName As DynamicControlName = New DynamicControlName
        'dyCtrlName.JobMaker_HINInfo()

        'If JobMaker_Form.Spec_LiftNum_NumericUpDown.Value <> 0 Then
        '    For lift_i As Integer = 1 To CInt(JobMaker_Form.Spec_LiftNum_NumericUpDown.Value)
        '        If coverFile_bool = False Then
        '            If lift_i < JobMaker_Form.Spec_LiftNum_NumericUpDown.Value Then
        '                Insert_DbmsData(dyCtrlName.JobMaker_HINInfoName_Array(0),
        '                                SQLite_tableName_Important,
        '                                SQLite_connectionPath_Job,
        '                                SQLite_JobDBMS_Name)
        '            End If
        '            For Each mFlowLayoutPanel As Control In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls
        '                For ctrl_j As Integer = 1 To dyCtrlName.JobMaker_HINInfoName_Array.Count
        '                    If mFlowLayoutPanel.Name = $"{dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1)}_{lift_i}" Then
        '                        update_DbmsData(dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1),
        '                                        mFlowLayoutPanel.Text,
        '                                        SQLite_tableName_Important,
        '                                        SQLite_connectionPath_Job,
        '                                        SQLite_JobDBMS_Name,
        '                                        lift_i)
        '                    ElseIf dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1) = dyCtrlName.JobMaker_HIN_FL_ChkB Or
        '                           dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1) = dyCtrlName.JobMaker_HIN_FL_CmbB Then

        '                        For stopFL_k As Integer = 1 To JobMaker_Form.arr_liftStopFL(lift_i - 1)
        '                            If mFlowLayoutPanel.Name = $"{stopFL_k}{dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1)}_{lift_i}" Then
        '                                update_DbmsData(dyCtrlName.JobMaker_HINInfoName_Array(ctrl_j - 1),
        '                                                mFlowLayoutPanel.Text,
        '                                                SQLite_tableName_Important,
        '                                                SQLite_connectionPath_Job,
        '                                                SQLite_JobDBMS_Name,
        '                                                lift_i)
        '                            End If
        '                        Next
        '                    End If
        '                Next
        '            Next
        '        Else
        '            Dim temp_specBasic_liftNumber As String
        '            temp_specBasic_liftNumber =
        '                read_DbmsData(SpecBasic_LiftNumber,
        '                              SQLite_tableName_SpecBasic,
        '                              SQLite_connectionPath_Job,
        '                              SQLite_JobDBMS_Name)
        '            Dim overwrite_liftNumber_bool As Boolean


        '            If temp_specBasic_liftNumber <> JobMaker_Form.Spec_LiftNum_NumericUpDown.Value Then
        '                '比對電梯總數不相同，需要更改
        '                overwrite_liftNumber_bool = True

        '                '如果新的電梯數量比舊的多，則要插入新的行在SQLite中 ---------------------------------
        '                If JobMaker_Form.Spec_LiftNum_NumericUpDown.Value > temp_specBasic_liftNumber Then
        '                    Dim tempSub_num As Integer
        '                    tempSub_num = CInt(JobMaker_Form.Spec_LiftNum_NumericUpDown.Value) - CInt(temp_specBasic_liftNumber)
        '                    For insertRow_i = 1 To tempSub_num
        '                        Insert_DbmsData(dyCtrlName.JobMaker_LiftInfoName_Array(0),
        '                                        SQLite_tableName_SpecBasic,
        '                                        SQLite_connectionPath_Job,
        '                                        SQLite_JobDBMS_Name)
        '                    Next
        '                End If
        '                '---------------------------------如果新的數量比舊的多，則要插入新的行在SQLite中 
        '            Else
        '                '數量相同但內容不同，需要更改
        '                For Each tempCtrl As Control In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls
        '                    For hin_j As Integer = 1 To dyCtrlName.JobMaker_HINInfoName_Array.Count
        '                        If tempCtrl.Name = $"{dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1)}_{lift_i}" Then
        '                            If tempCtrl.Text <> read_DbmsData_RowID(dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1),
        '                                                                    SQLite_tableName_Important,
        '                                                                    SQLite_connectionPath_Job,
        '                                                                    SQLite_JobDBMS_Name,
        '                                                                    lift_i) Then
        '                                overwrite_liftNumber_bool = True
        '                                Exit For
        '                            Else
        '                                overwrite_liftNumber_bool = False
        '                            End If
        '                        ElseIf dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1) = dyCtrlName.JobMaker_HIN_FL_ChkB Or
        '                               dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1) = dyCtrlName.JobMaker_HIN_FL_CmbB Then
        '                            If tempCtrl.Text <> read_DbmsData_RowID(dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1),
        '                                                                    SQLite_tableName_Important,
        '                                                                    SQLite_connectionPath_Job,
        '                                                                    SQLite_JobDBMS_Name,
        '                                                                    lift_i) Then
        '                                overwrite_liftNumber_bool = True
        '                                Exit For
        '                            Else
        '                                overwrite_liftNumber_bool = False
        '                            End If
        '                        End If
        '                    Next
        '                    If overwrite_liftNumber_bool = True Then
        '                        Exit For
        '                    End If
        '                Next
        '            End If

        '            If overwrite_liftNumber_bool Then
        '                '當下更新的電梯內容與紀錄中的比較，如果有一處不同就全數刪除設="" -------
        '                If lift_i <= 1 Then
        '                    For hin_j As Integer = 1 To dyCtrlName.JobMaker_HINInfoName_Array.Count
        '                        update_DbmsData(dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1),
        '                                        "",
        '                                        SQLite_tableName_Important,
        '                                        SQLite_connectionPath_Job,
        '                                        SQLite_JobDBMS_Name)
        '                    Next
        '                End If
        '                '-------當下更新的電梯內容與紀錄中的比較，如果有一處不同就全數刪除設=""

        '                '更新新的CheckListBox ----------------------------------------------------------------
        '                For Each tempCtrl As Control In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls

        '                    For hin_j As Integer = 1 To dyCtrlName.JobMaker_HINInfoName_Array.Count
        '                        If tempCtrl.Name = $"{dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1)}_{lift_i}" Then
        '                            update_DbmsData(dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1),
        '                                            tempCtrl.Text,
        '                                            SQLite_tableName_Important,
        '                                            SQLite_connectionPath_Job,
        '                                            SQLite_JobDBMS_Name,
        '                                            lift_i)
        '                        ElseIf dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1) = dyCtrlName.JobMaker_HIN_FL_ChkB Or
        '                               dyCtrlName.JobMaker_HINInfoName_Array(hin_j - 1) = dyCtrlName.JobMaker_HIN_FL_CmbB Then
        '                            update_DbmsData(dyCtrlName.JobMaker_LiftInfoName_Array(hin_j - 1),
        '                                            tempCtrl.Text,
        '                                            SQLite_tableName_Important,
        '                                            SQLite_connectionPath_Job,
        '                                            SQLite_JobDBMS_Name,
        '                                            lift_i)
        '                        End If
        '                    Next
        '                Next
        '                '---------------------------------------------------------------- 更新新的CheckListBox 
        '            End If
        '        End If
        '    Next
        'End If
        '------------------------------------------------------------------- Hall Indicator中的號機基本資訊
    End Sub
    Private Sub MMIC_TabPage_Load()
        'MMIC CheckBox
        chkbox_and_radioBtn_checkState_when_load(MMIC_Use_ChkBox,
                                                 SQLite_tableName_MMIC,
                                                 JobMaker_Form.Use_mmic_CheckBox)
        'MMIC 機種 ComboBox
        JobMaker_Form.MMIC_MachineType_ComboBox.Text =
           read_DbmsData(MMIC_MachineType,
                         SQLite_tableName_MMIC,
                         SQLite_connectionPath_Job,
                         SQLite_JobDBMS_Name)
        'MMIC FLEX-N ComboBox
        JobMaker_Form.MMIC_FLEX_N_ComboBox.Text =
           read_DbmsData(MMIC_FLEX,
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
        With JobMaker_Form
            .MMIC_MR_NumericUpDown.Value = 0
            .MMIC_MR_Panel.Controls.Clear()

            set_numericUpDown_value_when_load(MMIC_MR_Number,
                                              SQLite_tableName_MMIC,
                                              .MMIC_MR_NumericUpDown)
        End With
        'MR EEPROM Number
        With JobMaker_Form
            .MMIC_MR_E_NumericUpDown.Value = 0
            .MMIC_MR_E_Panel.Controls.Clear()

            set_numericUpDown_value_when_load(MMIC_MR_ENumber,
                                              SQLite_tableName_MMIC,
                                              .MMIC_MR_E_NumericUpDown)
        End With
        'SV Number 
        With JobMaker_Form
            .MMIC_SV_NumericUpDown.Value = 0
            .MMIC_SV_Panel.Controls.Clear()

            set_numericUpDown_value_when_load(MMIC_SV_Number,
                                              SQLite_tableName_MMIC,
                                              .MMIC_SV_NumericUpDown)
        End With
        'SV EEPROM Number
        With JobMaker_Form
            .MMIC_SV_E_NumericUpDown.Value = 0
            .MMIC_SV_E_Panel.Controls.Clear()

            set_numericUpDown_value_when_load(MMIC_SV_ENumber,
                                              SQLite_tableName_MMIC,
                                              .MMIC_SV_E_NumericUpDown)
        End With
        'VD10 Number 
        With JobMaker_Form
            .MMIC_VD10_NumericUpDown.Value = 0
            .MMIC_VD10_Panel.Controls.Clear()

            set_numericUpDown_value_when_load(MMIC_VD10_Number,
                                              SQLite_tableName_MMIC,
                                              .MMIC_VD10_NumericUpDown)
        End With

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

    ''' <summary>
    ''' 解決numericUpDown在讀取時的bug，利用迴圈給予值，而不是直接付予
    ''' </summary>
    ''' <param name="spec_name"></param>
    ''' <param name="sqlite_tableName"></param>
    ''' <param name="numUpDown"></param>
    Private Sub set_numericUpDown_value_when_load(spec_name As String, sqlite_tableName As String,
                                             numUpDown As NumericUpDown)
        Dim num As Integer =
            read_DbmsData(spec_name,
                          sqlite_tableName,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)
        If num <> 0 Then
            For i As Integer = 1 To num
                numUpDown.Value = i
            Next i
        End If
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
    ''' [Load > 載入SQLite > 計算Loading檔案數量的Label、移動Loading gif圖檔]
    ''' </summary>
    Private Shared Sub loadingControllerState_whenLoading()
        With JobMaker_Form
            .SQLite_EachDataLoading_Label.Text = Val(JobMaker_Form.SQLite_EachDataLoading_Label.Text) + 1
            If .SQLite_Loading_PictureBox.Location.X < 0 - (.SQLite_Loading_PictureBox.Width) Then
                .SQLite_Loading_PictureBox.Location = New Point(450 + .SQLite_Loading_PictureBox.Width, .SQLite_Loading_PictureBox.Location.Y)
            Else
                .SQLite_Loading_PictureBox.Location = New Point(.SQLite_Loading_PictureBox.Location.X - 1,
                                                                .SQLite_Loading_PictureBox.Location.Y)
            End If
            If Val(.SQLite_EachDataLoading_Label.Text) Mod 10 = 0 Then
                .Refresh()
            End If
        End With
    End Sub
    ''' <summary>
    ''' 輸出文字至TextBox中，並將插入符號保持在最下方
    ''' </summary>
    ''' <param name="outputText">要輸出的文字</param>
    Public Sub outputText_toTextBox_focusOnBelow(tb As TextBox, outputText As String)
        With tb
            .Text += $"{outputText}"
            .SelectionStart = .TextLength
            .ScrollToCaret()
        End With
    End Sub

    ''' <summary>
    ''' 載入時設定CheckBox的Checked狀態
    ''' </summary>
    ''' <param name="spec_name"></param>
    ''' <param name="sqlite_tablename"></param>
    ''' <param name="chkbox"></param>
    Private Overloads Sub chkbox_and_radioBtn_checkState_when_load(spec_name As String, sqlite_tablename As String, chkbox As CheckBox)
        Dim temp_controler_state As String =
                read_DbmsData(spec_name,
                              sqlite_tablename,
                              SQLite_connectionPath_Job,
                              SQLite_JobDBMS_Name)

        If temp_controler_state <> "" Then
            chkbox.Checked = temp_controler_state
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"{spec_name} : 成功設定 : {temp_controler_state}{vbCrLf}{vbCrLf}")
        End If

        loadingControllerState_whenLoading()
    End Sub


    ''' <summary>
    ''' 載入時設定RadioButton的Checked狀態
    ''' </summary>
    ''' <param name="spec_name"></param>
    ''' <param name="sqlite_tablename"></param>
    ''' <param name="radioBtn"></param>
    Private Overloads Sub chkbox_and_radioBtn_checkState_when_load(spec_name As String, sqlite_tablename As String, radioBtn As RadioButton)
        Dim temp_controler_state As String =
            read_DbmsData(spec_name,
                          sqlite_tablename,
                          SQLite_connectionPath_Job,
                          SQLite_JobDBMS_Name)

        If temp_controler_state <> "" Then
            radioBtn.Checked = temp_controler_state
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"{spec_name} : 成功設定 : {temp_controler_state}{vbCrLf}{vbCrLf}")
        End If

        loadingControllerState_whenLoading()
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

            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"{SQLite_CellName} : {SQLite_CellName_value}成功更新{vbCrLf}{vbCrLf}")
            JobMaker_Form.Result_Loading_PictureBox.Refresh()
        Catch e As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.update_DbmsData")
            errorInfo.writeInfoError_InfoTxt($"{SQLite_CellName} : {SQLite_CellName_value} : {e.Message}")
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultFailOutput_TextBox,
                                             $"{SQLite_CellName} : {SQLite_CellName_value}失敗更新{vbCrLf}{vbCrLf}")
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
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"{SQLite_CellName} : {SQLite_CellName_value}成功更新{vbCrLf}{vbCrLf}")
            JobMaker_Form.Result_Loading_PictureBox.Refresh()

        Catch e As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.update_DbmsData")
            errorInfo.writeInfoError_InfoTxt($"{SQLite_CellName} : {SQLite_CellName_value} : {e.Message}")
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultFailOutput_TextBox,
                                              $"{SQLite_CellName} : {SQLite_CellName_value}失敗更新{vbCrLf}{vbCrLf}")

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
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.Insert_DbmsData")
            errorInfo.writeInfoError_InfoTxt($"{SQLite_CellName} : 插入空值 : {e.Message}")
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"{SQLite_CellName} : 插入空值 失敗更新{vbCrLf}{vbCrLf}")
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
                                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"{tableName} 的 {selectName} 成功讀取{vbCrLf}{vbCrLf}")
                                loadingControllerState_whenLoading()
                                Return read_string
                            End If
                        End While
                        msqlite_dataReader.Close()
                        msqlite_command.Dispose()
                        msqlite_connect.Close()
                    End Using
                End Using
            End Using
        Catch e As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.read_DbmsData")
            errorInfo.writeInfoError_InfoTxt($"{tableName} : {selectName} 失敗讀取 : {e.Message}{vbCrLf}")
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultFailOutput_TextBox,
                                              $"{tableName} 的 {selectName} 失敗讀取{vbCrLf}{vbCrLf}")
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
                                outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                    $"{tableName} 的 {selectName} 成功讀取{vbCrLf}{vbCrLf}")
                                Return read_string
                            End If
                        End While
                        msqlite_dataReader.Close()
                        msqlite_command.Dispose()
                        msqlite_connect.Close()
                    End Using
                End Using
            End Using
        Catch e As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.read_DbmsData_RowID")
            errorInfo.writeInfoError_InfoTxt($"{tableName} 的 {selectName} 失敗讀取 : {e.Message}{vbCrLf}")
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultFailOutput_TextBox,
                                              $"{tableName} 的 {selectName} 失敗讀取{vbCrLf}{vbCrLf}")
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
                        outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"{tableName} 的 {selectName} 成功讀取{vbCrLf}{vbCrLf}")
                        Return RowCount
                        msqlite_dataReader.Close()
                        msqlite_command.Dispose()
                        msqlite_connect.Close()
                    End Using
                End Using
            End Using
        Catch e As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.read_DbmsData_CountRow")
            errorInfo.writeInfoError_InfoTxt($"{tableName} 的 {selectName} 失敗讀取{e.Message}{vbCrLf}")
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultFailOutput_TextBox,
                                              $"{tableName} 的 {selectName} 失敗讀取{vbCrLf}{vbCrLf}")
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
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultOutput_TextBox,
                                              $"{tableName} 的 {selectName} 成功讀取{vbCrLf}{vbCrLf}")
        Catch e As Exception
            errorInfo.writeTitleIntoError_InfoTxt("Spec_StoredJobData.read_DbmsData_catalogPage")
            errorInfo.writeInfoError_InfoTxt($"{tableName} 的 {selectName} 失敗讀取{e.Message}{vbCrLf}")
            outputText_toTextBox_focusOnBelow(JobMaker_Form.ResultFailOutput_TextBox,
                                              $"{tableName} 的 {selectName} 失敗讀取{vbCrLf}{vbCrLf}")
        End Try

        '----------------------- SQLite Reading -----------------------------
    End Sub

End Class
