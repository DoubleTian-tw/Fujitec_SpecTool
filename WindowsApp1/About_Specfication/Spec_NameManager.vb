Imports System.Configuration 'dbms
Imports System.Data.SQLite
Imports System.IO

''' <summary>
''' 儲存所有在Tool DataBase.SQLite中的名稱
''' </summary>
Public Class Spec_NameManager

    '全部仕樣確認表 ----------------
    Public FinalCheck_Item As String = "FinalCheck_Item"
    Public FinalCheck_State As String = "FinalCheck_State"
    Public FinalCheck_Spec As String = "FinalCheck_Spec"
    '---------------- 全部仕樣確認表

    '----------- 基本內容 -----------------
    Public STD_JobNo_New As String = "例 : TW-5566-68"
    Public STD_JobNo_Old As String = "例 : EXF-9487"
    Public STD_JobNo_Mod As String = "例 : MOD-0520"
    '----------- 基本內容 -----------------


    '----------- OTHERS ------------------
    Public TB_O As String = "○"
    Public TB_X As String = "×"
    Public TB_WITH As String = "WITH"
    Public TB_WITHOUT As String = "WITHOUT"
    Public TB_NO As String = "N/O"
    Public TB_NC As String = "N/C"
    Public TB_DR_OPEN As String = "DR OPEN"
    Public TB_DR_CLOSE As String = "DR CLOSE"
    Public TB_CarTop As String = "CAR TOP"
    Public TB_CarTopBtm As String = "CAR TOP And BOTTOM"
    Public TB_WithCOB As String = "WITH COB"
    Public TB_InVONIC As String = "IN VONIC"
    '----------- OTHERS ------------------

    '[Tool_DataBase > BasicSetting] ---------------------
    Public EmployeeTwn_None As String = "EmployeeTwn_None" 'e.g 2100
    Public EmployeeTwn_Upper As String = "EmployeeTwn_Upper" 'e.g TWN2100
    Public EmployeeTwn_Lower As String = "EmployeeTwn_Lower" 'e.g twn2100
    Public AllEmployee_Type() As String = {EmployeeTwn_None, EmployeeTwn_Upper, EmployeeTwn_Lower} '(不是初始值)
    Public EmployeeRow As Integer '輸入員工工號在SQLite中是第幾行(不是初始值)

    Public EmployeeChinese As String = "EmployeeChinese"
    Public EmployeeEnglish As String = "EmployeeEnglish"
    Public ApproverChinese As String = "ApproverChinese"
    Public ApproverEnglish As String = "ApproverEnglish"
    Public Local As String = "Local"
    Public OperationType As String = "OperationType"
    Public FLEX As String = "FLEX"
    'Public PRK_Name As String = "PRK_Name"
    Public AllMachineType As String = "AllMachineType"
    Public mmicType As String = "mmicType"
    Public mmicTypeName As String = "mmicTypeName"
    Public mmicEEPROM_Base As String = "mmicEEPROM_Base"
    Public mmicEEPROM_DataName As String = "mmicEEPROM_DataName"
    Public gspType As String = "gspType"
    Public gspEEPROM_Base As String = "gspEEPROM_Base"
    Public gspEEPROM_DataName As String = "gspEEPROM_DataName"
    Public OverBalance As String = "OverBalance"
    Public IMP_HIN_FL_Content As String = "IMP_HIN_FL_Content"
    Public Spec_TW_EmerInput As String = "Spec_TW_EmerInput"
    Public Spec_TW_EmerAddress As String = "Spec_TW_EmerAddress"
    Public Spec_MachineType As String = "Spec_MachineType"
    Public Spec_ControlWay As String = "Spec_ControlWay"
    Public Spec_Purpose As String = "Spec_Purpose"
    '--------------------- [Tool_DataBase > BasicSetting]

    '[Tool_DataBase > GSP_ProgramType] --------------------------------

    Public gsp_N100_PC8 As String = "gsp_N100_PC8"
    Public gsp_N100_PC9 As String = "gsp_N100_PC9"
    Public gsp_OverN200 As String = "gsp_OverN200"
    Public gsp_ELVIC_TW As String = "gsp_ELVIC_TW"
    Public gsp_GsoTo1Car As String = "gsp_GsoTo1Car"
    Public gsp_EvaucationOpe_SP As String = "gsp_EvaucationOpe_SP"
    Public gsp_IndepPowerOpe As String = "gsp_IndepPowerOpe"
    Public gsp_EOP As String = "gsp_EOP"
    Public gsp_Double2Car As String = "gsp_Double2Car"
    '[MMIC > FLEX-N幾百 Combobox] --------------------------
    Public FLEX_NX100_PC8 As String = "FLEX-N(X)100(PC8)"
    Public FLEX_NX100_PC9 As String = "FLEX-N(X)100(PC9)"
    Public FLEX_NX200 As String = "FLEX-N(X)200"
    Public FLEX_NX300 As String = "FLEX-N(X)300"
    '--------------------------- [MMIC > FLEX-N幾百 Combobox]
    '[MMIC > SV BASE ComboBox] ---------------------------------------------
    Public FLEX_NX100_PC8_FileName As String = "FLEX-N(X)100(PC8)_FileName"
    Public FLEX_NX100_PC9_FileName As String = "FLEX-N(X)100(PC9)_FileName"
    Public FLEX_NX200_FileName As String = "FLEX-N(X)200_FileName"
    Public FLEX_NX300_FileName As String = "FLEX-N(X)300_FileName"
    '--------------------------------------------- [MMIC > SV BASE ComboBox] 
    '-------------------------------- [Tool_DataBase > GSP_ProgramType]

    '[Tool_DataBase > GSP_ProgramTypeName] ------------------------------------
    '[SV > FlashRom Object Name > Type TextBox] ----------------
    Public gspTypeName_Array As String = "gspTypeName_Array"
    Public gspName_N100_PC8 As String = "gspName_N100(PC8)"
    Public gspName_N100_PC9 As String = "gspName_N100(PC9)"
    Public gspName_OverN200 As String = "gspName_OverN200"
    Public gspName_Double2Car As String = "gspName_Double2Car"
    Public gspName_ELVIC_TW As String = "gspName_ELVIC_TW"
    Public gspName_GsoTo1Car As String = "gspName_GsoTo1Car"
    Public gspName_EvaucationOpe_SP As String = "gspName_EvaucationOpe_SP"
    Public gspName_IndependentPowerOpe As String = "gspName_IndependentPowerOpe"
    Public gspName_EOP As String = "gspName_EOP"
    '---------------- [SV > FlashRom Object Name > Type TextBox] 
    '------------------------------------ [Tool_DataBase > GSP_ProgramTypeName] 

    '[Tool_DataBase > MMIC_ProgramType] ------------------------------------
    Public mmic_IDU_ZT_TW As String = "mmic_IDU_ZT_TW"
    Public mmic_IDU_ZT_TW_PC8 As String = "mmic_IDU_ZT_TW_PC8"
    Public mmic_IDU_RT_TW As String = "mmic_IDU_RT_TW"
    Public mmic_FP17_ZR_TW As String = "mmic_FP17_ZR_TW"
    Public mmic_FP17_ZR_TW_PC8 As String = "mmic_FP17_ZR_TW_PC8"
    Public mmic_FP17_ZR_TW_FrontRearDoor As String = "mmic_FP17_ZR_TW_FrontRearDoor"
    Public mmic_FP17_ZR_HK As String = "mmic_FP17_ZR_HK"
    Public mmic_GLVF_HK_Hallbus As String = "mmic_GLVF_HK_Hallbus"
    Public mmic_GLVF_HK_SelcomDoor As String = "mmic_GLVF_HK_SelcomDoor"
    Public mmic_GLVF_E_SP As String = "mmic_GLVF_E_SP"
    Public mmic_REXIAa_TW As String = "mmic_REXIAa_TW"
    Public mmic_TP09_TW As String = "mmic_TP09_TW"
    Public mmic_XIOR_TW As String = "mmic_XIOR_TW"
    Public mmic_GLVF_HK_Millnet As String = "mmic_GLVF_HK_Millnet"
    Public mmic_GLVF_D_SP As String = "mmic_GLVF_D_SP"
    '------------------------------------ [Tool_DataBase > MMIC_ProgramType] 

    '[Tool_DataBase > MMIC_ProgramTypeName] ------------------------------------
    Public mmicN_IDU_ZT_TW As String = "mmicN_IDU_ZT_TW"
    Public mmicN_IDU_RT_TW As String = "mmicN_IDU_RT_TW"
    Public mmicN_FP17_ZR_TW As String = "mmicN_FP17_ZR_TW"
    Public mmicN_FP17_ZR_TW_FrontRearDoor As String = "mmicN_FP17_ZR_TW_FrontRearDoor"
    Public mmicN_FP17_ZR_HK As String = "mmicN_FP17_ZR_HK"
    Public mmicN_GLVF_HK_Hallbus As String = "mmicN_GLVF_HK_Hallbus"
    Public mmicN_GLVF_HK_SelcomDoor As String = "mmicN_GLVF_HK_SelcomDoor"
    Public mmicN_GLVF_E_SP As String = "mmicN_GLVF_E_SP"
    Public mmicN_REXIAa_TW As String = "mmicN_REXIAa_TW"
    Public mmicN_TP09_TW As String = "mmicN_TP09_TW"
    Public mmicN_XIOR_TW As String = "mmicN_XIOR_TW"
    Public mmicN_GLVF_HK_Millnet As String = "mmicN_GLVF_HK_Millnet"
    Public mmicN_GLVF_D_SP As String = "mmicN_GLVF_D_SP"
    '------------------------------------ [Tool_DataBase > MMIC_ProgramTypeName] 

    '[Standard_StoredJobData > CheckListSetting] -----------------
    Public ChkList_JOBNO As String = "ChkList_JOBNO"
    Public ChkList_JOBNAME As String = "ChkList_JOBNAME"

    Public ChkList_PA_Year As String = "ChkList_PA_Year"
    Public ChkList_PA_Month As String = "ChkList_PA_Month"
    Public ChkList_PA_Day As String = "ChkList_PA_Day"
    Public ChkList_OS_Year As String = "ChkList_OS_Year"
    Public ChkList_OS_Month As String = "ChkList_OS_Month"
    Public ChkList_OS_Day As String = "ChkList_OS_Day"
    Public ChkList_CFM_Year As String = "ChkList_CFM_Year"
    Public ChkList_CFM_Month As String = "ChkList_CFM_Month"
    Public ChkList_CFM_Day As String = "ChkList_CFM_Day"
    Public ChkList_ELE_Year As String = "ChkList_ELE_Year"
    Public ChkList_ELE_Month As String = "ChkList_ELE_Month"
    Public ChkList_ELE_Day As String = "ChkList_ELE_Day"
    Public ChkList_PA_ChkBox As String = "ChkList_PA_ChkBox"
    Public ChkList_OS_ChkBox As String = "ChkList_OS_ChkBox"
    Public ChkList_CFM_ChkBox As String = "ChkList_CFM_ChkBox"
    Public ChkList_ELE_ChkBox As String = "ChkList_ELE_ChkBox"

    Public ChkList_Q1No_ChkBox As String = "ChkList_Q1No_ChkBox"
    Public ChkList_Q1Yes_ChkBox As String = "ChkList_Q1Yes_ChkBox"
    Public ChkList_Q2No_ChkBox As String = "ChkList_Q2No_ChkBox"
    Public ChkList_Q2Yes_ChkBox As String = "ChkList_Q2Yes_ChkBox"
    Public ChkList_Q3No_ChkBox As String = "ChkList_Q3No_ChkBox"
    Public ChkList_Q3Yes_ChkBox As String = "ChkList_Q3Yes_ChkBox"
    Public ChkList_Q5No_ChkBox As String = "ChkList_Q5No_ChkBox"
    Public ChkList_Q5Std_ChkBox As String = "ChkList_Q5Std_ChkBox"
    Public ChkList_Q5NoStd_ChkBox As String = "ChkList_Q5NoStd_ChkBox"
    Public ChkList_Q6No_ChkBox As String = "ChkList_Q6No_ChkBox"
    Public ChkList_Q6Yes_ChkBox As String = "ChkList_Q6Yes_ChkBox"
    Public ChkList_Q6YesChk_ChkBox As String = "ChkList_Q6YesChk_ChkBox"
    Public ChkList_Q6YesItem_ChkBox As String = "ChkList_Q6YesItem_ChkBox"
    Public ChkList_Q7No_ChkBox As String = "ChkList_Q7No_ChkBox"
    Public ChkList_Q7Yes_ChkBox As String = "ChkList_Q7Yes_ChkBox"
    Public ChkList_Q8No_ChkBox As String = "ChkList_Q8No_ChkBox"
    Public ChkList_Q8Yes_ChkBox As String = "ChkList_Q8Yes_ChkBox"
    Public ChkList_Q8Item_ChkBox As String = "ChkList_Q8Item_ChkBox"
    Public ChkList_Q9No_ChkBox As String = "ChkList_Q9No_ChkBox"
    Public ChkList_Q9Yes_ChkBox As String = "ChkList_Q9Yes_ChkBox"

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
    '----------------- [Standard_StoredJobData > CheckListSetting] 

    '[Standard_StoredJobData > CheckList_PrgmSetting] -----------------
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
    Public ChkList_Prgm_4_TestContent As String = "ChkList_Prgm_TestContent"
    '----------------- [Standard_StoredJobData > CheckList_PrgmSetting] 


    Public ChkList_P1_PageName As String = "CheckList_P1"
    Public ChkList_P2_PageName As String = "CheckList_P2"

    '----------- NOPRINT_INFO Sheet ------------------
    Public JOBNO As String = "JOBNO"
    Public JOBNO_OLD As String = "JOBNO_OLD"
    Public JOBNO_MOD As String = "JOBNO_MOD"
    Public JOBNAME As String = "JOBNAME"
    Public DESIGENED As String = "DESIGENED"
    Public CHECKED As String = "CHECKED"
    Public APPROVED As String = "APPROVED"
    Public DRAW_DATE As String = "DRAW_DATE"
    'Public PAGE_NUM_SPEC As String = "PAGE_NUM_SPEC"
    'Public PAGE_NUM_FM As String = "PAGE_NUM_FM"
    '----------- NOPRINT_INFO Sheet ------------------



    '[Tool_Database > NameManagerSetting] ---------------------
    'Public SPEC_SheetName_1 As String = "SPEC_SheetName_1"
    'Public SPEC_SheetName_5 As String = "SPEC_SheetName_5"
    Public SPEC_CAR_NAME As String = "SPEC_CAR_NAME"
    Public SPEC_CAR_NO As String = "SPEC_CAR_NO"
    Public SPEC_CAR_OPE As String = "SPEC_CAR_OPE"
    Public SPEC_CAR_TOPFL As String = "SPEC_CAR_TOPFL"
    Public SPEC_CAR_BTMFL As String = "SPEC_CAR_BTMFL"
    Public SPEC_CAR_STOP As String = "SPEC_CAR_STOP"
    Public SPEC_CAR_SPEED As String = "SPEC_CAR_SPEED"
    Public SPEC_CAR_FLNAME As String = "SPEC_CAR_FLNAME"
    Public SPEC_CAR_MACHINE_TYPE As String = "SPEC_CAR_MACHINE_TYPE"
    Public SPEC_CAR_CONTROL_WAY As String = "SPEC_CAR_CONTROL_WAY"
    Public SPEC_CAR_PURPOSE As String = "SPEC_CAR_PURPOSE"
    Public SPEC_CAR_LOCATION As String = "SPEC_CAR_LOCATION"

    Public SPEC_AUTO_DR As String = "SPEC_AUTO_DR"
    Public SPEC_AUTO_DR_PHOTOEYE As String = "SPEC_AUTO_DR_PHOTOEYE"
    Public SPEC_AUTO_DR_SAFETY As String = "SPEC_AUTO_DR_SAFETY"
    Public SPEC_CANCELL_CALL As String = "SPEC_CANCELL_CALL"
    Public SPEC_CANCELL_CALL_SCOB As String = "SPEC_CANCELL_CALL_SCOB"
    Public SPEC_CANCELL_CALL_SIX As String = "SPEC_CANCELL_CALL_SIX"
    Public SPEC_CANCELL_BEHIND As String = "SPEC_CANCELL_BEHIND"
    Public SPEC_LAMP_CHK As String = "SPEC_LAMP_CHK"
    'Public SPEC_EC_BOOK As String = "SPEC_EC_BOOK"
    'Public SPEC_INSTALL_BOOK As String = "SPEC_INSTALL_BOOK"
    Public SPEC_AUTO_FAN As String = "SPEC_AUTO_FAN"
    Public SPEC_AUTO_FAN_ION As String = "SPEC_AUTO_FAN_ION"
    'Public SPEC_AUTO_LIGHT As String = "SPEC_AUTO_LIGHT"
    'Public SPEC_RUN_OPEN As String = "SPEC_RUN_OPEN"
    Public SPEC_CC_CANCEL As String = "SPEC_CC_CANCEL"
    Public SPEC_AUTO_PASS As String = "SPEC_AUTO_PASS"
    'Public SPEC_AUTO_LEVEL As String = "SPEC_AUTO_LEVEL"
    'Public SPEC_OPERATION As String = "SPEC_OPERATION"
    Public SPEC_OPERATION_TYPE As String = "SPEC_OPERATION_TYPE"
    Public SPEC_INSTALL_OPE As String = "SPEC_INSTALL_OPE"
    Public SPEC_INDEP_OPE As String = "SPEC_INDEP_OPE"
    Public SPEC_INDEP_OPE_CMD As String = "SPEC_INDEP_OPE_CMD"
    Public SPEC_UCMP As String = "SPEC_UCMP"
    Public SPEC_HIN_CPI As String = "SPEC_HIN_CPI"
    Public SPEC_FIRE_OPE As String = "SPEC_FIRE_OPE"
    Public SPEC_FIRE_OPE_SIGNAL As String = "SPEC_FIRE_OPE_SIGNAL"
    Public SPEC_FIREMAN As String = "SPEC_FIREMAN"
    Public SPEC_PARKING As String = "SPEC_PARKING"
    Public SPEC_PK_CMD1 As String = "SPEC_PK_CMD1"
    Public SPEC_PK_CMD2 As String = "SPEC_PK_CMD2"
    Public SPEC_PK_EN_CMD1 As String = "SPEC_PK_EN_CMD1"
    Public SPEC_SEISMIC As String = "SPEC_SEISMIC"
    Public SPEC_SEISMIC_CANCEL As String = "SPEC_SEISMIC_CANCEL"
    Public SPEC_CPI As String = "SPEC_CPI"
    Public SPEC_CPI_EMER As String = "SPEC_CPI_EMER"
    Public SPEC_CPI_FM As String = "SPEC_CPI_FM"
    Public SPEC_CPI_OLT As String = "SPEC_CPI_OLT"
    Public SPEC_CAR_GONG As String = "SPEC_CAR_GONG"
    Public SPEC_CAR_GONG_POS As String = "SPEC_CAR_GONG_POS"
    Public SPEC_CAR_GONG_CARTOP As String = "SPEC_CAR_GONG_CARTOP"
    Public SPEC_CAR_GONG_CARTOPBTM As String = "SPEC_CAR_GONG_CARTOPBTM"
    Public SPEC_CAR_GONG_COB As String = "SPEC_CAR_GONG_COB"
    Public SPEC_CAR_GONG_VONIC As String = "SPEC_CAR_GONG_VONIC"
    Public SPEC_HALL_GONG As String = "SPEC_HALL_GONG"
    Public SPEC_HPI As String = "SPEC_HPI"
    Public SPEC_HPI_MSG As String = "SPEC_HPI_MSG"
    Public SPEC_HPI_MAIN As String = "SPEC_HPI_MAIN"
    Public SPEC_DR_HOLD As String = "SPEC_DR_HOLD"
    Public SPEC_CRD As String = "SPEC_CRD"
    Public SPEC_CRD_TYPE As String = "SPEC_CRD_TYPE"
    Public SPEC_CRD_SPEC As String = "SPEC_CRD_SPEC"
    Public SPEC_CRD_RVS_CALL As String = "SPEC_CRD_RVS_CALL"
    Public SPEC_CRD_INPUT As String = "SPEC_CRD_INPUT"
    Public SPEC_CRD_ANTI As String = "SPEC_CRD_ANTI"
    Public SPEC_CRD_AUTO_Y As String = "SPEC_CRD_AUTO_Y"
    Public SPEC_CRD_AUTO_N As String = "SPEC_CRD_AUTO_N"
    Public SPEC_CRD_RGL4_Y As String = "SPEC_CRD_RGL4_Y"
    Public SPEC_CRD_RGL4_N As String = "SPEC_CRD_RGL4_N"
    Public SPEC_CRD_RGL5_Y As String = "SPEC_CRD_RGL5_Y"
    Public SPEC_CRD_RGL5_N As String = "SPEC_CRD_RGL5_N"
    Public SPEC_FORCE_CLOSE As String = "SPEC_FORCE_CLOSE"

    Public SPEC_EMER_POWER As String = "SPEC_EMER_POWER"
    Public SPEC_EMER_POWER_GROUP As String = "SPEC_EMER_POWER_GROUP"
    Public SPEC_EMER_POWER_CarName As String = "SPEC_EMER_POWER_CarName"
    Public SPEC_EMER_POWER_EscapeFL As String = "SPEC_EMER_POWER_EscapeFL"
    Public SPEC_EMER_POWER_RETURN As String = "SPEC_EMER_POWER_RETURN"
    Public SPEC_EMER_POWER_CONTINUE As String = "SPEC_EMER_POWER_CONTINUE"
    Public SPEC_EMER_SIGNAL As String = "SPEC_EMER_SIGNAL"
    Public SPEC_EMER_CAPCITY As String = "SPEC_EMER_CAPCITY"
    Public SPEC_EMER_INPUT As String = "SPEC_EMER_INPUT"
    Public SPEC_EMER_ADDRESS As String = "SPEC_EMER_ADDRESS"
    Public SPEC_LANDIC As String = "SPEC_LANDIC"
    Public SPEC_MLF_RETURN As String = "SPEC_MLF_RETURN"
    Public SPEC_VONIC As String = "SPEC_VONIC"
    Public SPEC_VONIC_BZ As String = "SPEC_VONIC_BZ"
    Public SPEC_VONIC_STD_C As String = "SPEC_VONIC_STD_C"
    Public SPEC_VONIC_STD_E As String = "SPEC_VONIC_STD_E"
    Public SPEC_VONIC_NSTD_C As String = "SPEC_VONIC_NSTD_C"
    Public SPEC_VONIC_NSTD_E As String = "SPEC_VONIC_NSTD_E"

    Public SPEC_WCOB As String = "SPEC_WCOB"
    Public SPEC_WCOB_SUB As String = "SPEC_WCOB_SUB"
    Public SPEC_WCOB_RING As String = "SPEC_WCOB_RING"
    Public SPEC_WCOB_BZ As String = "SPEC_WCOB_BZ"
    Public SPEC_ELVIC As String = "SPEC_ELVIC"
    Public SPEC_ELVIC_CMD As String = "SPEC_ELVIC_CMD"
    Public SPEC_HLL As String = "SPEC_HLL"
    Public SPEC_ATT As String = "SPEC_ATT"
    Public SPEC_FLOOD As String = "SPEC_FLOOD"
    Public SPEC_LS1M As String = "SPEC_LS1M"
    Public SPEC_PRU As String = "SPEC_PRU"
    Public SPEC_FRONT_REAR_DR As String = "SPEC_FRONT_REAR_DR"
    Public SPEC_LOAD_CELL As String = "SPEC_LOAD_CELL"
    Public SPEC_LOAD_CELL_CAR_BTM As String = "SPEC_LOAD_CELL_CAR_BTM"
    Public SPEC_LOAD_CELL_CAR_BTM_POS As String = "SPEC_LOAD_CELL_CAR_BTM_POS"
    Public SPEC_LOAD_CELL_MR As String = "SPEC_LOAD_CELL_MR"
    Public SPEC_LOAD_CELL_MR_POS As String = "SPEC_LOAD_CELL_MR_POS"
    'Public SPEC_FORCE_CLOSE As String = "SPEC_FORCE_CLOSE"
    Public SPEC_OPE_SW As String = "SPEC_OPE_SW"
    Public SPEC_OPE_SW_POS As String = "SPEC_OPE_SW_POS"
    Public SPEC_OPE_SW_ADDRESS As String = "SPEC_OPE_SW_ADDRESS"
    Public SPEC_WTB As String = "SPEC_WTB"
    Public SPEC_WTB_ERROR As String = "SPEC_WTB_ERROR"
    Public SPEC_WTB_STOP As String = "SPEC_WTB_STOP"
    Public SPEC_WTB_FM As String = "SPEC_WTB_FM"
    Public SPEC_WTB_EQ As String = "SPEC_WTB_EQ"
    Public SPEC_WTB_INDEP As String = "SPEC_WTB_INDEP"
    Public SPEC_WTB_NORMAL As String = "SPEC_WTB_NORMAL"
    Public SPEC_WTB_URGENT As String = "SPEC_WTB_URGENT"
    Public SPEC_WTB_FO As String = "SPEC_WTB_FO"
    Public SPEC_WTB_EMERPOWER As String = "SPEC_WTB_EMERPOWER"
    Public SPEC_WTB_ALART As String = "SPEC_WTB_ALART"
    Public SPEC_WTB_EQMAC As String = "SPEC_WTB_EQMAC"
    '--------------------- [Tool_Database > NameManagerSetting] 

    ' Setting Table Sheet -----------------------------
    ' Only --------------
    Public SetTable_PHOTOEYE_ONLY As String = "SetTable_PHOTOEYE_ONLY"
    Public SetTable_MechSafety_ONLY As String = "SetTable_MechSafety_ONLY"
    Public SetTable_SCOB_ONLY As String = "SetTable_SCOB_ONLY"
    Public SetTable_ION_ONLY As String = "SetTable_ION_ONLY"
    Public SetTable_AutoPass_ONLY As String = "SetTable_AutoPass_ONLY"
    Public SetTable_Indep_ONLY As String = "SetTable_Indep_ONLY"
    Public SetTable_HinCpi_ONLY As String = "SetTable_HinCpi_ONLY"
    Public SetTable_FIRE_ONLY As String = "SetTable_FIRE_ONLY"
    Public SetTable_Fireman_ONLY As String = "SetTable_Fireman_ONLY"
    Public SetTable_PARKING_ONLY As String = "SetTable_PARKING_ONLY"
    Public SetTable_Seismic_ONLY As String = "SetTable_Seismic_ONLY"
    Public SetTable_SeismicSW_ONLY As String = "SetTable_SeismicSW_ONLY"
    Public SetTable_Seismic_SENSOR_ONLY As String = "SetTable_Seismic_SENSOR_ONLY"
    Public SetTable_CPI_FM_ONLY As String = "SetTable_CPI_FM_ONLY"
    Public SetTable_CPI_OLT_ONLY As String = "SetTable_CPI_OLT_ONLY"
    Public SetTable_HallGong_ONLY As String = "SetTable_HallGong_ONLY"
    Public SetTable_HpiFM_ONLY As String = "SetTable_HpiFM_ONLY"
    Public SetTable_VonicBz_ONLY As String = "SetTable_VonicBz_ONLY"
    Public SetTable_DrHold_ONLY As String = "SetTable_DrHold_ONLY"
    Public SetTable_Landic_ONLY As String = "SetTable_Landic_ONLY"
    Public SetTable_MFLReturn_ONLY As String = "SetTable_MFLReturn_ONLY"
    Public SetTable_MFLReturn_FL_ONLY As String = "SetTable_MFLReturn_FL_ONLY"
    Public SetTable_VONIC_ONLY As String = "SetTable_VONIC_ONLY"
    Public SetTable_ELVIC_ONLY As String = "SetTable_ELVIC_ONLY"
    Public SetTable_ELVIC_ParkingFL_ONLY As String = "SetTable_ELVIC_ParkingFL_ONLY"
    Public SetTable_WCOB_ONLY As String = "SetTable_WCOB_ONLY"
    Public SetTable_WSCOB_ONLY As String = "SetTable_WSCOB_ONLY"
    Public SetTable_HLL_ONLY As String = "SetTable_HLL_ONLY"
    Public SetTable_ATT_ONLY As String = "SetTable_ATT_ONLY"
    Public SetTable_LS1M_ONLY As String = "SetTable_LS1M_ONLY"
    Public SetTable_PRU_ONLY As String = "SetTable_PRU_ONLY"
    Public SetTable_LoadCellPos_MR_ONLY As String = "SetTable_LoadCellPos_MR_ONLY"
    Public SetTable_LoadCellPos_CarBtm_ONLY As String = "SetTable_LoadCellPos_CarBtm_ONLY"
    Public SetTable_ForceClose_ONLY As String = "SetTable_ForceClose_ONLY"
    Public SetTable_FrontRearDr_ONLY As String = "SetTable_FrontRearDr_ONLY"
    Public SetTable_OpeSw_ONLY As String = "SetTable_OpeSw_ONLY"
    '-------------- Only 

    'Public SetTable_MACHINE_TYPE As String = "SetTable_MACHINE_TYPE"
    Public SetTable_PARKING_FL As String = "SetTable_PARKING_FL"
    Public SetTable_ESCAPE_FL As String = "SetTable_ESCAPE_FL"
    Public SetTable_ESCAPE_FL_ONLY As String = "SetTable_ESCAPE_FL_ONLY"
    Public SetTable_MAIN_FL As String = "SetTable_MAIN_FL"
    Public SetTable_FLOOD_FL As String = "SetTable_FLOOD_FL"
    Public SetTable_RESULT_WITH As String = "SetTable_RESULT_WITH"
    Public SetTable_RESULT_WITHOUT As String = "SetTable_RESULT_WITHOUT"
    Public SetTable_NO As String = "SetTable_NO"
    Public SetTable_NC As String = "SetTable_NC"
    Public SetTable_PK_ELVIC As String = "SetTable_PK_ELVIC"
    Public SetTable_PK_COB As String = "SetTable_PK_COB"
    Public SetTable_PK_WTB As String = "SetTable_PK_WTB"
    Public SetTable_PK_SW As String = "SetTable_PK_SW"
    Public SetTable_PK_DROPEN As String = "SetTable_PK_DROPEN"
    Public SetTable_PK_DRCLOSE As String = "SetTable_PK_DRCLOSE"
    Public SetTable_PK_EN_DROPEN As String = "SetTable_PK_EN_DROPEN"
    Public SetTable_PK_EN_DRCLOSE As String = "SetTable_PK_EN_DRCLOSE"

    Public SetTable_SeismicSW_WITH As String = "SetTable_SeismicSW_WITH"
    Public SetTable_SeismicSW_WITHOUT As String = "SetTable_SeismicSW_WITHOUT"
    Public SetTable_Seismic_SENSOR As String = "SetTable_Seismic_SENSOR"

    Public SetTable_CPI_SEISMIC As String = "SetTable_CPI_SEISMIC"
    Public SetTable_CPI_FIRE As String = "SetTable_CPI_FIRE"
    Public SetTable_CPI_EMER As String = "SetTable_CPI_EMER"
    Public SetTable_CPI_FIREMAN As String = "SetTable_CPI_FIREMAN"
    Public SetTable_CPI_OLT As String = "SetTable_CPI_OLT"
    Public SetTable_CAR_TOP As String = "SetTable_CAR_TOP"
    Public SetTable_CAR_TOP_BTM As String = "SetTable_CAR_TOP_BTM"
    Public SetTable_CAR_COB As String = "SetTable_CAR_COB"
    Public SetTable_CAR_VONIC As String = "SetTable_CAR_VONIC"
    Public SetTable_HALL_OLT As String = "SetTable_HALL_OLT"
    Public SetTable_HALL_MAIN As String = "SetTable_HALL_MAIN"
    Public SetTable_HALL_INDEP As String = "SetTable_HALL_INDEP"
    Public SetTable_HALL_FM As String = "SetTable_HALL_FM"
    Public SetTable_CRD_TYPE_NOTALL As String = "SetTable_CRD_TYPE_NOTALL"
    Public SetTable_CRD_TYPE_ALL As String = "SetTable_CRD_TYPE_ALL"
    Public SetTable_CRD_SPEC_Y As String = "SetTable_CRD_SPEC_Y"
    Public SetTable_CRD_SPEC_N As String = "SetTable_CRD_SPEC_N"
    Public SetTable_CRD_RVS_CALL_Y As String = "SetTable_CRD_RVS_CALL_Y"
    Public SetTable_CRD_RVS_CALL_N As String = "SetTable_CRD_RVS_CALL_N"
    Public SetTable_CRD_ANTI_Y As String = "SetTable_CRD_ANTI_Y"
    Public SetTable_CRD_ANTI_N As String = "SetTable_CRD_ANTI_N"
    Public SetTable_CRD_TIME_SET As String = "SetTable_CRD_TIME_SET"
    Public SetTable_CRD_ID_4 As String = "SetTable_CRD_ID_4"
    Public SetTable_CRD_ID_5 As String = "SetTable_CRD_ID_5"

    Public SetTable_OpeSW_Content As String = "SetTable_OpeSW_Content"

    Public SetTable_WCOB_BZ_Y As String = "SetTable_WCOB_BZ_Y"
    Public SetTable_WCOB_BZ_N As String = "SetTable_WCOB_BZ_N"
    Public SetTable_WCOB_RING_Y As String = "SetTable_WCOB_RING_Y"
    Public SetTable_WCOB_RING_N As String = "SetTable_WCOB_RING_N"

    Public SetTable_CP43x_WITH As String = "SetTable_CP43x_WITH"
    Public SetTable_CP43x_WITHOUT As String = "SetTable_CP43x_WITHOUT"
    Public SetTable_FLOOR_N1 As String = "SetTable_FLOOR_N1"
    Public SetTable_FLOOR_N2 As String = "SetTable_FLOOR_N2"
    Public SetTable_FLOOR_N3 As String = "SetTable_FLOOR_N3"
    '----------------------------- Setting Table Sheet 

    '送狀 Sheet [暫停使用]------------------------------------
    'Public DWG_SheetName As String = "DWG_SheetName"
    Public DWG_StdPage As String = "DWG_StdPage"
    Public DWG_StdPage_withoutVonic As String = "DWG_StdPage_withoutVonic"
    Public DWG_PRK As String = "DWG_PRK"
    Public DWG_Start_GrNo As String = "DWG_Start_GrNo"
    Public DWG_JOBNO As String = "DWG_JOBNO"
    Public DWG_Start_Construction As String = "DWG_Start_Construction"

    Public DWG_HsinChu As String = "DWG_HsinChu"
    Public DWG_Tainan As String = "DWG_Tainan"
    Public DWG_Taipei As String = "DWG_Taipei"
    Public DWG_Taichung As String = "DWG_Taichung"
    Public DWG_Kaohsiung As String = "DWG_Kaohsiung"
    Public DWG_Taoyuan As String = "DWG_Taoyuan"

    Public DWG_Start_HsinChu As String = "DWG_Start_HsinChu"
    Public DWG_Start_Tainan As String = "DWG_Start_Tainan"
    Public DWG_Start_Taipei As String = "DWG_Start_Taipei"
    Public DWG_Start_Taichung As String = "DWG_Start_Taichung"
    Public DWG_Start_Kaohsiung As String = "DWG_Start_Kaohsiung"
    Public DWG_Start_Taoyuan As String = "DWG_Start_Taoyuan"
    Public DWG_Start_Produce As String = "DWG_Start_Produce"

    Public DWG_Chinese_HsinChu As String = "新竹"
    Public DWG_Chinese_Tainan As String = "台南"
    Public DWG_Chinese_Taipei As String = "台北"
    Public DWG_Chinese_Taichung As String = "台中"
    Public DWG_Chinese_Kaohsiung As String = "高雄"
    Public DWG_Chinese_Taoyuan As String = "桃園"
    '------------------------------------ 送狀 Sheet [暫停使用]
    'Public IMPORTANT_SheetName As String = "IMPORTANT_SheetName"
    Public IMPORTANT_FAN As String = "IMPORTANT_FAN"
    Public IMPORTANT_FAN_CONTENT As String = "IMPORTANT_FAN_CONTENT"
    Public IMPORTANT_BALANCE As String = "IMPORTANT_BALANCE"
    Public IMPORTANT_WCOB As String = "IMPORTANT_WCOB"
    Public IMPORTANT_DOOR As String = "IMPORTANT_DOOR"
    Public IMPORTANT_HIN As String = "IMPORTANT_HIN"
    Public IMPORTANT_HIN_FL As String = "IMPORTANT_HIN_FL"
    Public IMPORTANT_HIN_PCB As String = "IMPORTANT_HIN_PCB"

    Public MMIC_MACHINE_TYPE As String = "MMIC_MACHINE_TYPE"
    Public MMIC_OPERATION As String = "MMIC_OPERATION"
    Public MMIC_FLEX_N_SV As String = "MMIC_FLEX_N_SV"
    Public MMIC_CP43x As String = "MMIC_CP43x"
    Public MMIC_EBase As String = "MMIC_EBase"

    Public MMIC_CarNo As String = "MMIC_CarNo"
    Public MMIC_CarObj As String = "MMIC_CarObj"
    Public MMIC_ECarNo As String = "MMIC_ECarNo"
    Public MMIC_ECarObj As String = "MMIC_ECarObj"
    Public SV_CarNo As String = "SV_CarNo"
    Public SV_CarObj As String = "SV_CarObj"
    Public SV_EBase As String = "SV_EBase"
    Public SV_ECarNo As String = "SV_ECarNo"
    Public SV_ECarObj As String = "SV_ECarObj"
    Public VONIC_ROM_Device As String = "VONIC_ROM_Device"
    Public VONIC_Quantity As String = "VONIC_Quantity"
    Public VONIC_CarNo As String = "VONIC_CarNo"
    Public VONIC_CarObj As String = "VONIC_CarObj"
    '--------------- [Tool_Database > NameManagerSetting]

    '[Tool_Database > VD10_ProgramType] -----------------------------
    '[MMIC > VD10 > Type Combobox]
    Public VD10TypeName_Array As String = "VD10TypeName_Array"
    '[MMIC > VD10 > Base TextBox]
    Public VD10_TW_STD_LOWER As String = "VD10_TW_STD_LOWER"
    Public VD10_TW_STD_HIGHER As String = "VD10_TW_STD_HIGHER"
    Public VD10_SP_STD_STOREY As String = "VD10_SP_STD_STOREY"
    Public VD10_SP_STD_FLOOR As String = "VD10_SP_STD_FLOOR"
    Public VD10_TW_NSTD_Lobby_R As String = "VD10_TW_NSTD_Lobby_R"
    Public VD10_TW_NSTD_1M_2M As String = "VD10_TW_NSTD_1M_2M"
    Public VD10_TW_NSTD_Taiwanese As String = "VD10_TW_NSTD_Taiwanese"
    Public VD10_TW_NSTD_Taiwanese_B As String = "VD10_TW_NSTD_Taiwanese_B"
    Public VD10_HK_NSTD_B_G As String = "VD10_HK_NSTD_B_G"
    Public VD10_SP_NSTD_M As String = "VD10_SP_NSTD_M"
    Public VD10_Other_Path As String = "VD10_Other_Path"
    '----------------------------- [Tool_Database > VD10_ProgramType] 



    'SQL-------------
    Public SQLite_connectionPath_Tool As String = "M:\DESIGN\BACK UP\yc_tian\Tool Application\SQLite\" 'SQLite的檔案位置
    Public SQLite_ToolDBMS_Name As String = "Tool_Database.sqlite"
    Public SQLite_StdJobDataDBMS_Name As String = "Standard_StoredJobData.sqlite"

    'Public SQLite_tableName_AllEmployee As String = "AllEmployeeNumber"
    Public SQLite_tableName_Basic As String = "BasicSetting"
    'Public SQLite_tableName_NameManager As String = "NameManagerSetting"
    Public SQLite_tableName_NameManager_FinalCheck As String = "NameManagerSetting_FinalCheck"
    Public SQLite_tableName_NameManager_TW As String = "NameManagerSetting_TW"
    Public SQLite_tableName_NameManager_CheckList As String = "NameManagerSetting_CheckList"
    Public SQLite_tableName_GSP_ProgramType As String = "GSP_ProgramType"
    Public SQLite_tableName_GSP_ProgramTypeName As String = "GSP_ProgramTypeName"
    Public SQLite_tableName_MMIC_ProgramType As String = "MMIC_ProgramType"
    Public SQLite_tableName_MMIC_ProgramTypeName As String = "MMIC_ProgramTypeName"
    Public SQLite_tableName_VD10_ProgramType As String = "VD10_ProgramType"
    Public SQLite_tableName_VD10_ProgramTypeName As String = "VD10_ProgramTypeName"
    '-------------SQL



    Dim sqlite_connect As SQLiteConnection
    Dim sqlite_cmd As SQLiteCommand

    Dim sqlite_dataReader As SQLiteDataReader

    Dim read_txt As String


    ''' <summary>
    ''' 從SQL tableName(資料表)中選出selectName(項目)的全部內容填進wCmbBox(ComboBox)中
    ''' </summary>
    ''' <param name="selectName"> selectName(項目) </param>
    ''' <param name="tableName"> tableName(資料表) </param>
    ''' <param name="wCmbBox"> wCmbBox(ComboBox) </param>
    ''' <returns></returns>
    Overloads Function read_DbmsData(selectName As String, tableName As String,
                                     wCmbBox As ComboBox,
                                     sqlite_path As String, sqlite_name As String)
        '----------------------- SQLite Reading -----------------------------
        Try
            sqlite_connect = New SQLiteConnection("Data Source=" & sqlite_path & sqlite_name)

            sqlite_connect.Open()
            sqlite_cmd = sqlite_connect.CreateCommand()

            sqlite_cmd.CommandText = "SELECT * FROM " & tableName '可依照自行需求變動
            sqlite_dataReader = sqlite_cmd.ExecuteReader()

            If wCmbBox.Items.Count = 0 Then
                While sqlite_dataReader.Read
                    read_txt = sqlite_dataReader(selectName).ToString()
                    If read_txt <> "" Then
                        wCmbBox.Items.Add(read_txt)
                    End If
                End While
            End If
            sqlite_connect.Close()
        Catch ex As Exception
            '寫入errorInfo.log中(尚未設計)
        End Try
        '----------------------- SQLite Reading -----------------------------
    End Function
    ''' <summary>
    ''' 從SQL tableName(資料表)中選出selectName(項目)的全部內容填進wTxtBox(TextBox)中
    ''' </summary>
    ''' <param name="selectName"> selectName(項目) </param>
    ''' <param name="tableName"> tableName(資料表) </param>
    ''' <param name="wTxtBox"> wTxtBox(TextBox) </param>
    ''' <returns></returns>
    Overloads Function read_DbmsData(selectName As String, tableName As String,
                                     wTxtBox As TextBox,
                                     sqlite_path As String, sqlite_name As String)
        '----------------------- SQLite Reading -----------------------------
        sqlite_connect = New SQLiteConnection("Data Source=" & sqlite_path & sqlite_name)

        sqlite_connect.Open()
        sqlite_cmd = sqlite_connect.CreateCommand()

        sqlite_cmd.CommandText = "SELECT * FROM " & tableName '可依照自行需求變動
        sqlite_dataReader = sqlite_cmd.ExecuteReader()

        If wTxtBox.Text = Nothing Then
            While sqlite_dataReader.Read
                read_txt = sqlite_dataReader(selectName).ToString()
                If read_txt <> "" Then
                    wTxtBox.Text = read_txt
                End If
            End While
        End If
        sqlite_connect.Close()
        '----------------------- SQLite Reading -----------------------------
    End Function
    ''' <summary>
    ''' 從SQL tableName(資料表)中選出selectName(項目)的單獨內容
    ''' </summary>
    ''' <param name="selectName"> selectName(項目) </param>
    ''' <param name="tableName"> tableName(資料表) </param>
    ''' <returns></returns>
    Overloads Function read_DbmsData(selectName As String, tableName As String,
                                     sqlite_path As String, sqlite_name As String)
        '----------------------- SQLite Reading -----------------------------
        Dim output_ToSpec As Output_ToSpec = New Output_ToSpec()

        sqlite_connect = New SQLiteConnection("Data Source=" & sqlite_path & sqlite_name)

        sqlite_connect.Open()
        sqlite_cmd = sqlite_connect.CreateCommand()

        sqlite_cmd.CommandText = "SELECT * FROM " & tableName '可依照自行需求變動
        sqlite_dataReader = sqlite_cmd.ExecuteReader()


        While sqlite_dataReader.Read
            read_txt = sqlite_dataReader(selectName).ToString()
            If read_txt <> "" Then
                Return read_txt
                output_ToSpec.returnError_specName = read_txt
            End If
        End While

        sqlite_connect.Close()
        '----------------------- SQLite Reading -----------------------------
    End Function




    Sub read_DbmsData_catalogPage(selectName As String, tableName As String, wChkListBox As CheckedListBox,
                                  sqlite_path As String, sqlite_name As String)
        '----------------------- SQLite Reading -----------------------------
        sqlite_connect = New SQLiteConnection("Data Source=" & sqlite_path & sqlite_name)

        sqlite_connect.Open()
        sqlite_cmd = sqlite_connect.CreateCommand()

        sqlite_cmd.CommandText = "SELECT * FROM " & tableName
        sqlite_dataReader = sqlite_cmd.ExecuteReader()

        If wChkListBox.Items.Count = 0 Then
            While sqlite_dataReader.Read
                '
                'read_txt = Left(JobMaker_Form.Basic_JobNoNew_TextBox.Text, 7) & sqlite_dataReader(selectName).ToString()
                read_txt = sqlite_dataReader(selectName).ToString()
                If read_txt <> "" Then
                    wChkListBox.Items.Add(read_txt)
                End If
            End While
        End If
        sqlite_connect.Close()
        '----------------------- SQLite Reading -----------------------------
    End Sub

    Function read_DbmsData_Employee(selectName As String, tableName As String, inputNum As String, sqlite_path As String, sqlite_name As String) '進入SQL比對工號是否相同並且回傳第N行

        '----------------------- SQLite Reading -----------------------------
        sqlite_connect = New SQLiteConnection("Data Source=" & sqlite_path & sqlite_name)

        sqlite_connect.Open()
        sqlite_cmd = sqlite_connect.CreateCommand()

        sqlite_cmd.CommandText = "SELECT * FROM " & tableName
        sqlite_dataReader = sqlite_cmd.ExecuteReader()


        While sqlite_dataReader.Read
            read_txt = sqlite_dataReader(selectName).ToString()
            If read_txt <> "" And read_txt = inputNum Then
                EmployeeRow = EmployeeRow + 1
                Return read_txt
                Exit While
            ElseIf read_txt = "" Then
                EmployeeRow = 0
                Exit While
            Else
                EmployeeRow = EmployeeRow + 1
            End If
        End While
        sqlite_connect.Close()
        '----------------------- SQLite Reading -----------------------------
    End Function

    Overloads Function read_DbmsData_Employee_getRow(selectName As String, tableName As String, inputNum As String, sqlite_path As String, sqlite_name As String) '進入SQL中判斷工號是第N行
        '----------------------- SQLite Reading -----------------------------
        sqlite_connect = New SQLiteConnection("Data Source=" & sqlite_path & sqlite_name)

        sqlite_connect.Open()
        sqlite_cmd = sqlite_connect.CreateCommand()

        sqlite_cmd.CommandText = "SELECT * FROM " & tableName
        sqlite_dataReader = sqlite_cmd.ExecuteReader()

        Dim SQL_Row As Integer

        While sqlite_dataReader.Read
            read_txt = sqlite_dataReader(selectName).ToString()
            If read_txt <> "" Then
                SQL_Row = SQL_Row + 1
                If read_txt = inputNum Then
                    Return SQL_Row
                End If
            End If
        End While
        sqlite_connect.Close()
        '----------------------- SQLite Reading -----------------------------
    End Function

    Overloads Function read_DbmsData_Employee_getRow(selectName As String, tableName As String, sql_row As Integer, sqlite_path As String, sqlite_name As String) '取得工號後進SQL中取得第N行的值
        '----------------------- SQLite Reading -----------------------------
        sqlite_connect = New SQLiteConnection("Data Source=" & sqlite_path & sqlite_name)

        sqlite_connect.Open()
        sqlite_cmd = sqlite_connect.CreateCommand()

        sqlite_cmd.CommandText = "SELECT " & selectName & " FROM " & tableName & " WHERE ROWID=" & sql_row
        sqlite_dataReader = sqlite_cmd.ExecuteReader()


        While sqlite_dataReader.Read
            read_txt = sqlite_dataReader(selectName).ToString()
            If read_txt <> "" Then
                Return read_txt
            End If
        End While
        sqlite_connect.Close()
        '----------------------- SQLite Reading -----------------------------
    End Function
End Class
