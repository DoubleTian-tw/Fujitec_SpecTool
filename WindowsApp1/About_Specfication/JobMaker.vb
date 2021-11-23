Imports System.Text
Imports Microsoft.Office.Interop
'Imports System.IO.Directory
'Imports System.Runtime.InteropServices
Imports System.IO
'Imports System.Text.RegularExpressions

Public Class JobMaker_Form
    '其他form
    Dim chalink As ChangeLink = New ChangeLink()
    Dim get_nameManager As Spec_NameManager = New Spec_NameManager()
    Dim output_ToSpec As Output_ToSpec = New Output_ToSpec()

    '<基本>
    ''' <summary>
    ''' use_basic_chkbox按下次數
    ''' </summary>
    Dim use_basic_chkbox_clickTimes As Integer
    ''' <summary>
    ''' use_chkList_chkbox按下次數
    ''' </summary>
    Dim use_chkList_chkbox_clickTimes As Integer
    ''' <summary>
    ''' use_program_chkbox按下次數
    ''' </summary>
    Dim use_program_chkbox_clickTimes As Integer
    ''' <summary>
    ''' use_dwg_chkbox按下次數
    ''' </summary>
    Dim use_DWG_chkbox_clickTimes As Integer
    ''' <summary>
    ''' use_spec_chkbox按下次數
    ''' </summary>
    Dim use_spec_chkbox_clickTimes As Integer
    ''' <summary>
    ''' use_important_chkbox按下次數
    ''' </summary>
    Dim use_important_chkbox_clickTimes As Integer
    ''' <summary>
    ''' use_mmic_chkbox按下次數
    ''' </summary>
    Dim use_mmic_chkbox_clickTimes As Integer
    ''' <summary>
    ''' use_EepData_chkbox按下次數
    ''' </summary>
    Dim use_EepData_chkbox_clickTimes As Integer
    ''' <summary>
    ''' FinalCheck_Button按下次數
    ''' </summary>
    'Dim finalCheck_Btm_clickTimes As Integer

    ''' <summary>
    ''' 目前使用者的工號
    ''' </summary>
    Dim currentEmployee_Number As String

    ''' <summary>
    ''' 目前使用者的中文姓名
    ''' </summary>
    Dim currentEmployee_ChineseName As String

    ''' <summary>
    ''' 提示資訊欄
    ''' </summary>
    Dim Load_info_txt As String = "請拖曳檔案至文字框中 或 複製檔案路徑包含檔案的附檔名"

    ''' <summary>
    ''' 原始【表單】的長度
    ''' </summary>
    Const iniForm_width As Integer = 715
    ''' <summary>
    ''' 原始【表單】的高度
    ''' </summary>
    Const iniForm_height As Integer = 670
    ''' <summary>
    ''' 原始【關閉視窗】的Position X 
    ''' </summary>
    Const iniCloseBtn_X As Integer = 660
    ''' <summary>
    ''' 原始【關閉視窗】的Position Y 
    ''' </summary>
    Const iniCloseBtn_Y As Integer = 4
    ''' <summary>
    ''' 改變【表單】後的長度
    ''' </summary>
    Const reForm_width As Integer = 1150
    ''' <summary>
    ''' 改變【關閉視窗】後的Position X
    ''' </summary>
    Const reCloseBtn_X As Integer = 1095

    ''' <summary>
    ''' 地區選擇
    ''' </summary>
    Dim localSelect As String
    Const Taiwan As String = "台灣"
    Const HongKong As String = "香港"
    Const Singapore As String = "新加坡"

    'EXCEL use
    Dim msExcel_app As Excel.Application
    Dim msExcel_workbook As Excel.Workbook
    Dim msExcel_worksheet As Excel.Worksheet

    '--- 仕樣書 ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' 工番的仕樣書路徑
    ''' </summary>
    Private jobSpecPath As String
    Private jobDefaultPath As String
    '仕樣數量
    ''' <summary>
    ''' 使用者輸入的電梯總數
    ''' </summary>
    Public LiftNum As Integer

    Dim ContainNum As Integer

    '仕樣>重要設定>HIN
    ''' <summary>
    ''' [重要設定>HIN] 儲存自動產生樓層名稱的陣列
    ''' </summary>
    Public arr_liftName() As String     'HIN中自動產生樓層名稱
    ''' <summary>
    ''' [重要設定>HIN] 儲存自動產生樓層數量的陣列
    ''' </summary>
    Public arr_liftStopFL() As Integer  'HIN中自動產生樓層數量
    ''' <summary>
    ''' [重要設定>HIN] 儲存自動產生樓層最高的陣列
    ''' </summary>
    'Public arr_liftTopFL() As String  '
    ''' <summary>
    ''' 第一列儲存標準值，其他行列儲存當樓層選擇值的陣列，進行判斷後輸出 e.g #1,2:WITH/#3:WITHOUT
    ''' </summary>
    Public arr_liftStopFl_EachContent(,) As String
    ''' <summary>
    ''' [重要設定>HIN] 暫時存入WITH/WITHOUT等標準值的陣列 
    ''' </summary>
    Public arr_liftStopFl_StdContent() As String     '暫存HIN中自動產生樓層的內容
    ''' <summary>
    ''' [重要設定>HIN] 暫時存入使用者輸入的全樓層值的陣列
    ''' </summary>
    Public arr_liftStopFL_userContent(,) As String '暫存使用者在HIN中自動產生樓層選擇的內容
    '------------------------------------------------------------------------------------------------------- 仕樣書 
    ''' <summary>
    ''' 送狀自動生成控制項打勾的數量
    ''' </summary>
    Dim clp_count As Integer


    ''' <summary>
    ''' 原始或變更後表單大小
    ''' </summary>
    Enum JMForm_size
        ''' <summary>
        ''' 原始大小
        ''' </summary>
        ini_size = 0
        ''' <summary>
        ''' 改變後大小
        ''' </summary>
        re_size
    End Enum

    ''' <summary>
    ''' [方法 > 變更表單大小]
    ''' </summary>
    ''' <param name="mysize">原始或變更</param>
    Private Sub Resize_JMForm(mysize As JMForm_size)
        Select Case mysize
            Case mysize.ini_size
                With Me
                    .Width = iniForm_width
                    .Height = iniForm_height
                End With
                With JobMaker_Close_Button
                    .Location = New Point(iniCloseBtn_X, iniCloseBtn_Y)
                End With
                ResultOutput_TextBox.Visible = False
                ResultClose_Button.Visible = False
                With Result_Loading_PictureBox
                    .Enabled = False
                    .Visible = False
                End With

            Case mysize.re_size
                Me.Width = reForm_width
                JobMaker_Close_Button.Location = New Point(reCloseBtn_X, iniCloseBtn_Y)
                ResultOutput_TextBox.Visible = True
                ResultClose_Button.Visible = True
                With Result_Loading_PictureBox
                    .Enabled = True
                    .Visible = True
                End With

        End Select
    End Sub

    ''' <summary>
    ''' 當JobMaker打開時執行此程序
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobMaker_Form_Load(sender As Object, e As EventArgs) Handles Me.Load

        '判斷工號 -------------------------------------------------------------------------------------
        Dim em_i As Integer
        Dim em_bool As Boolean
        currentEmployee_Number =
            InputBox($"請輸入您的工號 >> {vbCrLf} 例如 : TWN2100 或 twn2100 或 2100", "員工工號輸入", "TWN")


        em_bool = False
        For em_i = 1 To get_nameManager.AllEmployee_Type.Count
            If currentEmployee_Number <> "" And
               currentEmployee_Number = get_nameManager.read_DbmsData_Employee(
                                        get_nameManager.AllEmployee_Type(em_i - 1),
                                        get_nameManager.SQLite_tableName_Basic,
                                        currentEmployee_Number,
                                        get_nameManager.SQLite_connectionPath_Tool,
                                        get_nameManager.SQLite_ToolDBMS_Name) Then
                em_bool = True
                Exit For
            End If
        Next


        If em_bool = False Then
            MsgBox("輸入錯誤",, "提醒")
            Me.Close()
        ElseIf em_bool = True Then
            MsgBox(currentEmployee_Number & "歡迎來到Fuji峽谷", , "Hello bro")
            currentEmployee_ChineseName =
                get_nameManager.read_DbmsData_Employee_getRow(get_nameManager.EmployeeChinese,
                                                              get_nameManager.SQLite_tableName_Basic,
                                                              get_nameManager.EmployeeRow,
                                                              get_nameManager.SQLite_connectionPath_Tool,
                                                              get_nameManager.SQLite_ToolDBMS_Name)
            If currentEmployee_Number = "2100" Or
               UCase(currentEmployee_Number) = "TWN2100" Then
                testBtn_GroupBox.Visible = True
                Load_AutoLoad_GroupBox.Visible = True
            End If
            '----------------------------------------------------------------------------------- 判斷工號

            ' SQLite 遺失 --------------------------------
            Dim fileExitPath As String
            fileExitPath = get_nameManager.SQLite_connectionPath_Tool & get_nameManager.SQLite_ToolDBMS_Name
            If Not File.Exists(fileExitPath) Then
                MsgBox($"未取得Sqlite檔案請確認路徑: {fileExitPath} 是否正確?")
                errorInfo.createError_InfoTxt("Sqlite路徑異常")
                errorInfo.writeInfoError_InfoTxt($"{fileExitPath} 是否正確?")
            End If
            '-------------------------------- SQLite 遺失 

            '時間start
            JobMaker_Timer.Enabled = True
            '讀取link
            chalink.Initialization_ini()

            chalink.Topmost_setting(Me, False)
            chalink.formPositionOnScreen_Setting(Me, chalink.sKeyValueScr.ToString, chalink.sKeyValuePos.ToString)


            '初始化load
            Resize_JMForm(JMForm_size.ini_size)

            JobMaker_initialization()
        End If
    End Sub

    ''' <summary>
    ''' 初始化JobMaker內資料
    ''' </summary>
    Private Sub JobMaker_initialization()

        Me.Text = $"JobMaker_價補妹可▼ω▼ User:[{currentEmployee_ChineseName}大濕]"

        '初始化地區選擇 
        localSelect = TW_ToolStripMenuItem.Text


        '初始化 Load > 仕樣書 分頁 ------------------------
        With JMFileCho_Spec_TextBox
            .Text = Load_info_txt
            .ForeColor = Color.Gray
        End With

        With JM_DefaultPath_Spec_Label
            .Text = chalink.ChgLink_DefaultPath_Spec_TextBox.Text
        End With

        jobDefaultPath = Load_Job_OutputPath_TextBox.Text
        '------------------------初始化 Load > 仕樣書 分頁

        '初始化 Load > ChkList 分頁----------------------
        With JMFileCho_ChkList_TextBox
            .Text = Load_info_txt
            .ForeColor = Color.Gray
        End With

        With JM_DefaultPath_CheckList_Label
            .Text = chalink.ChgLink_DefaultPath_CheckList_TextBox.Text
        End With
        '----------------------初始化 Load > ChkList 分頁

        '初始化 Load > 載入SQLite 分頁---------------------
        With Load_SQLite_Path_TextBox
            .Text = Load_info_txt
            .ForeColor = Color.Gray
        End With
        With JM_DefaultPath_SQLite_Label
            .Text = get_nameManager.SQLite_connectionPath_Tool
        End With
        '---------------------初始化 Load > 載入SQLite 分頁

        '初始化 Load > 工番路徑 分頁---------------------
        With Load_Job_JobSearch_TextBox '輸入工番
            For Each file In Directory.GetDirectories(Load_Job_OutputPath_TextBox.Text)
                .AutoCompleteCustomSource.Add(Path.GetFileName(file))
            Next
            .Text = "TW-"
        End With

        With Load_Job_BasePath_ComboBox '仕樣書Base路徑
            For Each file In Directory.GetDirectories(Load_Job_BasePath_ComboBox.Text)
                .Items.Add(file)
            Next
            .Text += "\FP-17 (TW)"
        End With
        '---------------------初始化 Load > 工番路徑 分頁


        '---------------------------------- 初始化 Load 分頁 結束


        '初始化 基本 分頁 開始 -----------------------------------
        If TW_ToolStripMenuItem.Checked Then
            Basic_JobNoNew_TextBox.Text = "TW-"
        End If
        '----------------------------------- 初始化 基本 分頁 結束

        '初始化 Check List 分頁 開始 -----------------------------------
        With CheckList_FlowLayoutPanel
            '設定CheckList Panel的排列順序
            .Controls.SetChildIndex(ChkList_1_Panel, 0)
            .Controls.SetChildIndex(ChkList_2_Panel, 1)
            .Controls.SetChildIndex(ChkList_3_Panel, 2)
            '.Enabled = False
        End With
        With CheckList2_FlowLayoutPanel
            .Controls.SetChildIndex(ChkList_4_Panel, 0)
            .Controls.SetChildIndex(ChkList_5_Panel, 1)
            .Controls.SetChildIndex(ChkList_6_Panel, 2)
        End With
        With CheckList3_FlowLayoutPanel
            .Controls.SetChildIndex(ChkList_7_Panel, 0)
            .Controls.SetChildIndex(ChkList_8_Panel, 1)
            .Controls.SetChildIndex(ChkList_9_Panel, 2)
        End With
        '----------------------------------- 初始化 Check List 分頁 結束 

        '初始化 程式變更表 分頁 開始 -----------------------------------
        With ProgramChange_FlowLayoutPanel
            '設定ProgramChange Panel的排列順序
            .Controls.SetChildIndex(use_ProgramChg_Panel1, 0)
            .Controls.SetChildIndex(use_ProgramChg_Panel2, 1)
            .Controls.SetChildIndex(use_ProgramChg_Panel3, 2)
        End With
        '----------------------------------- 初始化 程式變更表 分頁 結束

        '初始化 送狀 分頁 開始 -----------------------------------
        DWG_PrkName_ComboBox.Items.Clear()
        '----------------------------------- 初始化 送狀 分頁 結束

        '初始化 仕樣 分頁 開始 -----------------------------------
        Spec_EscapeFL_TextBox_height = Spec_EscapeFL_TextBox.Height
        Spec_Fire_Panel_height = Spec_Fire_Panel.Height
        Spec_Parking_FL_TextBox_height = Spec_Parking_FL_TextBox.Height
        Spec_Parking_Panel_height = Spec_Parking_Panel.Height
        Spec_MFLReturn_FL_TextBox_height = Spec_MFLReturn_FL_TextBox.Height
        Spec_MFLReturn_Panel_height = Spec_MFLReturn_Panel.Height
        Spec_Flood_FL_TextBox_height = Spec_Flood_FL_TextBox.Height
        Spec_Flood_Panel_height = Spec_Flood_Panel.Height

        With Spec_ParkingFL_DR_ComboBox
            .Items.Add(get_nameManager.TB_DR_CLOSE)
            .Items.Add(get_nameManager.TB_DR_OPEN)
        End With
        With Spec_EmerSignal_ComboBox
            .Items.Add(get_nameManager.TB_NO)
            .Items.Add(get_nameManager.TB_NC)
        End With

        Spec_CarGong_Top_TextBox.Text = get_nameManager.TB_CarTop          '車廂上到著鈴-車廂上
        Spec_CarGong_TopBtm_TextBox.Text = get_nameManager.TB_CarTopBtm    '車廂上到著鈴-車廂上下
        Spec_CarGong_COB_TextBox.Text = get_nameManager.TB_WithCOB         '車廂上到著鈴-COB
        Spec_CarGong_VONIC_TextBox.Text = get_nameManager.TB_InVONIC       '車廂上到著鈴-Vonic

        Spec_DRAuto_ComboBox.Text = Spec_DRAuto_ComboBox.Items(0)                       '開門
        Spec_CancellCall_ComboBox.Text = Spec_CancellCall_ComboBox.Items(0)             '取消嬉戲
        Spec_CancellBehind_ComboBox.Text = Spec_CancellBehind_ComboBox.Items(0)         '逆呼
        Spec_LampChk_ComboBox.Text = Spec_LampChk_ComboBox.Items(0)                     '檢點
        Spec_AutoFan_ComboBox.Text = Spec_AutoFan_ComboBox.Items(0)                     '風扇連動
        Spec_CCCancell_ComboBox.Text = Spec_CCCancell_ComboBox.Items(0)                 '取消叫車
        'Spec_Operation_ComboBox.Text = Spec_Operation_ComboBox.Items(0)                 '操作方式
        Spec_UCMP_ComboBox.Text = Spec_UCMP_ComboBox.Items(0)                           '戶開行走
        Spec_HinCpi_ComboBox.Text = Spec_HinCpi_ComboBox.Items(0)                       'HIN/CPI
        Spec_MFLReturn_ComboBox.Text = Spec_MFLReturn_ComboBox.Items(0)                 '基準階
        Spec_VonicBz_ComboBox.Text = Spec_VonicBz_ComboBox.Items(0)                     'Vonic BZ
        Spec_DrHold_ComboBox.Text = Spec_DrHold_ComboBox.Items(0)                        '開門延長按鈕
        Spec_LoadCell_ComboBox.Text = Spec_LoadCell_ComboBox.Items(0)                   'Load Cell
        Spec_install_ope_ComboBox.Text = Spec_install_ope_ComboBox.Items(0)             '拒付運轉
        Spec_FireSignal_ComboBox.Text = Spec_FireSignal_ComboBox.Items(0)               '火災運轉訊號
        Spec_ParkingFL_DR_ComboBox.Text = Spec_ParkingFL_DR_ComboBox.Items(1)           'Parking休止開關門
        Spec_WTB_ComboBox.Text = Spec_WTB_ComboBox.Items(0)                             'wWTB
        '----------------------------------- 初始化 仕樣 分頁 結束 
    End Sub


    'LOAD ------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Help_Button_Click(sender As Object, e As EventArgs) Handles Help_Button.Click
        MagicTool.open_DirectPath($"{Application.StartupPath}\{ProgramAllPath.folderName_ppt}\{ProgramAllName.fileName_Manualpptx}")
    End Sub
    ''' <summary>
    ''' [Load > 工具欄]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub 地區選擇ToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles 地區選擇ToolStripMenuItem.DropDownOpening
        TW_ToolStripMenuItem.Enabled =
            If(TW_ToolStripMenuItem.Checked, False, True)
        HK_ToolStripMenuItem.Enabled =
            If(HK_ToolStripMenuItem.Checked, False, True)
        SP_ToolStripMenuItem.Enabled =
            If(SP_ToolStripMenuItem.Checked, False, True)
    End Sub
    ''' <summary>
    ''' [Load > 工具欄 > 台灣]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TW_ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TW_ToolStripMenuItem.Click
        TW_ToolStripMenuItem.Checked = True
        HK_ToolStripMenuItem.Checked = False
        SP_ToolStripMenuItem.Checked = False

        Load_Job_JobSearch_TextBox.Text = "TW-"  '輸入工番
        Basic_JobNoNew_TextBox.Text = "TW-" '基本 > JobNo(新)
        ChkList_5_std_RadioButton.Checked = True
        Load_Job_BasePath_ComboBox.Text = "\\10.213.2.103\job\21 SPEC&EPROM DATA\FP-17 (TW)"
    End Sub
    ''' <summary>
    ''' [Load > 工具欄 > 香港]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub HK_ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HK_ToolStripMenuItem.Click
        TW_ToolStripMenuItem.Checked = False
        HK_ToolStripMenuItem.Checked = True
        SP_ToolStripMenuItem.Checked = False

        Load_Job_JobSearch_TextBox.Text = "MZH"  '輸入工番
        Basic_JobNoNew_TextBox.Text = "MZH" '基本 > JobNo(新)
        Basic_Local_ComboBox.Text = "Hong Kong"
        'ChkList_5_std_RadioButton.Checked = True
        Load_Job_BasePath_ComboBox.Text = "\\10.213.2.103\job\21 SPEC&EPROM DATA\FP-17 (HK_MOD)"
    End Sub
    ''' <summary>
    ''' [Load > 工具欄 > 新加坡]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SP_ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SP_ToolStripMenuItem.Click
        TW_ToolStripMenuItem.Checked = False
        HK_ToolStripMenuItem.Checked = False
        SP_ToolStripMenuItem.Checked = True

        Load_Job_JobSearch_TextBox.Text = "WMB"  '輸入工番
        Basic_JobNoNew_TextBox.Text = "WMB" '基本 > JobNo(新)
        Basic_Local_ComboBox.Text = "Singapore"
        'ChkList_5_std_RadioButton.Checked = True
        Load_Job_BasePath_ComboBox.Text = "\\10.213.2.103\job\21 SPEC&EPROM DATA\FP-17 (SP)"
    End Sub

    ''' <summary>
    ''' [Load > 仕樣書 > Check]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobMaker_LOAD_Spec_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles JobMaker_LOAD_Spec_CheckBox.CheckedChanged
        '仕樣書是否啟用?
        If JobMaker_LOAD_Spec_CheckBox.Checked Then
            With Use_Basic_CheckBox
                .Text = "基本資料必填"
                .ForeColor = Color.Red
            End With
            With Use_SpecBasic_CheckBox
                .Text = "基本仕樣必填"
                .ForeColor = Color.Red
            End With
            With Use_Imp_CheckBox
                .Text = "重要設定必填"
                .ForeColor = Color.Red
            End With
            With Use_mmic_CheckBox
                .Text = "MMIC必填"
                .ForeColor = Color.Red
            End With
            Load_Spec_GroupBox.Enabled = True
        Else
            With Use_Basic_CheckBox
                .Text = ""
            End With
            With Use_SpecBasic_CheckBox
                .Text = ""
            End With
            With Use_Imp_CheckBox
                .Text = ""
            End With
            With Use_mmic_CheckBox
                .Text = ""
            End With

            All_OutputButton.Enabled = False
            Spec_OutputButton.Enabled = False

            Load_Spec_GroupBox.Enabled = False
            'finalCheck_Btm_clickTimes = 0
        End If
    End Sub

    ''' <summary>
    ''' [Load > 仕樣書路徑 > 輸入工番的TextBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobPathSelect_TextBox_TextChanged(sender As Object, e As EventArgs) Handles Load_Job_JobSearch_TextBox.TextChanged
        jobSpecPath = ""
        Try
            With Load_Job_JobSearch_TextBox
                If UCase(.Text) <> "" Then
                    If .TextLength = 5 Then
                        Select Case Strings.Left(UCase(.Text), 3)
                            Case "TW-"
                                jobSpecPath = $"{jobDefaultPath}\TW-\{Strings.Left(UCase(.Text), 5)}00番台\"
                            Case "TMB"
                                jobSpecPath = $"{jobDefaultPath}\TMB\{Strings.Left(UCase(.Text), 5)}00\"
                            Case "MZH"
                                jobSpecPath = $"{jobDefaultPath}\MZH\{Strings.Left(UCase(.Text), 5)}00番台\"
                            Case "WMB"
                                jobSpecPath = $"{jobDefaultPath}\WMB\{Strings.Left(UCase(.Text), 5)}00番台\"
                        End Select
                        For Each file In Directory.GetDirectories(jobSpecPath)
                            .AutoCompleteCustomSource.Add(Path.GetFileName(file))
                        Next
                    ElseIf .TextLength > 9 Then
                        Dim folderName As String = ""
                        If Load_Job_JobSelect_RadioButton.Checked Then
                            folderName = "SPEC"
                        ElseIf Load_Job_ChkListSelect_RadioButton.Checked Then
                            folderName = "CHECK LIST"
                        End If
                        Select Case Strings.Left(UCase(.Text), 3)
                            Case "TW-"
                                jobSpecPath =
                                    $"{jobDefaultPath}\TW-\{Strings.Left(UCase(.Text), 5)}00番台\{Load_Job_JobSearch_TextBox.Text}\{folderName}"
                            Case "TMB"
                                jobSpecPath =
                                    $"{jobDefaultPath}\TMB\{Strings.Left(UCase(.Text), 5)}00\{Load_Job_JobSearch_TextBox.Text}\{folderName}"
                            Case "MZH"
                                jobSpecPath =
                                    $"{jobDefaultPath}\MZH\{Strings.Left(UCase(.Text), 5)}00番台\{Load_Job_JobSearch_TextBox.Text}\{folderName}"
                            Case "WMB"
                                jobSpecPath =
                                    $"{jobDefaultPath}\WMB\{Strings.Left(UCase(.Text), 5)}00番台\{Load_Job_JobSearch_TextBox.Text}\{folderName}"
                        End Select
                    Else
                        jobSpecPath = "M:\DESIGN\軟體設計\01 JOB"
                    End If
                End If
            End With
        Catch ex As Exception

        End Try
    End Sub
    ''' <summary>
    ''' [Load > 仕樣書路徑 > 按我Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobPathEnter_Button_Click(sender As Object, e As EventArgs) Handles JobPathEnter_Button.Click
        Load_Job_OutputPath_TextBox.Text = jobSpecPath
    End Sub
    ''' <summary>
    ''' [Load > 仕樣書路徑 > 最後輸出路徑 Open Diolog button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobPathSelect_Button_Click(sender As Object, e As EventArgs) Handles JobPathSelect_Button.Click
        Dim mpath As String

        If chalink.ChgLink_DefaultPath_Spec_TextBox.Text = "" Then
            '在ChangLink Form中沒有預設路徑就給"C:\"或其他
            mpath = "C:\"
        Else
            '在ChangLink Form中有預設路徑就給預設
            mpath = Load_Job_OutputPath_TextBox.Text
        End If

        '打開diologResult
        ChangeLink.OpenFilePath_event(Load_Job_OutputPath_TextBox, Load_Job_OutputPath_TextBox.Text)
    End Sub

    ''' <summary>
    ''' [Load > 仕樣書路徑 > 來源Excel Open Diolog button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobBasePathSelect_Button_Click(sender As Object, e As EventArgs) Handles JobBasePathSelect_Button.Click
        ChangeLink.OpenFile_event(Load_Job_BasePath_ComboBox, chalink.OpenFileType.mExcel, Load_Job_BasePath_ComboBox.Text)
    End Sub

    ''' <summary>
    ''' [Load > 仕樣書路徑 > 仕樣書RadioButton]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobPathSelect_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles Load_Job_JobSelect_RadioButton.CheckedChanged
        If Load_Job_JobSelect_RadioButton.Checked Then
            JobPathSelect_GroupBox.Enabled = True
            With Use_Basic_CheckBox
                .Text = "基本資料必填"
                .ForeColor = Color.Red
            End With
            With Use_SpecBasic_CheckBox
                .Text = "基本仕樣必填"
                .ForeColor = Color.Red
            End With
            With Use_Imp_CheckBox
                .Text = "重要設定必填"
                .ForeColor = Color.Red
            End With
            With Use_mmic_CheckBox
                .Text = "MMIC必填"
                .ForeColor = Color.Red
            End With
            JobBasePathSelect_GroupBox.Enabled = True
            '更新輸出的路徑
            If Load_Job_JobSearch_TextBox.TextLength > 9 Then
                Dim splitPath() As String
                splitPath = Split(jobSpecPath, "\")

                jobSpecPath = jobSpecPath.Replace(splitPath(splitPath.Length - 1), "SPEC")
                Load_Job_OutputPath_TextBox.Text = jobSpecPath
            End If
            '===============================更新輸出的路徑 

            'Else
            '    JobPathSelect_GroupBox.Enabled = False
            '    JobBasePathSelect_GroupBox.Enabled = False
        Else
            With Use_Basic_CheckBox
                .Text = ""
            End With
            With Use_SpecBasic_CheckBox
                .Text = ""
            End With
            With Use_Imp_CheckBox
                .Text = ""
            End With
            With Use_mmic_CheckBox
                .Text = ""
            End With

            All_OutputButton.Enabled = False
            Spec_OutputButton.Enabled = False

            'finalCheck_Btm_clickTimes = 0
        End If
    End Sub
    ''' <summary>
    ''' [Load > 仕樣書路徑 > CheckList RadioButton]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkListPathSelect_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles Load_Job_ChkListSelect_RadioButton.CheckedChanged
        If Load_Job_ChkListSelect_RadioButton.Checked Then
            JobPathSelect_GroupBox.Enabled = True
            JobBasePathSelect_GroupBox.Enabled = True

            With Use_Basic_CheckBox
                .Text = "基本資料必填"
                .ForeColor = Color.Red
            End With
            With Use_ChkList_CheckBox
                .Text = "Check List資料必填"
                .ForeColor = Color.Red
            End With
            With Use_Program_CheckBox
                .Text = "程式變更資料有改程式必填"
                .ForeColor = Color.Red
            End With

            Load_ChkList_GroupBox.Enabled = True
            '更新輸出的路徑 ===============================
            If Load_Job_JobSearch_TextBox.TextLength > 9 Then
                Dim splitPath() As String
                splitPath = Split(jobSpecPath, "\")

                jobSpecPath = jobSpecPath.Replace(splitPath(splitPath.Length - 1), "CHECK LIST")
                Load_Job_OutputPath_TextBox.Text = jobSpecPath
            End If
            '===============================更新輸出的路徑 
        Else

            CheckList_OutputButton.Enabled = False
            With Use_Basic_CheckBox
                .Text = ""
            End With
            With Use_ChkList_CheckBox
                .Text = ""
            End With
            With Use_Program_CheckBox
                .Text = ""
            End With
        End If
    End Sub
    ''' <summary>
    ''' [Load > CheckList > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobMaker_LOAD_ChkList_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles JobMaker_LOAD_ChkList_CheckBox.CheckedChanged
        'Check List是否啟用?
        If JobMaker_LOAD_ChkList_CheckBox.Checked Then

            With Use_Basic_CheckBox
                .Text = "基本資料必填"
                .ForeColor = Color.Red
            End With
            With Use_ChkList_CheckBox
                .Text = "Check List資料必填"
                .ForeColor = Color.Red
            End With
            With Use_Program_CheckBox
                .Text = "程式變更資料有改程式必填"
                .ForeColor = Color.Red
            End With

            Load_ChkList_GroupBox.Enabled = True
        Else
            With Use_Basic_CheckBox
                .Text = ""
            End With
            With Use_ChkList_CheckBox
                .Text = ""
            End With
            With Use_Program_CheckBox
                .Text = ""
            End With

            CheckList_OutputButton.Enabled = False

            Load_ChkList_GroupBox.Enabled = False
        End If
    End Sub

    Private Sub JobMaker_LOAD_AutoLoad_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles JobMaker_LOAD_AutoLoad_CheckBox.CheckedChanged
        If JobMaker_LOAD_AutoLoad_CheckBox.Checked Then
            Load_AutoLoad_GroupBox.Enabled = True
        Else
            Load_AutoLoad_GroupBox.Enabled = False
        End If
    End Sub
    'LOAD分頁 -> 自動讀取分頁 ------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' [DragEnter功能][Load > 自動讀取 > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_AutoLoad_TextBox_DragEnter(sender As Object, e As DragEventArgs) Handles JMFileCho_AutoLoad_TextBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    ''' <summary>
    ''' [DragDrop功能][Load > 自動讀取 > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_AutoLoad_TextBox_DragDrop(sender As Object, e As DragEventArgs) Handles JMFileCho_AutoLoad_TextBox.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each mpath In file
            If System.IO.File.Exists(mpath) Then
                JMFileCho_AutoLoad_TextBox.Text = mpath
                JMFileCho_AutoLoad_TextBox.ForeColor = Color.Black
            End If
        Next
    End Sub
    ''' <summary>
    ''' [Load > 自動讀取 > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_AutoLoad_TextBox_TextChanged(sender As Object, e As EventArgs) Handles JMFileCho_AutoLoad_TextBox.TextChanged
        If JMFileCho_AutoLoad_TextBox.Text <> Load_info_txt Then
            If JMFileCho_AutoLoad_TextBox.Text <> "" Then
                JMFileConfirm_AutoLoad_Button.Enabled = True
                Check_direction_file_is_needed_type({"xls", "xlsx", "xlsm"}, JMFileCho_AutoLoad_TextBox)
            Else
                JMFileConfirm_AutoLoad_Button.Enabled = False
            End If
        End If
    End Sub
    Private Sub JMFileCho_AutoLoad_Button_Click(sender As Object, e As EventArgs) Handles JMFileCho_AutoLoad_Button.Click
        'Dim mpath As String

        'If chalink.ChgLink_DefaultPath_Spec_TextBox.Text = "" Then
        '    '在ChangLink Form中沒有預設路徑就給"C:\"或其他
        '     mpath = "M:\DESIGN\BACK UP\"
        'Else
        '    '在ChangLink Form中有預設路徑就給預設
        '    'mpath = chalink.ChgLink_DefaultPath_Spec_TextBox.Text
        'End If

        '打開diologResult
        ChangeLink.OpenFile_event(JMFileCho_AutoLoad_TextBox,
                                  ChangeLink.OpenFileType.mExcel,
                                  "M:\DESIGN\BACK UP\")
    End Sub
    Private Sub JMFileConfirm_AutoLoad_Button_Click(sender As Object, e As EventArgs) Handles JMFileConfirm_AutoLoad_Button.Click
        Output_new_excel_and_open_from_textbox(JMFileCho_AutoLoad_TextBox.Text)
        msExcel_app.Visible = True
        'Dim autoLoad As AutoLoad_intoJobMaker = New AutoLoad_intoJobMaker

        Try
            AutoLoad_inJobMaker.readData_fromExcel(msExcel_workbook)
        Catch ex As Exception
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub
    '------------------------------------------------------------------------------------------------------------LOAD分頁 -> 自動讀取分頁 

    'LOAD分頁 -> 仕樣書分頁 ------------------------------------------------------------------------------------------------------------

    Private Sub JM_Spec_JobSelect_TextBox_TextChanged(sender As Object, e As EventArgs) Handles JM_JobSelect_Spec_TextBox.TextChanged
        Dim default_path As String
        If JM_DefaultPath_Spec_Label.Text = "" Then
            default_path = "C:\"
        Else
            default_path = JM_DefaultPath_Spec_Label.Text & "\"
        End If
        JobSelect_type_into_textBox({"*.xls", "*.xlsx", "*.xlsm"},
                                    default_path,
                                    JM_JobSelect_Spec_ComboBox, JM_JobSelect_Spec_TextBox)
    End Sub
    Private Sub JM_Spec_JobSelect_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles JM_JobSelect_Spec_ComboBox.TextChanged
        JobSelect_add_into_comboBox_and_textBox(JM_DefaultPath_Spec_Label.Text & "\",
                                                JM_JobSelect_Spec_ComboBox,
                                                JMFileCho_Spec_TextBox)
    End Sub
    ''' <summary>
    ''' [DragEnter功能][Load > 仕樣書路徑 > 按我]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobDirectPath_TextBox_DragEnter(sender As Object, e As DragEventArgs) Handles Load_Job_OutputPath_TextBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    ''' <summary>
    ''' [DragDrop功能][Load > 仕樣書路徑 > 按我]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobDirectPath_TextBox_DragDrop(sender As Object, e As DragEventArgs) Handles Load_Job_OutputPath_TextBox.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each mpath In file
            If System.IO.File.Exists(mpath) Then
                Load_Job_OutputPath_TextBox.Text = Path.GetDirectoryName(mpath)
                Load_Job_OutputPath_TextBox.ForeColor = Color.Black
            End If
        Next
    End Sub
    ''' <summary>
    ''' [DragEnter功能][Load > 仕樣書 > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_Spec_TextBox_DragEnter(sender As Object, e As DragEventArgs) Handles JMFileCho_Spec_TextBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    ''' <summary>
    ''' [DragDrop功能][Load > 仕樣書 > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_Spec_TextBox_DragDrop(sender As Object, e As DragEventArgs) Handles JMFileCho_Spec_TextBox.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In file
            If System.IO.File.Exists(path) Then
                JMFileCho_Spec_TextBox.Text = path
                JMFileCho_Spec_TextBox.ForeColor = Color.Black
            End If
        Next
    End Sub
    ''' <summary>
    ''' [Load > 仕樣書 > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_Spec_TextBox_TextChanged(sender As Object, e As EventArgs) Handles JMFileCho_Spec_TextBox.TextChanged
        If JMFileCho_Spec_TextBox.Text <> Load_info_txt Then
            If JMFileCho_Spec_TextBox.Text <> "" Then
                Check_direction_file_is_needed_type({"xls", "xlsx", "xlsm"}, JMFileCho_Spec_TextBox)
            End If
        End If
    End Sub
    ''' <summary>
    ''' 檢查目標路徑的檔案是否為指定檔案格式
    ''' </summary>
    ''' <param name="filter_name">請以Array方式存取</param>
    Private Sub Check_direction_file_is_needed_type(filter_name() As String, select_textbox As TextBox)
        Dim filter_bool As Boolean
        Try
            For f As Integer = 1 To (filter_name).Count
                For i As Integer = 1 To Len(filter_name(f - 1))
                    If Strings.Right(select_textbox.Text, i) = Strings.Right(filter_name(f - 1), i) Then
                        filter_bool = True
                        If i = Len(filter_name(f - 1)) Then
                            Exit For
                        End If
                    Else
                        filter_bool = False
                        Exit For
                    End If
                Next

                If filter_bool Then
                    Exit For
                End If
            Next

            If filter_bool = False Then
                Dim output_msg As String
                output_msg = $"載入檔案非"
                For k As Integer = 1 To (filter_name).Count
                    output_msg += $"{filter_name(k - 1)},"
                Next
                If select_textbox.Text <> "" Then
                    MsgBox(output_msg)
                    select_textbox.Text = ""
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' [Load > 仕樣書 > 路徑 > Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_Spec_Button_Click(sender As Object, e As EventArgs) Handles JMFileCho_Spec_Button.Click
        Dim mpath As String

        If chalink.ChgLink_DefaultPath_Spec_TextBox.Text = "" Then
            '在ChangLink Form中沒有預設路徑就給"C:\"或其他
            mpath = "C:\"
        Else
            '在ChangLink Form中有預設路徑就給預設
            mpath = chalink.ChgLink_DefaultPath_Spec_TextBox.Text
        End If

        '打開diologResult
        ChangeLink.OpenFile_event(JMFileCho_Spec_TextBox,
                                  ChangeLink.OpenFileType.mExcel,
                                  mpath)
    End Sub
    '------------------------------------------------------------------------------------------------------------ LOAD分頁 -> 仕樣書分頁

    'LOAD分頁 -> CheckList分頁 ------------------------------------------------------------------------------------------------------------


    Private Sub JM_CheckList_JobSelect_TextBox_TextChanged(sender As Object, e As EventArgs) Handles JM_JobSelect_CheckList_TextBox.TextChanged
        Dim default_path As String
        If JM_DefaultPath_CheckList_Label.Text = "" Then
            default_path = "C:\"
        Else
            default_path = JM_DefaultPath_CheckList_Label.Text & "\"
        End If
        JobSelect_type_into_textBox({"*.xls", "*.xlsx", "*.xlsm"},
                                    default_path,
                                    JM_JobSelect_CheckList_ComboBox, JM_JobSelect_CheckList_TextBox)
    End Sub
    Private Sub JM_CheckList_JobSelect_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles JM_JobSelect_CheckList_ComboBox.TextChanged
        JobSelect_add_into_comboBox_and_textBox(JM_DefaultPath_CheckList_Label.Text & "\",
                                                JM_JobSelect_CheckList_ComboBox,
                                                JMFileCho_ChkList_TextBox)
    End Sub
    ''' <summary>
    ''' [DragEnter功能][Load > CheckList > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_ChkList_TextBox_DragEnter(sender As Object, e As DragEventArgs) Handles JMFileCho_ChkList_TextBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
    ''' <summary>
    ''' [DragDrop功能][Load > CheckList > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_ChkList_TextBox_DragDrop(sender As Object, e As DragEventArgs) Handles JMFileCho_ChkList_TextBox.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In file
            If System.IO.File.Exists(path) Then
                JMFileCho_ChkList_TextBox.Text = path
            End If
        Next
    End Sub
    ''' <summary>
    ''' [Load > CheckList > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_ChkList_TextBox_TextChanged(sender As Object, e As EventArgs) Handles JMFileCho_ChkList_TextBox.TextChanged
        If JMFileCho_ChkList_TextBox.Text <> Load_info_txt Then
            If JMFileCho_ChkList_TextBox.Text <> "" Then
                'CheckList_OutputButton.Enabled = True
                Check_direction_file_is_needed_type({"xls", "xlsx", "xlsm"}, JMFileCho_ChkList_TextBox)
                'Else
                '    CheckList_OutputButton.Enabled = False
            End If
        End If
    End Sub
    ''' <summary>
    ''' [Load > CheckList > 路徑 > Button ]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_ChkList_Button_Click(sender As Object, e As EventArgs) Handles JMFileCho_ChkList_Button.Click
        Dim mpath As String

        If chalink.ChgLink_DefaultPath_CheckList_TextBox.Text = "" Then
            mpath = "C:\"
        Else
            mpath = chalink.ChgLink_DefaultPath_Spec_TextBox.Text
        End If
        ChangeLink.OpenFile_event(JMFileCho_ChkList_TextBox,
                                  ChangeLink.OpenFileType.mExcel,
                                  mpath)
    End Sub
    '------------------------------------------------------------------------------------------------------------ LOAD分頁 -> CheckList分頁 




    'LOAD分頁 -> 載入SQLite分頁 -------------------------------------------------------------------------------------------------------
    Private Sub JM_SQlite_JobSelect_TextBox_TextChanged(sender As Object, e As EventArgs) Handles Load_SQLite_JobSearch_TextBox.TextChanged
        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        JobSelect_type_into_textBox({"*.sqlite"},
                                    spec_stored.SQLite_connectionPath_Job,
                                    Load_SQLite_JobSearch_ComboBox, Load_SQLite_JobSearch_TextBox)
    End Sub

    Private Sub JobSelect_type_into_textBox(select_type() As String, default_path As String, select_cb As ComboBox, select_tb As TextBox)
        Dim file_Cho As String '目前選擇的檔案名稱 
        'Dim filter_name() As String '要讀取資料夾內的副檔名種類
        'Dim filePath As String '目前路徑


        'filter_name = {"*.sqlite"}
        'filePath = spec_stored.SQLite_connectionPath_Job

        select_cb.Text = ""
        select_cb.Items.Clear()
        'JMFileCho_SQLite_TextBox.Text = ""
        Try
            For Each myFilter In select_type
                For Each file In Directory.GetFileSystemEntries(default_path, myFilter)

                    file_Cho = Strings.Right(file, Len(file) - (Len(default_path)))

                    '將英文轉換大小寫後與目前檔案名稱相比，相同的加入COMBOBOX
                    For i As Integer = 1 To Len(file_Cho)
                        If select_tb.Text.ToUpperInvariant = Strings.Left(file_Cho, i) Or
                           select_tb.Text.ToLowerInvariant = Strings.Left(file_Cho, i) Then
                            select_cb.Items.Add(file_Cho)
                            'JMFileCho_SQLite_TextBox.Text = filePath & JM_SQlite_JobSelect_ComboBox.Text
                        End If
                    Next
                Next
            Next
            If select_cb.Items.Count <> 0 Then
                select_cb.Text = select_cb.Items(0)
            End If
        Catch ex As Exception
            MsgBox("指定常用工番路徑已刪除變動，系統找不到相對應資料夾", vbCritical, "ERROR常用工番路徑ERROR")
        End Try
    End Sub

    Private Sub JM_SQlite_JobSelect_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Load_SQLite_JobSearch_ComboBox.TextChanged
        'Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        JobSelect_add_into_comboBox_and_textBox(JM_DefaultPath_SQLite_Label.Text,
                                                Load_SQLite_JobSearch_ComboBox,
                                                Load_SQLite_Path_TextBox)
    End Sub
    Private Sub JobSelect_add_into_comboBox_and_textBox(default_path As String, select_cb As ComboBox, choosePath_tb As TextBox)
        If select_cb.Text <> "" Then
            choosePath_tb.Text =
                $"{default_path}{select_cb.Text}"
        End If
    End Sub
    ''' <summary>
    ''' [DragDrop功能][Load > 載入SQLite > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_SQLite_Button_DragDrop(sender As Object, e As DragEventArgs) Handles Load_SQLite_Path_TextBox.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In file
            If System.IO.File.Exists(path) Then
                Load_SQLite_Path_TextBox.Text = path
                Load_SQLite_Path_TextBox.ForeColor = Color.Black
            End If
        Next
    End Sub
    ''' <summary>
    ''' [DragEnter][Load > 載入SQLite > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_SQLite_Button_DragEnter(sender As Object, e As DragEventArgs) Handles Load_SQLite_Path_TextBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    ''' <summary>
    ''' [Load > 載入SQLite > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_SQLite_TextBox_TextChanged(sender As Object, e As EventArgs) Handles Load_SQLite_Path_TextBox.TextChanged
        If Load_SQLite_Path_TextBox.Text <> Load_info_txt Then
            If Load_SQLite_Path_TextBox.Text <> "" Then
                JMFileConfirm_SQLite_Button.Enabled = True
                JMFileConfirm_SQLite_FixBug_Button.Enabled = True
                Check_direction_file_is_needed_type({"sqlite"}, Load_SQLite_Path_TextBox)
            End If
        Else
            JMFileConfirm_SQLite_Button.Enabled = False
            JMFileConfirm_SQLite_FixBug_Button.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' [Load > 載入SQLite > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobMaker_LOAD_SQLite_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Load_SQLite_Loading_CheckBox.CheckedChanged
        If Load_SQLite_Loading_CheckBox.Checked Then
            Load_SQLite_GroupBox.Enabled = True
        Else
            Load_SQLite_GroupBox.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' [Load > 載入SQLite > Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_SQLite_Button_Click(sender As Object, e As EventArgs) Handles JMFileCho_SQLite_Button.Click
        Dim mpath As String

        If chalink.ChgLink_DefaultPath_SQLite_TextBox.Text = "" Then
            'spec_stored 中沒有預設路徑就給"C:\"或其他
            mpath = "C:\"
        Else
            'spec_stored 中有預設路徑就給預設
            mpath = chalink.ChgLink_DefaultPath_SQLite_TextBox.Text
        End If

        ChangeLink.OpenFile_event(Load_SQLite_Path_TextBox,
                                  ChangeLink.OpenFileType.mOther,
                                  mpath)

        If Load_SQLite_Path_TextBox.Text <> "" Then
            JMFileConfirm_SQLite_Button.Enabled = True
        End If
    End Sub

    'Public Sub loadStored_controllerCount()
    '    For Each ctrl As Control In JobPath_TabPage.Controls
    '        If GetType(ctrl) Is "GroupBox" Then
    '            For Each ctrl_grp As Control In ctrl.Controls

    '            Next
    '        End If
    '    Next
    'End Sub
    ''' <summary>
    ''' [Load > 載入SQLite > 確認Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileConfirm_SQLite_Button_Click(sender As Object, e As EventArgs) Handles JMFileConfirm_SQLite_Button.Click
        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        Resize_JMForm(JMForm_size.re_size)

        With spec_stored
            .SQLiteLoading_Stored(Path.GetFileName(Load_SQLite_Path_TextBox.Text))
            .outputText_toTextBox_focusOnBelow(ResultOutput_TextBox, "")
        End With

    End Sub

    Dim SQLite_FixBug_Button_ClickCount As Integer = 0
    Private Sub SpecBasic_BugFix_Button_Click(sender As Object, e As EventArgs) Handles JMFileConfirm_SQLite_FixBug_Button.Click
        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        Resize_JMForm(JMForm_size.re_size)

        With spec_stored
            .SQLiteLoading_FixBug_Stored(Path.GetFileName(Load_SQLite_Path_TextBox.Text))
            .outputText_toTextBox_focusOnBelow(ResultOutput_TextBox, "")
        End With

        SQLite_FixBug_Button_ClickCount += 1

        MsgBox("讀取結束")
    End Sub
    '--------------------------------------------------------------------------------------------------------LOAD分頁 -> 載入SQLite分頁 

    ''' <summary>
    ''' [LOAD > 輸出 > 仕樣書]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_OutputButton_Click(sender As Object, e As EventArgs) Handles Spec_OutputButton.Click
        '開啟excel
        Try
            Output_new_excel_and_open_from_textbox(Load_Job_BasePath_ComboBox.Text)
            'msExcel_app.Visible = True

            Resize_JMForm(JMForm_size.re_size) '重新變大小

            output_ToSpec.Spec_FinalCheck(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_Spec_Std(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_Basic(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_TW(LiftNum, ContainNum, msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_Important(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_MMIC(msExcel_workbook, msExcel_app)

            Output_open_excel_folder_and_saveAs_when_done($"{Load_Job_OutputPath_TextBox.Text}\{Basic_JobNoNew_TextBox.Text}-SPEC",
                                                          Load_Job_OutputPath_TextBox.Text)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.Spec_OutputButton_Click")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub

    ''' <summary>
    ''' [LOAD > 輸出 > Check List]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CheckList_OutputButton_Click(sender As Object, e As EventArgs) Handles CheckList_OutputButton.Click
        Try
            Output_new_excel_and_open_from_textbox(JMFileCho_ChkList_TextBox.Text)
            msExcel_app.Visible = True

            Resize_JMForm(JMForm_size.re_size) '重新變大小
            output_ToSpec.Spec_CheckList(msExcel_workbook, msExcel_app)

            Output_open_excel_folder_and_saveAs_when_done($"{Load_Job_OutputPath_TextBox.Text}\{Basic_JobNoNew_TextBox.Text}-SPEC",
                                                          Load_Job_OutputPath_TextBox.Text)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.CheckList_OutputButton_Click")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub
    '-------------------------------------------------------------------------------------------------------------------- Check List.
    ''' <summary>
    ''' [Load > 輸出 > DWG送狀]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DWG_OutputButton_Click_1(sender As Object, e As EventArgs) Handles DWG_OutputButton.Click
        Try
            Output_new_excel_and_open_from_textbox(JMFileCho_Spec_TextBox.Text)
            msExcel_app.Visible = True

            Resize_JMForm(JMForm_size.re_size) '重新變大小
            'output_ToSpec.Spec_DWG(msExcel_workbook, msExcel_app)
            'output_ToSpec.Spec_SPEC_TW(LiftNum, ContainNum, msExcel_workbook, msExcel_app)
            'output_ToSpec.Spec_Important(msExcel_workbook, msExcel_app)
            'output_ToSpec.Spec_MMIC(msExcel_workbook, msExcel_app)

            Output_open_excel_folder_and_save_when_done(JMFileCho_Spec_TextBox)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.DWG_OutputButton_Click_1")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub

    Private Sub testBasic_Button_Click_1(sender As Object, e As EventArgs) Handles testBasic_Button.Click
        Try
            Output_new_excel_and_open_from_textbox(Load_Job_BasePath_ComboBox.Text)
            Resize_JMForm(JMForm_size.re_size) '重新變大小

            output_ToSpec.Spec_Spec_Std(msExcel_workbook, msExcel_app)
            Output_open_excel_folder_and_saveAs_when_done($"{Load_Job_OutputPath_TextBox.Text}\{Basic_JobNoNew_TextBox.Text}-SPEC",
                                                          Load_Job_OutputPath_TextBox.Text)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.testBasic_Button_Click_1")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub

    Private Sub testSpec_Button_Click_1(sender As Object, e As EventArgs) Handles testSpec_Button.Click
        Try
            Output_new_excel_and_open_from_textbox(Load_Job_BasePath_ComboBox.Text)
            Resize_JMForm(JMForm_size.re_size) '重新變大小

            output_ToSpec.Spec_SPEC_Basic(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_TW(LiftNum, ContainNum, msExcel_workbook, msExcel_app)
            Output_open_excel_folder_and_saveAs_when_done($"{Load_Job_OutputPath_TextBox.Text}\{Basic_JobNoNew_TextBox.Text}-SPEC",
                                                          Load_Job_OutputPath_TextBox.Text)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.testSpec_Button_Click_1")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub

    Private Sub testImp_Button_Click_1(sender As Object, e As EventArgs) Handles testImp_Button.Click
        Try
            Output_new_excel_and_open_from_textbox(Load_Job_BasePath_ComboBox.Text)
            Resize_JMForm(JMForm_size.re_size) '重新變大小

            output_ToSpec.Spec_Important(msExcel_workbook, msExcel_app)
            Output_open_excel_folder_and_saveAs_when_done($"{Load_Job_OutputPath_TextBox.Text}\{Basic_JobNoNew_TextBox.Text}-SPEC",
                                                          Load_Job_OutputPath_TextBox.Text)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.testImp_Button_Click_1")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub
    Private Sub testCheckList_Button_Click_1(sender As Object, e As EventArgs) Handles testCheckList_Button.Click
        Try
            Output_new_excel_and_open_from_textbox(Load_Job_BasePath_ComboBox.Text)
            Resize_JMForm(JMForm_size.re_size) '重新變大小

            output_ToSpec.Spec_CheckList(msExcel_workbook, msExcel_app)
            Output_open_excel_folder_and_saveAs_when_done($"{Load_Job_OutputPath_TextBox.Text}\{Basic_JobNoNew_TextBox.Text}-SPEC",
                                                          Load_Job_OutputPath_TextBox.Text)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.testCheckList_Button_Click_1")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub
    Private Sub testMMIC_Button_Click_1(sender As Object, e As EventArgs) Handles testMMIC_Button.Click
        Try
            Output_new_excel_and_open_from_textbox(Load_Job_BasePath_ComboBox.Text)
            Resize_JMForm(JMForm_size.re_size) '重新變大小

            output_ToSpec.Spec_MMIC(msExcel_workbook, msExcel_app)
            Output_open_excel_folder_and_saveAs_when_done($"{Load_Job_OutputPath_TextBox.Text}\{Basic_JobNoNew_TextBox.Text}-SPEC",
                                                          Load_Job_OutputPath_TextBox.Text)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.testMMIC_Button_Click_1")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()
        End Try
    End Sub

    ''' <summary>
    ''' [Load > 輸出 > All全部]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub All_OutputButton_Click(sender As Object, e As EventArgs) Handles All_OutputButton.Click
        Try
            Output_new_excel_and_open_from_textbox(Load_Job_BasePath_ComboBox.Text)
            'msExcel_app.Visible = True

            Resize_JMForm(JMForm_size.re_size) '重新變大小
            output_ToSpec.Spec_FinalCheck(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_Spec_Std(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_Basic(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_TW(LiftNum, ContainNum, msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_Important(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_MMIC(msExcel_workbook, msExcel_app)

            Output_open_excel_folder_and_saveAs_when_done($"{Load_Job_OutputPath_TextBox.Text}\{Basic_JobNoNew_TextBox.Text}-SPEC",
                                                          Load_Job_OutputPath_TextBox.Text)
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.All_OutputButton_Click")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox(ex.Message)
        Finally
            Output_kill_excel_when_done()

        End Try
    End Sub

    ''' <summary>
    ''' 完成輸出後打開目標Excel的資料夾
    ''' </summary>
    ''' <param name="openPath_textBox"></param>
    Private Sub Output_open_excel_folder_and_save_when_done(openPath_textBox As TextBox)
        msExcel_workbook.Save()
        Process.Start(Path.GetDirectoryName(openPath_textBox.Text))
        MsgBox("完成")
    End Sub

    Private Sub Output_open_excel_folder_and_saveAs_when_done(saveAs_FullPath As String, openFolder_Path As String)
        msExcel_workbook.SaveAs(saveAs_FullPath)
        Process.Start(Path.GetDirectoryName(openFolder_Path))
        MsgBox("完成",, "輸出Excel訊息")
    End Sub

    ''' <summary>
    ''' 完成輸出後完全Kill掉所有執行的Excel
    ''' </summary>
    Private Sub Output_kill_excel_when_done()
        msExcel_workbook.Close()
        msExcel_app.Quit()
        Dim excelPro As Process() = Process.GetProcessesByName("Excel")

        For Each mPro As Process In excelPro
            mPro.Kill()
        Next
    End Sub
    ''' <summary>
    ''' 新增一個Excel並開啟
    ''' </summary>
    ''' <param name="openPath_textBox"></param>
    Private Sub Output_new_excel_and_open_from_textbox(openPath_textBox As String)
        msExcel_app = New Excel.Application
        msExcel_workbook = msExcel_app.Workbooks.Open(openPath_textBox)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'for  test
        LoadStored_ProgressBar_Form.Show()
    End Sub
    '------------------------------------------------------------------------------------------------------------ LOAD分頁 -> 送狀分頁 

















    '基本 --------------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' [基本 > JobNo New]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_NoNew_TextBox_MouseClick(sender As Object, e As MouseEventArgs) Handles Basic_JobNoNew_TextBox.MouseClick
        With Basic_JobNoNew_TextBox
            .ForeColor = Color.Black
        End With
    End Sub
    ''' <summary>
    ''' [基本 > JobNo Old]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_JobNoOld_TextBox_MouseClick(sender As Object, e As MouseEventArgs) Handles Basic_JobNoOld_TextBox.MouseClick
        With Basic_JobNoOld_TextBox
            .ForeColor = Color.Black
        End With
    End Sub

    ''' <summary>
    ''' [基本 > JobNo Mod]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_JobNoMOD_TextBox_MouseClick(sender As Object, e As MouseEventArgs) Handles Basic_JobNoMOD_TextBox.MouseClick
        With Basic_JobNoMOD_TextBox
            .ForeColor = Color.Black
        End With
    End Sub

    ''' <summary>
    ''' [基本仕樣 > use CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_Basic_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_Basic_CheckBox.CheckedChanged
        '基本是否啟用
        use_basic_chkbox_clickTimes += 1

        If Use_Basic_CheckBox.Checked Then
            Basic_GroupBox.Enabled = True

            If use_basic_chkbox_clickTimes = 1 Then
                With Basic_JobNoNew_TextBox '基本->目前工番號
                    If .Text = "" Then
                        .ForeColor = Color.Gray
                        .Text = get_nameManager.STD_JobNo_New
                    End If
                End With
                With Basic_JobNoOld_TextBox '基本->舊工番號
                    If .Text = "" Then
                        .ForeColor = Color.Gray
                        .Text = get_nameManager.STD_JobNo_Old
                    End If
                End With
                With Basic_JobNoMOD_TextBox '基本->MOD工番號
                    If .Text = "" Then
                        .ForeColor = Color.Gray
                        .Text = get_nameManager.STD_JobNo_Mod
                    End If
                End With
                With Basic_DesingerChinese_ComboBox '基本->設計者名字
                    get_nameManager.read_DbmsData(get_nameManager.EmployeeChinese,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  Basic_DesingerChinese_ComboBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                    If .Items.Count <> 0 Then
                        .Text = get_nameManager.read_DbmsData_Employee_getRow(get_nameManager.EmployeeChinese,
                                                                              get_nameManager.SQLite_tableName_Basic,
                                                                              get_nameManager.EmployeeRow,
                                                                              get_nameManager.SQLite_connectionPath_Tool,
                                                                              get_nameManager.SQLite_ToolDBMS_Name)

                    End If
                End With
                With Basic_DesingerEnglish_ComboBox '基本->設計者英文名字
                    get_nameManager.read_DbmsData(get_nameManager.EmployeeEnglish,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  Basic_DesingerEnglish_ComboBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                End With
                With Basic_CheckerChinese_ComboBox '基本->覆核者名字
                    get_nameManager.read_DbmsData(get_nameManager.EmployeeChinese,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  Basic_CheckerChinese_ComboBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                End With
                With Basic_CheckerEnglish_ComboBox '基本->覆核者英文名字
                    get_nameManager.read_DbmsData(get_nameManager.EmployeeEnglish,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  Basic_CheckerEnglish_ComboBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                End With
                With Basic_ApproverChinese_ComboBox '基本->承認者名字
                    get_nameManager.read_DbmsData(get_nameManager.ApproverChinese,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  Basic_ApproverChinese_ComboBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                    If .Items.Count <> 0 Then
                        .Text = .Items(0)
                    End If
                End With
                With Basic_ApproverEnglish_ComboBox '基本->承認者英文名字
                    get_nameManager.read_DbmsData(get_nameManager.ApproverEnglish,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  Basic_ApproverEnglish_ComboBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                End With
                With Basic_Local_ComboBox '基本->地區名
                    get_nameManager.read_DbmsData(get_nameManager.Local,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  Basic_Local_ComboBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                End With
            End If
        Else
            Basic_GroupBox.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' [基本仕樣 > 中文 > 設計者]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_Desinger_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Basic_DesingerChinese_ComboBox.TextChanged
        If Basic_DesingerChinese_ComboBox.Text <> "" Or Basic_CheckerChinese_ComboBox.Text <> "" Or Basic_ApproverChinese_ComboBox.Text <> "" Then
            Basic_DesingerEnglish_ComboBox.Enabled = False
            Basic_CheckerEnglish_ComboBox.Enabled = False
            Basic_ApproverEnglish_ComboBox.Enabled = False
        Else
            Basic_DesingerEnglish_ComboBox.Enabled = True
            Basic_CheckerEnglish_ComboBox.Enabled = True
            Basic_ApproverEnglish_ComboBox.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' [基本仕樣 > 中文 > 檢查者]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_Checker_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Basic_CheckerChinese_ComboBox.TextChanged
        If Basic_DesingerChinese_ComboBox.Text <> "" Or Basic_CheckerChinese_ComboBox.Text <> "" Or Basic_ApproverChinese_ComboBox.Text <> "" Then
            Basic_DesingerEnglish_ComboBox.Enabled = False
            Basic_CheckerEnglish_ComboBox.Enabled = False
            Basic_ApproverEnglish_ComboBox.Enabled = False
        Else
            Basic_DesingerEnglish_ComboBox.Enabled = True
            Basic_CheckerEnglish_ComboBox.Enabled = True
            Basic_ApproverEnglish_ComboBox.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' [基本仕樣 > 中文 > 承認者]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_Approver_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Basic_ApproverChinese_ComboBox.TextChanged
        If Basic_DesingerChinese_ComboBox.Text <> "" Or Basic_CheckerChinese_ComboBox.Text <> "" Or Basic_ApproverChinese_ComboBox.Text <> "" Then
            Basic_DesingerEnglish_ComboBox.Enabled = False
            Basic_CheckerEnglish_ComboBox.Enabled = False
            Basic_ApproverEnglish_ComboBox.Enabled = False
        Else
            Basic_DesingerEnglish_ComboBox.Enabled = True
            Basic_CheckerEnglish_ComboBox.Enabled = True
            Basic_ApproverEnglish_ComboBox.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' [基本仕樣 > 英文 > 設計者]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_DesingerEnglish_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Basic_DesingerEnglish_ComboBox.TextChanged
        If Basic_DesingerEnglish_ComboBox.Text <> "" Or Basic_CheckerEnglish_ComboBox.Text <> "" Or Basic_ApproverEnglish_ComboBox.Text <> "" Then
            Basic_DesingerChinese_ComboBox.Enabled = False
            Basic_ApproverChinese_ComboBox.Enabled = False
            Basic_CheckerChinese_ComboBox.Enabled = False
        Else
            Basic_DesingerChinese_ComboBox.Enabled = True
            Basic_ApproverChinese_ComboBox.Enabled = True
            Basic_CheckerChinese_ComboBox.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' [基本 > 英文 > Checker]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_CheckerEnglish_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Basic_CheckerEnglish_ComboBox.TextChanged
        If Basic_DesingerEnglish_ComboBox.Text <> "" Or Basic_CheckerEnglish_ComboBox.Text <> "" Or Basic_ApproverEnglish_ComboBox.Text <> "" Then
            Basic_DesingerChinese_ComboBox.Enabled = False
            Basic_ApproverChinese_ComboBox.Enabled = False
            Basic_CheckerChinese_ComboBox.Enabled = False
        Else
            Basic_DesingerChinese_ComboBox.Enabled = True
            Basic_ApproverChinese_ComboBox.Enabled = True
            Basic_CheckerChinese_ComboBox.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' [基本 > 英文 > Approver]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_ApproverEnglish_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Basic_ApproverEnglish_ComboBox.TextChanged
        If Basic_DesingerEnglish_ComboBox.Text <> "" Or Basic_CheckerEnglish_ComboBox.Text <> "" Or Basic_ApproverEnglish_ComboBox.Text <> "" Then
            Basic_DesingerChinese_ComboBox.Enabled = False
            Basic_ApproverChinese_ComboBox.Enabled = False
            Basic_CheckerChinese_ComboBox.Enabled = False
        Else
            Basic_DesingerChinese_ComboBox.Enabled = True
            Basic_ApproverChinese_ComboBox.Enabled = True
            Basic_CheckerChinese_ComboBox.Enabled = True
        End If
    End Sub
    '-------------------------------------------------------------------------------------------------------------------- 基本 

    'Check List --------------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' [Check List > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_ChkList_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_ChkList_CheckBox.CheckedChanged
        'check list是否啟用
        use_chkList_chkbox_clickTimes += 1

        If Use_ChkList_CheckBox.Checked Then
            CheckList_GroupBox.Enabled = True

            'If use_chkList_chkbox_clickTimes = 1 Then
            '    ChkList_Confirm_CheckBox.Checked = True     '確認圖ChkBox
            '    ChkList_1_no_RadioButton.Checked = True     '1 不清楚仕樣
            '    ChkList_2_no_RadioButton.Checked = True     '2 法規、安全
            '    ChkList_3_no_RadioButton.Checked = True     '3 迴路圖面是否不清楚
            '    ChkList_5_no_RadioButton.Checked = True     '5 VONIC
            '    ChkList_6_no_RadioButton.Checked = True     '6 確認式樣動作
            '    ChkList_7_no_RadioButton.Checked = True     '7 參考資料
            '    ChkList_8_yes_RadioButton.Checked = True    '8 最後確認
            '    ChkList_8Item_RadioButton.Checked = True    '8 滿足特記事項
            '    ChkList_9_yes_RadioButton.Checked = True    '9 自我檢查表
            'End If
        Else
            CheckList_GroupBox.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' [CheckList > 品目明細日期 > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_PaSheet_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_PaSheet_CheckBox.CheckedChanged
        If ChkList_PaSheet_CheckBox.CheckState = CheckState.Checked Then
            ChkList_PaSheet_DateTimePicker.Enabled = False
            'MsgBox("year:" + usr_PaSheet_DateTimePicker.Value.Year.ToString() + "/month:" + usr_PaSheet_DateTimePicker.Value.Month.ToString() _
            '+ "/date:" + usr_PaSheet_DateTimePicker.Value.Day.ToString())
        Else
            ChkList_PaSheet_DateTimePicker.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' [CheckList > Order Spec > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_Os_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_OS_CheckBox.CheckedChanged
        If ChkList_OS_CheckBox.CheckState = CheckState.Checked Then
            ChkList_OS_DateTimePicker.Enabled = False
        Else
            ChkList_OS_DateTimePicker.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' [CheckList > 確認圖 > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_Confirm_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_Confirm_CheckBox.CheckedChanged
        If ChkList_Confirm_CheckBox.CheckState = CheckState.Checked Then
            ChkList_Confirm_DateTimePicker.Enabled = False
        Else
            ChkList_Confirm_DateTimePicker.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' [CheckList > 電器圖面 > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub usr_Elec_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_Elec_CheckBox.CheckedChanged
        If ChkList_Elec_CheckBox.CheckState = CheckState.Checked Then
            ChkList_Elec_DateTimePicker.Enabled = False
        Else
            ChkList_Elec_DateTimePicker.Enabled = True
        End If
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_1_no_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_1_no_RadioButton.CheckedChanged
        'If ChkList_1_no_RadioButton.Checked Then
        '    ChkList_1_yes_Content_TextBox.Enabled = False
        '    ChkList_1_yes_result_TextBox.Enabled = False
        'End If
    End Sub
    ''' <summary>
    ''' [CheckList > 1.主式樣 > 有，討論內容]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_1_yes_Content_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_1_yes_Content_TextBox.TextChanged
        If ChkList_1_yes_RadioButton.Checked = False Then
            If ChkList_1_yes_Content_TextBox.Text <> "" Or ChkList_1_yes_result_TextBox.Text <> "" Then
                ChkList_1_yes_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 1.主式樣 > 有，結果]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_1_yes_result_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_1_yes_result_TextBox.TextChanged
        If ChkList_1_yes_RadioButton.Checked = False Then
            If ChkList_1_yes_Content_TextBox.Text <> "" Or ChkList_1_yes_result_TextBox.Text <> "" Then
                ChkList_1_yes_RadioButton.Checked = True
            End If
        End If
    End Sub

    Private Sub ChkList_2_no_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_2_no_RadioButton.CheckedChanged
        'If ChkList_2_no_RadioButton.Checked Then
        '    ChkList_2_yes_Content_TextBox.Enabled = False
        '    ChkList_2_yes_Result_TextBox.Enabled = False
        'End If
    End Sub
    ''' <summary>
    ''' [CheckList > 2.法規問題 > 有，指出內容]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_2_yes_Content_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_2_yes_Content_TextBox.TextChanged
        If ChkList_2_yes_RadioButton.Checked = False Then
            If ChkList_2_yes_Content_TextBox.Text <> "" Or ChkList_2_yes_Result_TextBox.Text <> "" Then
                ChkList_2_yes_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 2.法規問題 > 有，結果]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_2_yes_Result_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_2_yes_Result_TextBox.TextChanged
        If ChkList_2_yes_RadioButton.Checked = False Then
            If ChkList_2_yes_Content_TextBox.Text <> "" Or ChkList_2_yes_Result_TextBox.Text <> "" Then
                ChkList_2_yes_RadioButton.Checked = True
            End If
        End If
    End Sub
    Private Sub ChkList_3_no_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_3_no_RadioButton.CheckedChanged
        'If ChkList_3_no_RadioButton.Checked Then
        '    ChkList_3_yes_Man_TextBox.Enabled = False
        '    ChkList_3_yes_Content_TextBox.Enabled = False
        '    ChkList_3_yes_Result_TextBox.Enabled = False
        'End If
    End Sub
    ''' <summary>
    ''' [CheckList > 3.電器不清楚 > 有，討論者]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_3_yes_Man_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_3_yes_Man_TextBox.TextChanged
        If ChkList_3_yes_RadioButton.Checked = False Then
            If ChkList_3_yes_Content_TextBox.Text <> "" Or
                ChkList_3_yes_Result_TextBox.Text <> "" Or
                ChkList_3_yes_Man_TextBox.Text <> "" Then
                ChkList_3_yes_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 3.電器不清楚 > 有，內容]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_3_yes_Content_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_3_yes_Content_TextBox.TextChanged
        If ChkList_3_yes_RadioButton.Checked = False Then
            If ChkList_3_yes_Content_TextBox.Text <> "" Or ChkList_3_yes_Result_TextBox.Text <> "" Or ChkList_3_yes_Man_TextBox.Text Then
                ChkList_3_yes_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 3.電器不清楚 > 有，結果]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_3_yes_Result_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_3_yes_Result_TextBox.TextChanged
        If ChkList_3_yes_RadioButton.Checked = False Then
            If ChkList_3_yes_Content_TextBox.Text <> "" Or ChkList_3_yes_Result_TextBox.Text <> "" Or ChkList_3_yes_Man_TextBox.Text Then
                ChkList_3_yes_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 5.VONIC > 無]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_5_no_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_5_no_RadioButton.CheckedChanged
        If ChkList_5_no_RadioButton.Checked Then
            ChkList_5_std_Content_TextBox.Text = ""
            ChkList_5_nstd_Content_TextBox.Text = ""
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 5.VONIC > 標準]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_5_std_Content_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_5_std_Content_TextBox.TextChanged
        If ChkList_5_std_RadioButton.Checked = False Then
            If ChkList_5_std_Content_TextBox.Text <> "" Then
                ChkList_5_std_RadioButton.Checked = True
                ChkList_5_nstd_Content_TextBox.Text = ""
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 5.VONIC > 工直]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_5_nstd_Content_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_5_nstd_Content_TextBox.TextChanged
        If ChkList_5_nstd_RadioButton.Checked = False Then
            If ChkList_5_nstd_Content_TextBox.Text <> "" Then
                ChkList_5_nstd_RadioButton.Checked = True
                ChkList_5_std_Content_TextBox.Text = ""
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 6.確認 > Check Sheet]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_6_yesChk_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_6_yesChk_RadioButton.CheckedChanged
        If ChkList_6_yesChk_RadioButton.Checked Then
            ChkList_6_yes_Content_TextBox.Text = ""
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 6.確認 > 有，檢驗項目]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_6_yes_Content_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_6_yes_Content_TextBox.TextChanged
        If ChkList_6_yes_Content_TextBox.Text <> "" Then
            ChkList_6_yes_RadioButton.Checked = True
            ChkList_6_yesItem_RadioButton.Checked = True
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 6.確認 > 無]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_6_no_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_6_no_RadioButton.CheckedChanged
        If ChkList_6_no_RadioButton.Checked Then
            ChkList_6_yesItem_RadioButton.Checked = False
            ChkList_6_yesChk_RadioButton.Checked = False
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 7.參考資料 > 有，文書]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_7_yes1_content_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_7_yes1_content_TextBox.TextChanged
        If ChkList_7_yes_RadioButton.Checked = False Then
            If ChkList_7_yes1_content_TextBox.Text <> "" Then
                ChkList_7_yes_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    '''  [CheckList > 7.參考資料 > 無]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_7_no_RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ChkList_7_no_RadioButton.CheckedChanged
        'If ChkList_7_no_RadioButton.Checked Then
        '    ChkList_7_yes1_content_TextBox.Enabled = False
        'End If
    End Sub
    '------------------------------------------------------------------------------------------------------------------- Check List

    '程式變更 --------------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' [程式變更 > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_Program_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_Program_CheckBox.CheckedChanged
        '程式變更是否啟用
        use_program_chkbox_clickTimes += 1

        If Use_Program_CheckBox.Checked Then
            ProgramChange_TabControl.Enabled = True
            'If use_program_chkbox_clickTimes = 1 Then
            '    PrmList_2_test_CheckBox.Checked = True     '測試裝置
            '    PrmList_3_debug_CheckBox.Checked = True    'DEBUG
            '    PrmList_3_confirm_CheckBox.Checked = True  '一般動作確認
            '    PrmList_3_excute_CheckBox.Checked = True   '確認程式執行
            '    PrmList_4_yes1_RadioButton.Checked = True  '4-1 手動全自動
            '    PrmList_4_yes2_RadioButton.Checked = True  '4-2 入出力點一致
            '    PrmList_4_yes3_RadioButton.Checked = True  '4-3 變數初始化
            '    PrmList_4_yes4_RadioButton.Checked = True  '4-4 OTHER的CASE
            '    PrmList_4_yes5_RadioButton.Checked = True  '4-5 ELSE IF
            '    PrmList_4_yes6_RadioButton.Checked = True  '4-6 LOOP
            '    PrmList_4_yes7_RadioButton.Checked = True  '4-7 範圍內
            '    PrmList_4_no8_RadioButton.Checked = True   '4-8 CASTING
            '    PrmList_4_no9_RadioButton.Checked = True   '4-9 0除
            '    PrmList_4_yes10_RadioButton.Checked = True '4-10 運算子
            '    PrmList_4_yes11_RadioButton.Checked = True '4-11 ADDRESS
            '    PrmList_4_yes12_RadioButton.Checked = True '4-12 要求仕樣
            'End If
        Else
            ProgramChange_TabControl.Enabled = False
        End If
    End Sub






    ''' <summary>
    ''' [程式變更 > 2.使用裝置 > 測試裝置]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PrmList_2_test_TextBox_TextChanged(sender As Object, e As EventArgs) Handles PrmList_2_test_TextBox.TextChanged
        If PrmList_2_test_CheckBox.Checked = False Then
            If PrmList_2_test_TextBox.Text <> "" Then
                PrmList_2_test_CheckBox.Checked = True
            End If
        Else
            If PrmList_2_test_TextBox.Text = "" Then
                PrmList_2_test_CheckBox.Checked = False
            End If
        End If
    End Sub
    Private Sub PrmList_2_test_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles PrmList_2_test_CheckBox.CheckedChanged
        If PrmList_2_test_CheckBox.CheckState = CheckState.Unchecked Then
            PrmList_2_test_TextBox.Text = ""
        End If
    End Sub

    ''' <summary>
    ''' [程式變更 > 2.使用裝置 > COP控制盤]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PrmList_2_COP_TextBox_TextChanged(sender As Object, e As EventArgs) Handles PrmList_2_COP_TextBox.TextChanged
        If PrmList_2_COP_CheckBox.Checked = False Then
            If PrmList_2_COP_TextBox.Text <> "" Then
                PrmList_2_COP_CheckBox.Checked = True
            End If
        Else
            If PrmList_2_COP_TextBox.Text = "" Then
                PrmList_2_COP_CheckBox.Checked = False
            End If
        End If
    End Sub

    Private Sub PrmList_2_COP_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles PrmList_2_COP_CheckBox.CheckedChanged
        If PrmList_2_COP_CheckBox.CheckState = CheckState.Unchecked Then
            PrmList_2_COP_TextBox.Text = ""
        End If
    End Sub
    ''' <summary>
    ''' [程式變更 > 2.使用裝置 > 研修塔]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PrmList_2_tower_TextBox_TextChanged(sender As Object, e As EventArgs) Handles PrmList_2_tower_TextBox.TextChanged
        If PrmList_2_Tower_CheckBox.Checked = False Then
            If PrmList_2_tower_TextBox.Text <> "" Then
                PrmList_2_Tower_CheckBox.Checked = True
            End If
        Else
            If PrmList_2_tower_TextBox.Text = "" Then
                PrmList_2_Tower_CheckBox.Checked = False
            End If
        End If
    End Sub

    Private Sub PrmList_2_Tower_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles PrmList_2_Tower_CheckBox.CheckedChanged
        If PrmList_2_Tower_CheckBox.CheckState = CheckState.Unchecked Then
            PrmList_2_tower_TextBox.Text = ""
        End If
    End Sub
    ''' <summary>
    ''' [程式變更 > 2.使用裝置 > 其他]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PrmList_2_other_TextBox_TextChanged(sender As Object, e As EventArgs) Handles PrmList_2_other_TextBox.TextChanged
        If PrmList_2_Other_CheckBox.Checked = False Then
            If PrmList_2_other_TextBox.Text <> "" Then
                PrmList_2_Other_CheckBox.Checked = True
            End If
        Else
            If PrmList_2_other_TextBox.Text = "" Then
                PrmList_2_Other_CheckBox.Checked = False
            End If
        End If
    End Sub
    Private Sub PrmList_2_Other_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles PrmList_2_Other_CheckBox.CheckedChanged
        If PrmList_2_Other_CheckBox.CheckState = CheckState.Unchecked Then
            PrmList_2_other_TextBox.Text = ""
        End If
    End Sub
    ''' <summary>
    ''' [程式變更 > 3.檢查方法 > 其他]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PrmList_3_other_TextBox_TextChanged(sender As Object, e As EventArgs) Handles PrmList_3_other_TextBox.TextChanged
        If PrmList_3_other_Checkbox.Checked = False Then
            If PrmList_3_other_TextBox.Text <> "" Then
                PrmList_3_other_Checkbox.Checked = True
            End If
        End If
    End Sub
    Private Sub PrmList_3_other_Checkbox_CheckedChanged(sender As Object, e As EventArgs) Handles PrmList_3_other_Checkbox.CheckedChanged
        If PrmList_3_other_Checkbox.CheckState = CheckState.Unchecked Then
            PrmList_3_other_TextBox.Text = ""
        End If
    End Sub
    '--------------------------------------------------------------------------------------------------------------------程式變更 


    '仕樣 -------------------------------------------------------------------------------------------------------------------- 
    ''' <summary>
    ''' [仕樣 > TW > NumericUpDown > 機種/控制方式 ]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_MachineType_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Spec_MachineType_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_LiftInfo()
        '機種
        AddSub_Object_Sub(Spec_MachineType_NumericUpDown,
                          Spec_MachineType_Panel,
                          {Spec_Base_ComboBox},
                          dyCtrlName.JobMaker_MachinTypeInfoName_Array.Count,
                          dyCtrlName.JobMaker_MachinTypeInfoName_Array,
                          {get_nameManager.SQLite_tableName_Basic},
                          {get_nameManager.Spec_MachineType})
        '控制方式
        AddSub_Object_Sub(Spec_MachineType_NumericUpDown,
                          Spec_ControlWay_Panel,
                          {Spec_Base_ComboBox},
                          dyCtrlName.JobMaker_ControlWayInfoName_Array.Count,
                          dyCtrlName.JobMaker_ControlWayInfoName_Array,
                          {get_nameManager.SQLite_tableName_Basic},
                          {get_nameManager.Spec_ControlWay})
    End Sub
    ''' <summary>
    ''' [仕樣 > TW > NumericUpDown > 用途]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_Purpose_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Spec_Purpose_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_LiftInfo()
        AddSub_Object_Sub(Spec_Purpose_NumericUpDown,
                          Spec_Purpose_Panel,
                          {Spec_Base_ComboBox},
                          dyCtrlName.JobMaker_PurposeInfoName_Array.Count,
                          dyCtrlName.JobMaker_PurposeInfoName_Array,
                          {get_nameManager.SQLite_tableName_Basic},
                          {get_nameManager.Spec_Purpose})
    End Sub
    ''' <summary>
    ''' [仕樣 > TW > NumericUpDown > FLEX-N ]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_FLEX_N_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Spec_FLEX_N_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_LiftInfo()
        AddSub_Object_Sub(Spec_FLEX_N_NumericUpDown,
                          Spec_FLEX_N_Panel,
                          {Spec_Base_ComboBox},
                          dyCtrlName.JobMaker_FLEXInfoName_Array.Count,
                          dyCtrlName.JobMaker_FLEXInfoName_Array,
                          {get_nameManager.SQLite_tableName_Basic},
                          {get_nameManager.FLEX})
    End Sub
    ''' <summary>
    ''' [仕樣 > TW > NumericUpDown >自家發 ]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_EmerNum_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Spec_EmerNum_NumericUpDown.ValueChanged
        Dim TitleLabel_name As String() = {"Group:", "號機名:", "避難階:", "回歸順序:", "繼續運轉號機:"}
        Dim TitleLabel_PosX As Integer() = {5, 70, 160, 5, 160}
        Dim TitleLable_PosY As Integer() = {10, 10, 10, 60, 60}
        Dim ContentTextBox_PosX As Integer() = {5, 70, 160, 5, 160}
        Dim ContentTextBox_PosY As Integer() = {30, 30, 30, 85, 85}
        Dim dyCtrlName As DynamicControlName = New DynamicControlName

        Dim emer_tabPage As TabPage
        Dim emer_Label As Label
        Dim emer_TextBox As TextBox
        Dim emer_groupNum As Integer
        Try
            emer_groupNum = Spec_EmerNum_NumericUpDown.Value
            'EMER_AUTO_TabControl.TabPages.Clear()
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.Spec_EmerNum_NumericUpDown_ValueChanged")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox("請輸入整數")
        End Try

        Dim EmerGroupNum_Panel_count, i_start As Integer
        EmerGroupNum_Panel_count = Spec_emerGroup_TabControl.TabPages.Count
        'Console.WriteLine($"EmerGroupNum_Panel_count:{EmerGroupNum_Panel_count}")
        If EmerGroupNum_Panel_count = 1 Then
            Spec_emerGroup_TabControl.TabPages.Clear()
            i_start = 1
        Else
            i_start = EmerGroupNum_Panel_count + 1
        End If

        dyCtrlName.JobMaker_EmerInfo()
        If emer_groupNum <= 10 Then
            If i_start > emer_groupNum Then
                For Each ctrlName As Control In Spec_emerGroup_TabControl.TabPages
                    If ctrlName.Name = $"{dyCtrlName.JobMaker_EMER_TabPage}_{i_start - 1}" Then
                        Spec_emerGroup_TabControl.TabPages.Remove(ctrlName)
                    End If
                Next
            Else
                For i = i_start To emer_groupNum
                    emer_tabPage = New TabPage '要自動生成的Tabpage

                    Spec_emerGroup_TabControl.TabPages.Add(emer_tabPage)

                    With emer_tabPage
                        .Text = i
                        .Name = ($"{dyCtrlName.JobMaker_EMER_TabPage}_{i}")
                    End With

                    For j = 1 To TitleLabel_name.Length
                        emer_Label = New Label
                        emer_TextBox = New TextBox

                        emer_tabPage.Controls.Add(emer_Label)
                        emer_tabPage.Controls.Add(emer_TextBox)

                        With emer_Label
                            .AutoSize = True
                            .Text = TitleLabel_name(j - 1)
                            '.BackColor = Color.Red
                            '.Name = ($"{dyCtrlName.JobMaker_EMER_LB}_{i}_{j}")
                            .Name = ($"{dyCtrlName.JobMaker_EmerLBInfoName_Array(j - 1)}_{i}")
                            .Location = New Point(TitleLabel_PosX(j - 1), TitleLable_PosY(j - 1))
                        End With


                        With emer_TextBox
                            If j <= 3 Then
                                .Width = emer_Label.Width
                            Else
                                .Width = emer_Label.Width + 50
                            End If
                            '.Name = ($"{dyCtrlName.JobMaker_EMER_TB}_{i}_{j}")
                            .Name = ($"{dyCtrlName.JobMaker_EmerTBInfoName_Array(j - 1)}_{i}")

                            Select Case .Name
                                Case dyCtrlName.Spec_EmerGroup_TextBox
                                    .Text = Chr(64 + i)
                                Case dyCtrlName.Spec_EmerCarName_TextBox
                                    .Text = ""
                                Case dyCtrlName.Spec_EmerEscapeFL_TextBox
                                    .Text = Spec_EscapeFL_TextBox.Text
                                Case dyCtrlName.Spec_EmerReturnFL_TextBox
                                    .Text = ""
                                Case dyCtrlName.Spec_EmerContinue_TextBox
                                    .Text = ""
                            End Select

                            .Location = New Point(ContentTextBox_PosX(j - 1), ContentTextBox_PosY(j - 1))
                        End With
                    Next
                Next
                'For Each mCrtl As Control In EMER_AUTO_TabControl.Controls
                '    For Each mmCrtl As Control In mCrtl.Controls
                '        MsgBox(mmCrtl.Name)
                '    Next
                'Next
            End If
        Else
            MsgBox("目前群數上限為10群")
        End If
    End Sub


    ''' <summary>
    ''' [仕樣 > Basic_all CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_SpecBasic_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_SpecBasic_CheckBox.CheckedChanged
        'basic式樣是否啟用
        use_spec_chkbox_clickTimes += 1

        If Use_SpecBasic_CheckBox.Checked Then

            SpecBasic_GroupBox.Enabled = True
            SpecBasic_GroupBox2.Enabled = True
            Use_SpecTWIDU_CheckBox.Enabled = True
            Use_SpecTWFP17_CheckBox.Enabled = True

            'SpecBasic_LiftItem_Panel.Enabled = True
            'SpecBasic_LiftItem_Dynamic_Panel.Enabled = True

            If Use_SpecTWIDU_CheckBox.Checked Or Use_SpecTWFP17_CheckBox.Checked Then
                'Spec_OutputButton.Enabled = True
            End If

            If use_spec_chkbox_clickTimes = 1 Then
                '基本 > 用途
                'get_nameManager.read_DbmsData(get_nameManager.Spec_Purpose,
                '                              get_nameManager.SQLite_tableName_Basic,
                '                              Spec_Purpose_ComboBox,
                '                              get_nameManager.SQLite_connectionPath_Tool,
                '                              get_nameManager.SQLite_ToolDBMS_Name)
                'TW > 機種
                'get_nameManager.read_DbmsData(get_nameManager.Spec_MachineType,
                '                              get_nameManager.SQLite_tableName_Basic,
                '                              Spec_MachineType_ComboBox,
                '                              get_nameManager.SQLite_connectionPath_Tool,
                '                              get_nameManager.SQLite_ToolDBMS_Name)

                'TW > 自家發入力點
                get_nameManager.read_DbmsData(get_nameManager.Spec_TW_EmerInput,
                                          get_nameManager.SQLite_tableName_Basic,
                                          Spec_EmerInput_ComboBox,
                                          get_nameManager.SQLite_connectionPath_Tool,
                                          get_nameManager.SQLite_ToolDBMS_Name)

                'TW > 自家發入力地址
                get_nameManager.read_DbmsData(get_nameManager.Spec_TW_EmerAddress,
                                          get_nameManager.SQLite_tableName_Basic,
                                          Spec_EmerAddress_ComboBox,
                                          get_nameManager.SQLite_connectionPath_Tool,
                                          get_nameManager.SQLite_ToolDBMS_Name)

            End If

        Else
            SpecBasic_GroupBox.Enabled = False
            SpecBasic_GroupBox2.Enabled = False
            With Use_SpecTWIDU_CheckBox
                .Enabled = False
                .CheckState = CheckState.Unchecked
            End With
            With Use_SpecTWFP17_CheckBox
                .Enabled = False
                .CheckState = CheckState.Unchecked
            End With

            'SpecBasic_LiftItem_Panel.Enabled = False
            'SpecBasic_LiftItem_Dynamic_Panel.Enabled = False

            If Use_SpecTWIDU_CheckBox.Checked <> False Or Use_SpecTWFP17_CheckBox.Checked <> False Then
                'Spec_OutputButton.Enabled = False
            End If
        End If
    End Sub
    ''' <summary>
    ''' [仕樣 > TW台灣 > FP-17 CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_SpecFP17_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_SpecTWFP17_CheckBox.CheckedChanged
        'FP-17式樣是否啟用
        If Use_SpecBasic_CheckBox.Checked Then
            If Use_SpecTWFP17_CheckBox.Checked Then
                Spec_TW_FlowLayoutPanel1.Enabled = True
                Spec_TW_FlowLayoutPanel2.Enabled = True
                Spec_TW_FlowLayoutPanel3.Enabled = True
                Spec_TW_FlowLayoutPanel4.Enabled = True
                Spec_TW_FlowLayoutPanel5.Enabled = True
                Spec_TW_FlowLayoutPanel6.Enabled = True
                Spec_TW_FlowLayoutPanel7.Enabled = True
                Spec_Base_ComboBox.Enabled = True '機種

                Use_SpecTWIDU_CheckBox.CheckState = CheckState.Unchecked

                'Spec_OutputButton.Enabled = True
                Spec_IF79x_Panel.Enabled = False    'IF79入出力位置
                Spec_EachStop_Panel.Enabled = False '各停開關
                Spec_WTB_Panel.Enabled = True       'WTB
                Spec_LoadCell_Panel.Enabled = True  'Load Cell

            Else
                If Use_SpecTWIDU_CheckBox.CheckState = CheckState.Unchecked Then
                    Spec_TW_FlowLayoutPanel1.Enabled = False
                    Spec_TW_FlowLayoutPanel2.Enabled = False
                    Spec_TW_FlowLayoutPanel3.Enabled = False
                    Spec_TW_FlowLayoutPanel4.Enabled = False
                    Spec_TW_FlowLayoutPanel5.Enabled = False
                    Spec_TW_FlowLayoutPanel6.Enabled = False
                    Spec_TW_FlowLayoutPanel7.Enabled = False
                End If
                'Spec_OutputButton.Enabled = False
            End If
        Else
            Spec_TW_FlowLayoutPanel1.Enabled = False
            Spec_TW_FlowLayoutPanel2.Enabled = False
            Spec_TW_FlowLayoutPanel3.Enabled = False
            Spec_TW_FlowLayoutPanel4.Enabled = False
            Spec_TW_FlowLayoutPanel5.Enabled = False
            Spec_TW_FlowLayoutPanel6.Enabled = False
            Spec_TW_FlowLayoutPanel7.Enabled = False
            Spec_Base_ComboBox.Enabled = False '機種
        End If
    End Sub
    ''' <summary>
    ''' [仕樣 > TW台灣 > IDU CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_SpecIDU_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_SpecTWIDU_CheckBox.CheckedChanged
        'idu式樣是否啟用
        If Use_SpecBasic_CheckBox.Checked Then
            If Use_SpecTWIDU_CheckBox.Checked Then
                Spec_TW_FlowLayoutPanel1.Enabled = True
                Spec_TW_FlowLayoutPanel2.Enabled = True
                Spec_TW_FlowLayoutPanel3.Enabled = True
                Spec_TW_FlowLayoutPanel4.Enabled = True
                Spec_TW_FlowLayoutPanel5.Enabled = True
                Spec_TW_FlowLayoutPanel6.Enabled = True
                Use_SpecTWFP17_CheckBox.CheckState = CheckState.Unchecked

                'Spec_OutputButton.Enabled = True
                Spec_IF79x_Panel.Enabled = True    'IF79入出力位置
                Spec_EachStop_Panel.Enabled = True '各停開關
                Spec_WTB_Panel.Enabled = True      'WTB
                Spec_LoadCell_Panel.Enabled = True 'Load Cell
                Spec_Base_ComboBox.Enabled = True '機種
            Else
                If Use_SpecTWFP17_CheckBox.CheckState = CheckState.Unchecked Then
                    Spec_TW_FlowLayoutPanel1.Enabled = False
                    Spec_TW_FlowLayoutPanel2.Enabled = False
                    Spec_TW_FlowLayoutPanel3.Enabled = False
                    Spec_TW_FlowLayoutPanel4.Enabled = False
                    Spec_TW_FlowLayoutPanel5.Enabled = False
                    Spec_TW_FlowLayoutPanel6.Enabled = False
                    Spec_TW_FlowLayoutPanel7.Enabled = False
                End If
            End If
        Else
            Spec_TW_FlowLayoutPanel1.Enabled = False
            Spec_TW_FlowLayoutPanel2.Enabled = False
            Spec_TW_FlowLayoutPanel3.Enabled = False
            Spec_TW_FlowLayoutPanel4.Enabled = False
            Spec_TW_FlowLayoutPanel5.Enabled = False
            Spec_TW_FlowLayoutPanel6.Enabled = False
            Spec_Base_ComboBox.Enabled = False '機種
        End If
    End Sub

    '[仕樣 > TW台灣 > Only Checkbox] =======================================================================================================
    ''' <summary>
    ''' 台灣式樣Only的CheckBox控制TextBox的Enable狀態
    ''' </summary>
    ''' <param name="chkbox"></param>
    ''' <param name="tbox"></param>
    Private Sub spec_onlyCheckbox_ctrlTextbox(chkbox As CheckBox, tbox As TextBox)
        If chkbox.Checked Then
            tbox.Enabled = True
        Else
            tbox.Enabled = False
        End If
    End Sub
    Private Sub Spec_PhotoEye_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_PhotoEye_Only_CheckBox.CheckedChanged
        '光電裝置
        spec_onlyCheckbox_ctrlTextbox(Spec_PhotoEye_Only_CheckBox, Spec_PhotoEye_Only_TextBox)
    End Sub
    Private Sub Spec_MechSafety_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_MechSafety_Only_CheckBox.CheckedChanged
        '機械裝置
        spec_onlyCheckbox_ctrlTextbox(Spec_MechSafety_Only_CheckBox, Spec_MechSafety_Only_TextBox)
    End Sub
    Private Sub Spec_SCOB_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_SCOB_Only_CheckBox.CheckedChanged
        '副COB
        spec_onlyCheckbox_ctrlTextbox(Spec_SCOB_Only_CheckBox, Spec_SCOB_Only_TextBox)
    End Sub
    Private Sub Spec_ION_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_ION_Only_CheckBox.CheckedChanged
        '離子除菌
        spec_onlyCheckbox_ctrlTextbox(Spec_ION_Only_CheckBox, Spec_ION_Only_TextBox)
    End Sub
    Private Sub Spec_AutoPass_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_AutoPass_Only_CheckBox.CheckedChanged
        '自動滿員通過
        spec_onlyCheckbox_ctrlTextbox(Spec_AutoPass_Only_CheckBox, Spec_AutoPass_Only_TextBox)
    End Sub
    Private Sub Spec_Indep_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Indep_Only_CheckBox.CheckedChanged
        '專用運轉
        spec_onlyCheckbox_ctrlTextbox(Spec_Indep_Only_CheckBox, Spec_Indep_Only_TextBox)
    End Sub
    Private Sub Spec_HinCpi_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_HinCpi_Only_CheckBox.CheckedChanged
        'HIN/CPI
        spec_onlyCheckbox_ctrlTextbox(Spec_HinCpi_Only_CheckBox, Spec_HinCpi_Only_TextBox)
    End Sub
    Private Sub Spec_Fire_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Fire_Only_CheckBox.CheckedChanged
        '火災管制
        spec_onlyCheckbox_ctrlTextbox(Spec_Fire_Only_CheckBox, Spec_Fire_Only_TextBox)
    End Sub
    Private Sub Spec_EscapeFL_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Fireman_Only_CheckBox.CheckedChanged
        '火災管制 > 避難階
        spec_onlyCheckbox_ctrlTextbox(Spec_Fireman_Only_CheckBox, Spec_Fireman_Only_TextBox)
    End Sub
    Private Sub Spec_CpiFM_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CpiFM_Only_CheckBox.CheckedChanged
        '車廂管制運轉-緊急
        spec_onlyCheckbox_ctrlTextbox(Spec_CpiFM_Only_CheckBox, Spec_CpiFM_Only_TextBox)
    End Sub
    Private Sub Spec_CpiOLT_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CpiOLT_Only_CheckBox.CheckedChanged
        '車廂管制運轉-滿載
        spec_onlyCheckbox_ctrlTextbox(Spec_CpiOLT_Only_CheckBox, Spec_CpiOLT_Only_TextBox)
    End Sub
    Private Sub Spec_HallGong_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_HallGong_Only_CheckBox.CheckedChanged
        '乘場到著鈴
        spec_onlyCheckbox_ctrlTextbox(Spec_HallGong_Only_CheckBox, Spec_HallGong_Only_TextBox)
    End Sub
    Private Sub Spec_Seismic_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Seismic_Only_CheckBox.CheckedChanged
        '地震管制運轉
        spec_onlyCheckbox_ctrlTextbox(Spec_Seismic_Only_CheckBox, Spec_Seismic_Only_TextBox)
    End Sub

    Private Sub Spec_SeismicSensor_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_SeismicSensor_Only_CheckBox.CheckedChanged
        '地震管制運轉-感知器
        spec_onlyCheckbox_ctrlTextbox(Spec_SeismicSensor_Only_CheckBox, Spec_SeismicSensor_Only_TextBox)
    End Sub

    Private Sub Spec_SeismicSW_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_SeismicSW_Only_CheckBox.CheckedChanged
        ''地震管制運轉-自動解除
        spec_onlyCheckbox_ctrlTextbox(Spec_SeismicSW_Only_CheckBox, Spec_SeismicSW_Only_TextBox)
    End Sub

    Private Sub Spec_Parking_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Parking_Only_CheckBox.CheckedChanged
        '停車運轉
        spec_onlyCheckbox_ctrlTextbox(Spec_Parking_Only_CheckBox, Spec_Parking_Only_TextBox)
    End Sub

    Private Sub Spec_CarGong_TopBtm_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_TopBtm_Only_CheckBox.CheckedChanged
        '車廂上到著鈴-Top Bottom
        spec_onlyCheckbox_ctrlTextbox(Spec_CarGong_TopBtm_Only_CheckBox, Spec_CarGong_TopBtm_Only_TextBox)
    End Sub

    Private Sub Spec_CarGong_Top_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_Top_Only_CheckBox.CheckedChanged
        '車廂上到著鈴-Top
        spec_onlyCheckbox_ctrlTextbox(Spec_CarGong_Top_Only_CheckBox, Spec_CarGong_Top_Only_TextBox)
    End Sub

    Private Sub Spec_CarGong_COB_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_COB_Only_CheckBox.CheckedChanged
        '車廂上到著鈴-COB
        spec_onlyCheckbox_ctrlTextbox(Spec_CarGong_COB_Only_CheckBox, Spec_CarGong_COB_Only_TextBox)
    End Sub

    Private Sub Spec_CarGong_VONIC_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_VONIC_Only_CheckBox.CheckedChanged
        '車廂上到著鈴-VONIC
        spec_onlyCheckbox_ctrlTextbox(Spec_CarGong_VONIC_Only_CheckBox, Spec_CarGong_VONIC_Only_TextBox)
    End Sub

    Private Sub Spec_ForceClose_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_ForceClose_Only_CheckBox.CheckedChanged
        '強制關門
        spec_onlyCheckbox_ctrlTextbox(Spec_ForceClose_Only_CheckBox, Spec_ForceClose_Only_TextBox)
    End Sub
    Private Sub Spec_WSCOB_only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_WSCOB_Only_CheckBox.CheckedChanged
        '殘障-副COB
        spec_onlyCheckbox_ctrlTextbox(Spec_WSCOB_Only_CheckBox, Spec_WSCOB_Only_TextBox)
    End Sub

    Private Sub Spec_WCOB_only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_WCOB_Only_CheckBox.CheckedChanged
        '殘障
        spec_onlyCheckbox_ctrlTextbox(Spec_WCOB_Only_CheckBox, Spec_WCOB_Only_TextBox)
    End Sub
    Private Sub Spec_HpiFM_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_HpiFM_Only_CheckBox.CheckedChanged
        '乘場信號-緊急
        spec_onlyCheckbox_ctrlTextbox(Spec_HpiFM_Only_CheckBox, Spec_HpiFM_Only_TextBox)
    End Sub

    Private Sub Spec_VonicBz_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_VonicBz_Only_CheckBox.CheckedChanged
        'VONIC蜂鳴器
        spec_onlyCheckbox_ctrlTextbox(Spec_VonicBz_Only_CheckBox, Spec_VonicBz_Only_TextBox)
    End Sub

    Private Sub Spec_DrHold_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_DrHold_Only_CheckBox.CheckedChanged
        '開門延遲按鈕
        spec_onlyCheckbox_ctrlTextbox(Spec_DrHold_Only_CheckBox, Spec_DrHold_Only_TextBox)
    End Sub

    Private Sub Spec_Landic_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Landic_Only_CheckBox.CheckedChanged
        'LANDIC
        spec_onlyCheckbox_ctrlTextbox(Spec_Landic_Only_CheckBox, Spec_Landic_Only_TextBox)
    End Sub
    Private Sub Spec_HLL_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_HLL_Only_CheckBox.CheckedChanged
        '乘場廳燈
        spec_onlyCheckbox_ctrlTextbox(Spec_HLL_Only_CheckBox, Spec_HLL_Only_TextBox)
    End Sub

    Private Sub Spec_ATT_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_ATT_Only_CheckBox.CheckedChanged
        '運轉首盤
        spec_onlyCheckbox_ctrlTextbox(Spec_ATT_Only_CheckBox, Spec_ATT_Only_TextBox)
    End Sub
    Private Sub Spec_LS1M_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_LS1M_Only_CheckBox.CheckedChanged
        'LS1m
        spec_onlyCheckbox_ctrlTextbox(Spec_LS1M_Only_CheckBox, Spec_LS1M_Only_TextBox)
    End Sub

    Private Sub Spec_PRU_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_PRU_Only_CheckBox.CheckedChanged
        'PRU
        spec_onlyCheckbox_ctrlTextbox(Spec_PRU_Only_CheckBox, Spec_PRU_Only_TextBox)
    End Sub

    Private Sub Spec_FrontRearDr_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_FrontRearDr_Only_CheckBox.CheckedChanged
        '正背門
        spec_onlyCheckbox_ctrlTextbox(Spec_FrontRearDr_Only_CheckBox, Spec_FrontRearDr_Only_TextBox)
    End Sub

    Private Sub Spec_OpeSw_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_OpeSw_Only_CheckBox.CheckedChanged
        '單群控切換
        spec_onlyCheckbox_ctrlTextbox(Spec_OpeSw_Only_CheckBox, Spec_OpeSw_Only_TextBox)
    End Sub
    Private Sub Spec_MFLReturn_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_MFLReturn_Only_CheckBox.CheckedChanged
        '基準階賦歸
        spec_onlyCheckbox_ctrlTextbox(Spec_MFLReturn_Only_CheckBox, Spec_MFLReturn_Only_TextBox)
    End Sub

    Private Sub Spec_MFLReturn_FL_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_MFLReturn_FL_Only_CheckBox.CheckedChanged
        '基準階賦歸-樓層
        spec_onlyCheckbox_ctrlTextbox(Spec_MFLReturn_FL_Only_CheckBox, Spec_MFLReturn_FL_Only_TextBox)
    End Sub

    Private Sub Spec_Vonic_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Vonic_Only_CheckBox.CheckedChanged
        'VONIC-VD10
        spec_onlyCheckbox_ctrlTextbox(Spec_Vonic_Only_CheckBox, Spec_Vonic_Only_TextBox)
    End Sub

    Private Sub Spec_Elvic_ParkingFL_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Elvic_ParkingFL_Only_CheckBox.CheckedChanged
        'ELVIC-停車樓層
        spec_onlyCheckbox_ctrlTextbox(Spec_Elvic_ParkingFL_Only_CheckBox, Spec_Elvic_ParkingFL_Only_TextBox)
    End Sub

    Private Sub Spec_LoadCellPos_CarBtm_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_LoadCellPos_CarBtm_Only_CheckBox.CheckedChanged
        'Load Cell 車廂下
        spec_onlyCheckbox_ctrlTextbox(Spec_LoadCellPos_CarBtm_Only_CheckBox, Spec_LoadCellPos_CarBtm_Only_TextBox)
    End Sub

    Private Sub Spec_LoadCellPos_MR_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_LoadCellPos_MR_Only_CheckBox.CheckedChanged
        'Load Cell 機房
        spec_onlyCheckbox_ctrlTextbox(Spec_LoadCellPos_MR_Only_CheckBox, Spec_LoadCellPos_MR_Only_TextBox)
    End Sub
    '=======================================================================================================[仕樣 > TW台灣 > Only Checkbox] 

    '[仕樣 > TW台灣 > 標題 ComboBox] =======================================================================================================
    ''' <summary>
    ''' [仕樣 > Only CheckBox的狀態 > 如果標頭Combox是X時，指定的Only CheckBox 將會unabled and uncheck]
    ''' </summary>
    ''' <param name="cb"></param>
    Private Sub Spec_onlyChkBox_state_to_unable_uncheck(cb As CheckBox)
        With cb
            If .Checked Then
                .CheckState = CheckState.Unchecked
            End If
            .Enabled = False
        End With
    End Sub
    Private Sub Spec_DRAuto_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_DRAuto_ComboBox.TextChanged
        '開門自動調節 > ComboBox
        If Spec_DRAuto_ComboBox.Text = get_nameManager.TB_O Then
            With Spec_PhotoEye_ComboBox
                .Enabled = True
                If .Text <> "" Then
                    Spec_PhotoEye_Only_CheckBox.Enabled = True
                End If
            End With
            With Spec_MechSafety_ComboBox
                .Enabled = True
                If .Text <> "" Then
                    Spec_MechSafety_Only_CheckBox.Enabled = True
                End If
            End With
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_PhotoEye_Only_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_MechSafety_Only_CheckBox)
            Spec_PhotoEye_ComboBox.Enabled = False
            Spec_MechSafety_ComboBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_PhotoEye_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_PhotoEye_ComboBox.TextChanged
        '開門自動調節 > 光電 > ComboBox
        If Spec_PhotoEye_ComboBox.Text = get_nameManager.TB_WITH Then
            Spec_PhotoEye_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_PhotoEye_Only_CheckBox)
        End If
    End Sub

    Private Sub Spec_MechSafety_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_MechSafety_ComboBox.TextChanged
        '開門自動調節 > 機械 > ComboBox
        If Spec_MechSafety_ComboBox.Text = get_nameManager.TB_WITH Then
            Spec_MechSafety_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_MechSafety_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_CancellCall_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_CancellCall_ComboBox.TextChanged
        '取消嬉戲呼叫 > ComboBox
        If Spec_CancellCall_ComboBox.Text = get_nameManager.TB_O Then
            With Spec_SCOB_ComboBox
                .Enabled = True
                If .Text <> "" Then
                    Spec_SCOB_Only_CheckBox.Enabled = True
                End If
            End With
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_SCOB_Only_CheckBox)
            Spec_SCOB_ComboBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_SCOB_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_SCOB_ComboBox.TextChanged
        '取消嬉戲呼叫 > 副COB > ComboBox
        If Spec_SCOB_ComboBox.Text = get_nameManager.TB_WITH Then
            Spec_SCOB_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_SCOB_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_AutoFan_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_AutoFan_ComboBox.TextChanged
        '風扇連動 > ComboBox
        If Spec_AutoFan_ComboBox.Text = get_nameManager.TB_O Then
            With Spec_ION_ComboBox
                .Enabled = True
                If .Text <> "" Then
                    Spec_ION_Only_CheckBox.Enabled = True
                End If
            End With
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_ION_Only_CheckBox)
            Spec_ION_ComboBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_ION_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_ION_ComboBox.TextChanged
        '風扇連動 > 離子除菌 > ComboBox
        If Spec_ION_ComboBox.Text = get_nameManager.TB_WITH Then
            Spec_ION_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_ION_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_AutoPass_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_AutoPass_ComboBox.TextChanged
        '自動滿員通過 > ComboBox
        If Spec_AutoPass_ComboBox.Text = get_nameManager.TB_O Then
            Spec_AutoPass_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_AutoPass_Only_CheckBox)
        End If
    End Sub

    Private Sub Spec_Indep_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Indep_ComboBox.TextChanged
        '專用運轉 > ComboBox
        If Spec_Indep_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Indep_Only_CheckBox.Enabled = True
            Spec_HpiIndep_ComboBox.Text = get_nameManager.TB_O
            Spec_WTB_Indep_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Indep_Only_CheckBox)
            Spec_HpiIndep_ComboBox.Text = get_nameManager.TB_X
            Spec_WTB_Indep_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub
    Private Sub Spec_HinCpi_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_HinCpi_ComboBox.TextChanged
        'HIN/CPI > ComboBox
        If Spec_HinCpi_ComboBox.Text = get_nameManager.TB_O Then
            Spec_HinCpi_Only_CheckBox.Enabled = True
            Spec_ParkingFL_COB_ComboBox.Text = get_nameManager.TB_O
            Spec_ParkingFL_HALL_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_HinCpi_Only_CheckBox)
            Spec_ParkingFL_COB_ComboBox.Text = get_nameManager.TB_X
            Spec_ParkingFL_HALL_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub
    Private Sub Spec_Fire_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Fire_ComboBox.TextChanged
        '火災管制 > ComboBox
        If Spec_Fire_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Fire_Only_CheckBox.Enabled = True
            Spec_FireSignal_ComboBox.Enabled = True
            Spec_EscapeFL_TextBox.Enabled = True

            Spec_CpiFire_ComboBox.Text = get_nameManager.TB_O
            Spec_WTB_FO_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Fire_Only_CheckBox)
            Spec_FireSignal_ComboBox.Enabled = False
            Spec_EscapeFL_TextBox.Enabled = False

            Spec_CpiFire_ComboBox.Text = get_nameManager.TB_X
            Spec_WTB_FO_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub

    Private Sub Spec_EscapeFL_Copy_Button_Click(sender As Object, e As EventArgs) Handles Spec_EscapeFL_Copy_Button.Click
        '火災管制 > 複製
        With Spec_EscapeFL_TextBox
            If .Text <> "" Then
                Spec_Parking_FL_TextBox.Text = .Text
                Spec_MFLReturn_FL_TextBox.Text = .Text
                Spec_Elvic_ParkingFL_TextBox.Text = .Text
                Spec_Flood_FL_TextBox.Text = .Text
                MsgBox("已經完成將避難階複製到停車階、基準階",, "複製Done")
            Else
                MsgBox("避難階是空值",, "error")
            End If
        End With
    End Sub

    Private Sub Spec_Fireman_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Fireman_ComboBox.TextChanged
        '消防梯 > ComboBox
        If Spec_Fireman_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Fireman_Only_CheckBox.Enabled = True
            Spec_CpiFM_ComboBox.Text = get_nameManager.TB_O
            Spec_HpiFM_ComboBox.Text = get_nameManager.TB_O
            Spec_WTB_FM_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Fireman_Only_CheckBox)
            Spec_CpiFM_ComboBox.Text = get_nameManager.TB_X
            Spec_HpiFM_ComboBox.Text = get_nameManager.TB_X
            Spec_WTB_FM_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub

    Private Sub Spec_Parking_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Parking_ComboBox.TextChanged
        '停車階 > ComboBox
        If Spec_Parking_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Parking_Only_CheckBox.Enabled = True
            Spec_Parking_FL_TextBox.Enabled = True
            Spec_ParkingFL_ELVIC_ComboBox.Enabled = True
            Spec_ParkingFL_WTB_ComboBox.Enabled = True
            Spec_ParkingFL_DR_ComboBox.Enabled = True
            Spec_ParkingFL_COB_ComboBox.Enabled = True
            Spec_ParkingFL_HALL_ComboBox.Enabled = True

            Spec_WTB_PKSW_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Parking_Only_CheckBox)
            Spec_Parking_FL_TextBox.Enabled = False
            Spec_ParkingFL_ELVIC_ComboBox.Enabled = False
            Spec_ParkingFL_WTB_ComboBox.Enabled = False
            Spec_ParkingFL_DR_ComboBox.Enabled = False
            Spec_ParkingFL_COB_ComboBox.Enabled = False
            Spec_ParkingFL_HALL_ComboBox.Enabled = False

            Spec_WTB_PKSW_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub

    Private Sub Spec_Seismic_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Seismic_ComboBox.TextChanged
        '地震管制 > ComboBox
        If Spec_Seismic_ComboBox.Text = get_nameManager.TB_O Then
            'Only
            Spec_Seismic_Only_CheckBox.Enabled = True
            '感知器
            Spec_SeismicSensor_ComboBox.Enabled = True
            If Spec_SeismicSensor_ComboBox.Text <> "" Then
                Spec_SeismicSensor_Only_CheckBox.Enabled = True
            End If
            '自動解除開關
            Spec_SeismicSW_ComboBox.Enabled = True
            If Spec_SeismicSW_ComboBox.Text <> "" Then
                Spec_SeismicSW_Only_CheckBox.Enabled = True
            End If

            Spec_CpiSeismic_ComboBox.Text = get_nameManager.TB_O
            Spec_WTB_EQ_ComboBox.Text = get_nameManager.TB_O
            Spec_WTB_EQIND_ComboBox.Text = get_nameManager.TB_O
            Spec_WTB_EQMac_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Seismic_Only_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_SeismicSensor_Only_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_SeismicSW_Only_CheckBox)
            Spec_SeismicSensor_ComboBox.Enabled = False
            Spec_SeismicSW_ComboBox.Enabled = False

            Spec_CpiSeismic_ComboBox.Text = get_nameManager.TB_X
            Spec_WTB_EQ_ComboBox.Text = get_nameManager.TB_X
            Spec_WTB_EQIND_ComboBox.Text = get_nameManager.TB_X
            Spec_WTB_EQMac_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub
    Private Sub Spec_SeismicSensor_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_SeismicSensor_ComboBox.TextChanged
        '地震管制運轉 > 感知器N段 ComboBox
        If Spec_SeismicSensor_ComboBox.Text = "" Then
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_SeismicSensor_Only_CheckBox)
        Else
            Spec_SeismicSensor_Only_CheckBox.Enabled = True
        End If
    End Sub
    Private Sub Spec_SeismicSW_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_SeismicSW_ComboBox.TextChanged
        '地震管制運轉 > 自動解除開關 ComboBox
        If Spec_SeismicSW_ComboBox.Text = get_nameManager.TB_O Then
            Spec_SeismicSW_Only_CheckBox.Enabled = True
            Spec_WTB_EQSW_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_SeismicSW_Only_CheckBox)
            Spec_WTB_EQSW_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub
    Private Sub Spec_CPI_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_CPI_ComboBox.TextChanged
        '車廂管制運轉燈 > ComboBox
        If Spec_CPI_ComboBox.Text = get_nameManager.TB_O Then
            Spec_CpiSeismic_ComboBox.Enabled = True
            Spec_CpiFire_ComboBox.Enabled = True
            Spec_CpiFM_ComboBox.Enabled = True
            Spec_CpiEmer_ComboBox.Enabled = True
            Spec_CpiOLT_ComboBox.Enabled = True
            If Spec_CpiFM_ComboBox.Text = get_nameManager.TB_O Then
                Spec_CpiFM_Only_CheckBox.Enabled = True
            End If
            If Spec_CpiOLT_ComboBox.Text = get_nameManager.TB_O Then
                Spec_CpiOLT_Only_CheckBox.Enabled = True
            End If
        Else
            Spec_CpiSeismic_ComboBox.Enabled = False
            Spec_CpiFire_ComboBox.Enabled = False
            Spec_CpiFM_ComboBox.Enabled = False
            Spec_CpiEmer_ComboBox.Enabled = False
            Spec_CpiOLT_ComboBox.Enabled = False
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CpiOLT_Only_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CpiFM_Only_CheckBox)
        End If
    End Sub

    Private Sub Spec_CpiFM_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_CpiFM_ComboBox.TextChanged
        '車廂管制運轉燈 > 緊急 > ComboBox
        If Spec_CpiFM_ComboBox.Text = get_nameManager.TB_O Then
            Spec_CpiFM_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CpiFM_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_CpiOLT_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_CpiOLT_ComboBox.TextChanged
        '車廂管制運轉燈 > 滿載 > ComboBox
        If Spec_CpiOLT_ComboBox.Text = get_nameManager.TB_O Then
            Spec_CpiOLT_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CpiOLT_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_HallGong_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_HallGong_ComboBox.TextChanged
        '乘場到著鈴聲 > ComboBox
        If Spec_HallGong_ComboBox.Text = get_nameManager.TB_O Then
            Spec_HallGong_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_HallGong_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_HPIMsg_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_HPIMsg_ComboBox.TextChanged
        '乘場信號文字 > ComboBox
        If Spec_HPIMsg_ComboBox.Text = get_nameManager.TB_O Then
            Spec_HpiOLT_ComboBox.Enabled = True
            Spec_HpiMain_ComboBox.Enabled = True
            Spec_HpiIndep_ComboBox.Enabled = True
            With Spec_HpiFM_ComboBox
                .Enabled = True
                If .Text <> "" Then
                    Spec_HpiFM_Only_CheckBox.Enabled = True
                End If
            End With
        Else
            Spec_HpiOLT_ComboBox.Enabled = False
            Spec_HpiMain_ComboBox.Enabled = False
            Spec_HpiIndep_ComboBox.Enabled = False
            Spec_HpiFM_ComboBox.Enabled = False
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_HpiFM_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_HpiFM_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_HpiFM_ComboBox.TextChanged
        '乘場信號文字 > 緊急 > ComboBox
        If Spec_HpiFM_ComboBox.Text = get_nameManager.TB_O Then
            Spec_HpiFM_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_HpiFM_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_CarGong_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_ComboBox.TextChanged
        '車廂上到著鈴 > ComboBox
        If Spec_CarGong_ComboBox.Text = get_nameManager.TB_O Then
            '車廂上
            With Spec_CarGong_Top_CheckBox
                .Enabled = True
                If .Checked Then
                    Spec_CarGong_Top_Only_CheckBox.Enabled = True
                End If
            End With
            '車廂上下
            With Spec_CarGong_TopBtm_CheckBox
                .Enabled = True
                If .Checked Then
                    Spec_CarGong_TopBtm_Only_CheckBox.Enabled = True
                End If
            End With
            'COB
            With Spec_CarGong_COB_CheckBox
                .Enabled = True
                If .Checked Then
                    Spec_CarGong_COB_Only_CheckBox.Enabled = True
                End If
            End With
            'VONIC
            With Spec_CarGong_VONIC_CheckBox
                .Enabled = True
                If .Checked Then
                    Spec_CarGong_VONIC_Only_CheckBox.Enabled = True
                End If
            End With
        Else
            '車廂上
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CarGong_Top_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CarGong_Top_Only_CheckBox)
            '車廂上下
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CarGong_TopBtm_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CarGong_TopBtm_Only_CheckBox)
            'COB
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CarGong_COB_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CarGong_COB_Only_CheckBox)
            'VONIC
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CarGong_VONIC_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_CarGong_VONIC_Only_CheckBox)
        End If
    End Sub

    Private Sub Spec_CRD_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_CRD_ComboBox.TextChanged
        '刷卡機 > ComboBox
        If Spec_CRD_ComboBox.Text = get_nameManager.TB_O Then
            Spec_CRDType_ComboBox.Enabled = True
            Spec_CRDID4_ComboBox.Enabled = True
            Spec_CRDID5_ComboBox.Enabled = True
        Else
            Spec_CRDType_ComboBox.Enabled = False
            Spec_CRDID4_ComboBox.Enabled = False
            Spec_CRDID5_ComboBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_ForceClose_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_ForceClose_ComboBox.TextChanged
        '強制關門
        If Spec_ForceClose_ComboBox.Text = get_nameManager.TB_O Then
            Spec_ForceClose_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_ForceClose_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_VonicBz_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_VonicBz_ComboBox.TextChanged
        'VONIC蜂鳴器 > ComboBox
        If Spec_VonicBz_ComboBox.Text = get_nameManager.TB_O Then
            Spec_VonicBz_Only_CheckBox.Enabled = True
            Spec_Vonic_ComboBox.Text = get_nameManager.TB_X
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_VonicBz_Only_CheckBox)
            Spec_Vonic_ComboBox.Text = get_nameManager.TB_O
        End If
    End Sub

    Private Sub Spec_DrHold_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_DrHold_ComboBox.TextChanged
        '開門延長按鈕 > ComboBox
        If Spec_DrHold_ComboBox.Text = get_nameManager.TB_O Then
            Spec_DrHold_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_DrHold_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_MFLReturn_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_MFLReturn_ComboBox.TextChanged
        '基準階賦歸 > ComboBox
        If Spec_MFLReturn_ComboBox.Text = get_nameManager.TB_O Then
            Spec_MFLReturn_Only_CheckBox.Enabled = True
            '基準階
            With Spec_MFLReturn_FL_TextBox
                .Enabled = True
                If .Text <> "" Then
                    Spec_MFLReturn_FL_Only_CheckBox.Enabled = True
                End If
            End With
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_MFLReturn_Only_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_MFLReturn_FL_Only_CheckBox)
            Spec_MFLReturn_FL_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_Vonic_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Vonic_ComboBox.TextChanged
        'VONIC語音撥放器 > ComboBox
        If Spec_Vonic_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Vonic_Only_CheckBox.Enabled = True
            Spec_Vonic_standard_ComboBox.Enabled = True
            Spec_VonicBz_ComboBox.Text = get_nameManager.TB_X
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Vonic_Only_CheckBox)
            Spec_Vonic_standard_ComboBox.Enabled = False
            Spec_VonicBz_ComboBox.Text = get_nameManager.TB_O
        End If
    End Sub
    Private Sub Spec_Emer_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Emer_ComboBox.TextChanged
        '自家發 > ComboBox
        If Spec_Emer_ComboBox.Text = get_nameManager.TB_O Then
            Spec_EmerNum_NumericUpDown.Enabled = True
            Spec_EmerSignal_ComboBox.Enabled = True
            Spec_EmerCapacity_NumericUpDown.Enabled = True
            Spec_EmerInput_ComboBox.Enabled = True
            Spec_EmerAddress_ComboBox.Enabled = True

            Spec_WTB_EmerPow_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_EmerNum_NumericUpDown.Enabled = False
            Spec_EmerSignal_ComboBox.Enabled = False
            Spec_EmerCapacity_NumericUpDown.Enabled = False
            Spec_EmerInput_ComboBox.Enabled = False
            Spec_EmerAddress_ComboBox.Enabled = False

            Spec_WTB_EmerPow_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub

    Private Sub Spec_Elvic_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Elvic_ComboBox.TextChanged
        'ELVIC > ComboBox
        If Spec_Elvic_ComboBox.Text = get_nameManager.TB_O Then
            '1.
            Spec_Elvic_Only_CheckBox.Enabled = True
            Spec_Elvic_Only_TextBox.Enabled = True
            With Spec_Elvic_Parking_CheckBox
                .Enabled = True
                If .Checked Then
                    Spec_Elvic_ParkingFL_TextBox.Enabled = True
                    Spec_Elvic_ParkingFL_Only_CheckBox.Enabled = True
                End If
            End With
            Spec_Elvic_FloorLockOut_CheckBox.Enabled = True
            Spec_Elvic_VIP_CheckBox.Enabled = True
            Spec_Elvic_Express_CheckBox.Enabled = True
            Spec_Elvic_Indep_CheckBox.Enabled = True
            Spec_Elvic_ReturnFL_CheckBox.Enabled = True
            '2.
            Spec_Elvic_Traffic_Peak_CheckBox.Enabled = True
            Spec_Elvic_Traffic_UpPeak_CheckBox.Enabled = True
            Spec_Elvic_Traffic_DownPeak_CheckBox.Enabled = True
            Spec_Elvic_Traffic_Lunch_CheckBox.Enabled = True
            Spec_Elvic_MainFL_CheckBox.Enabled = True
            Spec_Elvic_Zoning_CheckBox.Enabled = True
            Spec_Elvic_FloorLockOut_CheckBox.Enabled = True
            Spec_Elvic_FloorLockOut_GR_CheckBox.Enabled = True
            Spec_Elvic_CarCall_CheckBox.Enabled = True
            '3.
            Spec_Elvic_Fire_CheckBox.Enabled = True
            Spec_Elvic_Wavic_CheckBox.Enabled = True
            Spec_Elvic_CRD_CheckBox.Enabled = True
        Else
            '1.
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Elvic_Only_CheckBox)
            Spec_Elvic_Only_TextBox.Enabled = False
            Spec_Elvic_Parking_CheckBox.Enabled = False
            Spec_Elvic_ParkingFL_TextBox.Enabled = False
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Elvic_ParkingFL_Only_CheckBox)
            Spec_Elvic_FloorLockOut_CheckBox.Enabled = False
            Spec_Elvic_VIP_CheckBox.Enabled = False
            Spec_Elvic_Express_CheckBox.Enabled = False
            Spec_Elvic_Indep_CheckBox.Enabled = False
            Spec_Elvic_ReturnFL_CheckBox.Enabled = False
            '2.
            Spec_Elvic_Traffic_Peak_CheckBox.Enabled = False
            Spec_Elvic_Traffic_UpPeak_CheckBox.Enabled = False
            Spec_Elvic_Traffic_DownPeak_CheckBox.Enabled = False
            Spec_Elvic_Traffic_Lunch_CheckBox.Enabled = False
            Spec_Elvic_MainFL_CheckBox.Enabled = False
            Spec_Elvic_Zoning_CheckBox.Enabled = False
            Spec_Elvic_FloorLockOut_GR_CheckBox.Enabled = False
            Spec_Elvic_CarCall_CheckBox.Enabled = False
            '3.
            Spec_Elvic_Fire_CheckBox.Enabled = False
            Spec_Elvic_Wavic_CheckBox.Enabled = False
            Spec_Elvic_CRD_CheckBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_Elvic_Parking_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Elvic_Parking_CheckBox.CheckedChanged
        'ELVIC > Parking > CheckBox
        If Spec_Elvic_Parking_CheckBox.Checked Then
            With Spec_Elvic_ParkingFL_TextBox
                .Enabled = True
                If .Text <> "" Then
                    Spec_Elvic_ParkingFL_Only_CheckBox.Enabled = True
                End If
            End With
            'Spec_Elvic_ParkingFL_Only_TextBox.Enabled = True
        Else
            Spec_Elvic_ParkingFL_TextBox.Enabled = False
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Elvic_ParkingFL_Only_CheckBox)
            Spec_Elvic_ParkingFL_Only_TextBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_Elvic_ParkingFL_TextBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Elvic_ParkingFL_TextBox.TextChanged
        'ELVIC > Parking > TextBox
        If Spec_Elvic_ParkingFL_TextBox.Text <> "" Then
            Spec_Elvic_ParkingFL_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Elvic_ParkingFL_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_WCOB_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_WCOB_ComboBox.TextChanged
        '殘障 > ComboBox
        If Spec_WCOB_ComboBox.Text = get_nameManager.TB_O Then
            Spec_WCOB_Only_CheckBox.Enabled = True
            '殘障SCOB
            With Spec_WSCOB_ComboBox
                .Enabled = True
                If .Text <> "" Then
                    Spec_WSCOB_Only_CheckBox.Enabled = True
                End If
            End With
            '鳴動
            With Spec_WCOB_Ring_ComboBox
                .Enabled = True
                .SelectedIndex = 1
            End With
            '重要設定 > WHB
            Imp_WHB_ComboBox.SelectedIndex = 2
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_WCOB_Only_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_WSCOB_Only_CheckBox)
            Spec_WSCOB_ComboBox.Enabled = False
            Spec_WCOB_Ring_ComboBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_WSCOB_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_WSCOB_ComboBox.TextChanged
        '殘障 > 副COB > ComboBox
        If Spec_WSCOB_ComboBox.Text = get_nameManager.TB_O Then
            Spec_WSCOB_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_WSCOB_Only_CheckBox)
        End If

    End Sub
    Private Sub Spec_Landic_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Landic_ComboBox.TextChanged
        'Landic > ComboBox
        If Spec_Landic_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Landic_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_Landic_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_HLL_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_HLL_ComboBox.TextChanged
        '乘場廳燈 > ComboBox
        If Spec_HLL_ComboBox.Text = get_nameManager.TB_O Then
            Spec_HLL_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_HLL_Only_CheckBox)
        End If
    End Sub

    Private Sub Spec_ATT_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_ATT_ComboBox.TextChanged
        '運轉手盤 > ComboBox
        If Spec_ATT_ComboBox.Text = get_nameManager.TB_O Then
            Spec_ATT_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_ATT_Only_CheckBox)
        End If
    End Sub
    Private Sub Spec_Flood_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Flood_ComboBox.TextChanged
        '浸水管制 > ComboBox
        If Spec_Flood_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Flood_FL_TextBox.Enabled = True
        Else
            Spec_Flood_FL_TextBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_LS1M_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_LS1M_ComboBox.TextChanged
        'LS1M > ComboBox
        If Spec_LS1M_ComboBox.Text = get_nameManager.TB_O Then
            Spec_LS1M_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_LS1M_Only_CheckBox)
        End If
    End Sub

    Private Sub Spec_PRU_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_PRU_ComboBox.TextChanged
        '電力回升 > ComboBox
        If Spec_PRU_ComboBox.Text = get_nameManager.TB_O Then
            Spec_PRU_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_PRU_Only_CheckBox)
        End If
    End Sub

    Private Sub Spec_FrontRearDr_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_FrontRearDr_ComboBox.TextChanged
        '正背門 > ComboBox
        If Spec_FrontRearDr_ComboBox.Text = get_nameManager.TB_O Then
            Spec_FrontRearDr_Only_CheckBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_FrontRearDr_Only_CheckBox)
        End If
    End Sub


    Private Sub Spec_LoadCell_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_LoadCell_ComboBox.TextChanged
        'Load Cell > ComboBox
        If Spec_LoadCell_ComboBox.Text = get_nameManager.TB_O Then
            '車廂下
            Spec_LoadCellPos_CarBtm_CheckBox.Enabled = True
            '機房
            Spec_LoadCellPos_MR_CheckBox.Enabled = True
            'Spec_LoadCellPos_MR_TextBox.Enabled = True
        Else
            '車廂下
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_LoadCellPos_CarBtm_CheckBox)
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_LoadCellPos_CarBtm_Only_CheckBox)
            Spec_LoadCellPos_CarBtm_Only_TextBox.Enabled = False
            '機房
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_LoadCellPos_MR_CheckBox)
            Spec_LoadCellPos_MR_TextBox.Enabled = False
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_LoadCellPos_MR_Only_CheckBox)
            Spec_LoadCellPos_MR_Only_TextBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_LoadCellPos_CarBtm_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_LoadCellPos_CarBtm_CheckBox.CheckedChanged
        'Load Cell 車廂下
        If Spec_LoadCellPos_CarBtm_CheckBox.Checked Then
            Spec_LoadCellPos_CarBtm_Only_CheckBox.Enabled = True
            'Spec_LoadCellPos_CarBtm_Only_TextBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_LoadCellPos_CarBtm_Only_CheckBox)
            Spec_LoadCellPos_CarBtm_Only_TextBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_LoadCellPos_MR_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_LoadCellPos_MR_CheckBox.CheckedChanged
        'Load Cell 機房
        If Spec_LoadCellPos_MR_CheckBox.Checked Then
            Spec_LoadCellPos_MR_Only_CheckBox.Enabled = True
            Spec_LoadCellPos_MR_TextBox.Enabled = True
            'Spec_LoadCellPos_MR_Only_TextBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_LoadCellPos_MR_Only_CheckBox)
            Spec_LoadCellPos_MR_TextBox.Enabled = False
            Spec_LoadCellPos_MR_Only_TextBox.Enabled = False
        End If
    End Sub
    Private Sub Spec_OpeSw_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_OpeSw_ComboBox.TextChanged
        '單群控切換 > ComboBox
        If Spec_OpeSw_ComboBox.Text = get_nameManager.TB_O Then
            Spec_OpeSw_Only_CheckBox.Enabled = True
            Spec_OpeSw_DevicePos_TextBox.Enabled = True
            Spec_OpeSw_InputPos_ComboBox.Enabled = True
            Spec_OpeSw_InputAddress_TextBox.Enabled = True
        Else
            Spec_onlyChkBox_state_to_unable_uncheck(Spec_OpeSw_Only_CheckBox)
            Spec_OpeSw_DevicePos_TextBox.Enabled = False
            Spec_OpeSw_InputPos_ComboBox.Enabled = False
            Spec_OpeSw_InputAddress_TextBox.Enabled = False
        End If
    End Sub


    Private Sub Spec_WTB_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_WTB_ComboBox.TextChanged
        '監視盤 > ComboBox
        'If Spec_WTB_ComboBox.Text = get_nameManager.TB_X Then
        '    Spec_WTB_Error_ComboBox.Enabled = False
        '    Spec_WTB_Stop_ComboBox.Enabled = False
        '    Spec_WTB_FM_ComboBox.Enabled = False
        '    Spec_WTB_Normal_ComboBox.Enabled = False
        '    Spec_WTB_Urgent_ComboBox.Enabled = False
        '    Spec_WTB_FO_ComboBox.Enabled = False
        '    Spec_WTB_EmerPow_ComboBox.Enabled = False
        '    Spec_WTB_Alart_ComboBox.Enabled = False
        '    Spec_WTB_EQ_ComboBox.Enabled = False
        '    Spec_WTB_Indep_ComboBox.Enabled = False
        '    Spec_WTB_EQSW_ComboBox.Enabled = False
        '    Spec_WTB_BZSW_ComboBox.Enabled = False
        '    Spec_WTB_ChkSW_ComboBox.Enabled = False
        '    Spec_WTB_PKSW_ComboBox.Enabled = False
        '    Spec_WTB_EQIND_ComboBox.Enabled = False
        '    Spec_WTB_EQMac_ComboBox.Enabled = False
        'Else
        '    Spec_WTB_Error_ComboBox.Enabled = True
        '    Spec_WTB_Stop_ComboBox.Enabled = True
        '    Spec_WTB_FM_ComboBox.Enabled = True
        '    Spec_WTB_Normal_ComboBox.Enabled = True
        '    Spec_WTB_Urgent_ComboBox.Enabled = True
        '    Spec_WTB_FO_ComboBox.Enabled = True
        '    Spec_WTB_EmerPow_ComboBox.Enabled = True
        '    Spec_WTB_Alart_ComboBox.Enabled = True
        '    Spec_WTB_EQ_ComboBox.Enabled = True
        '    Spec_WTB_Indep_ComboBox.Enabled = True
        '    Spec_WTB_EQSW_ComboBox.Enabled = True
        '    Spec_WTB_BZSW_ComboBox.Enabled = True
        '    Spec_WTB_ChkSW_ComboBox.Enabled = True
        '    Spec_WTB_PKSW_ComboBox.Enabled = True
        '    Spec_WTB_EQIND_ComboBox.Enabled = True
        '    Spec_WTB_EQMac_ComboBox.Enabled = True
        'End If
    End Sub

    Private Sub Spec_ParkingFL_ELVIC_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_ParkingFL_ELVIC_ComboBox.TextChanged
        If Spec_ParkingFL_ELVIC_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Elvic_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_Elvic_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub
    Private Sub Spec_ParkingFL_WTB_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_ParkingFL_WTB_ComboBox.TextChanged
        If Spec_ParkingFL_WTB_ComboBox.Text = get_nameManager.TB_O Then
            Spec_WTB_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_WTB_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub
    Private Sub Spec_CpiEmer_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_CpiEmer_ComboBox.TextChanged
        If Spec_CpiEmer_ComboBox.Text = get_nameManager.TB_O Then
            Spec_Emer_ComboBox.Text = get_nameManager.TB_O
        Else
            Spec_Emer_ComboBox.Text = get_nameManager.TB_X
        End If
    End Sub


    '======================================================================================================= [仕樣 > TW台灣 > 標題 ComboBox] 


    ''' <summary>
    ''' [仕樣 > Basic All > 電梯總數]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_LiftNum_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Spec_LiftNum_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName

        Dim TextBoxWidth, TextBox_XPosition, TextBox_YPosition As Integer()

        Dim ConNum_tb As TextBox
        Dim ConNum_cb As ComboBox
        '顯示內容數量及文字

        TextBoxWidth = {Spec_LiftName_TextBox.Width, Spec_LiftMem_ComboBox.Width, Spec_Control_ComboBox.Width,
                        Spec_TopFL_ComboBox.Width, Spec_BtmFL_ComboBox.Width, Spec_StopFL_ComboBox.Width,
                        Spec_Speed_ComboBox.Width, Spec_FLName_TextBox.Width}
        TextBox_XPosition = {Spec_LiftName_TextBox.Left, Spec_LiftMem_ComboBox.Left, Spec_Control_ComboBox.Left,
                             Spec_TopFL_ComboBox.Left, Spec_BtmFL_ComboBox.Left, Spec_StopFL_ComboBox.Left,
                             Spec_Speed_ComboBox.Left, Spec_FLName_TextBox.Left}
        TextBox_YPosition = {Spec_LiftName_TextBox.Top, Spec_LiftMem_ComboBox.Top, Spec_Control_ComboBox.Top,
                             Spec_TopFL_ComboBox.Top, Spec_BtmFL_ComboBox.Top, Spec_StopFL_ComboBox.Top,
                             Spec_Speed_ComboBox.Top, Spec_FLName_TextBox.Top}


        '嘗試得到電梯輸入之總數
        Try
            LiftNum = Spec_LiftNum_NumericUpDown.Value
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.Spec_LiftNum_NumericUpDown_ValueChanged")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
        End Try


        Dim whetherCopy As Boolean '是否 複製#n號機名資訊
        Dim nLift_isCopy As Integer '要複製的號機
        '檢查是否需要複製功能 ----------------------------------------
        If Spec_LiftCopyInfo_CheckBox.Checked And LiftNum > 1 Then
            whetherCopy = True
            nLift_isCopy = Spec_Check_CopyLiftNum() '取得號機
        Else
            whetherCopy = False
        End If
        '---------------------------------------- 檢查是否需要複製功能 

        '動態生成
        dyCtrlName.JobMaker_LiftInfo()

        Dim LiftNum_Panel_count, i_start As Integer
        LiftNum_Panel_count = SpecBasic_LiftItem_Dynamic_Panel.Controls.Count
        If LiftNum_Panel_count = 0 Then
            i_start = 1
        Else
            i_start = LiftNum_Panel_count / SpecBasic_LiftItem_Panel.Controls.Count + 1
        End If



        If i_start > LiftNum Then
            '刪除 ----------------------------------
            For Each CtrlName_main As Control In SpecBasic_LiftItem_Panel.Controls
                For Each CtrlName_dynamic As Control In SpecBasic_LiftItem_Dynamic_Panel.Controls
                    If CtrlName_dynamic.Name = $"{CtrlName_main.Name}_{i_start - 1}" Then
                        SpecBasic_LiftItem_Dynamic_Panel.Controls.Remove(CtrlName_dynamic)
                    End If
                Next
            Next
            '---------------------------------- 刪除 
        Else
            '增加 ----------------------------------
            For i As Integer = i_start To LiftNum
                For Each ctrlName As Control In SpecBasic_LiftItem_Panel.Controls
                    ConNum_tb = New TextBox()
                    ConNum_cb = New ComboBox()
                    Select Case ctrlName.GetType
                        Case GetType(ComboBox)
                            With ConNum_cb
                                .Width = ctrlName.Width
                                .Left = ctrlName.Left
                                .Top = ctrlName.Top + (i - 1) * 30
                                .Font = New System.Drawing.Font("微軟正黑體",
                                                                9.0!,
                                                                System.Drawing.FontStyle.Regular,
                                                                System.Drawing.GraphicsUnit.Point,
                                                                CType(136, Byte))
                                .Name = $"{ctrlName.Name}_{i}"


                                Select Case ctrlName.Name
                                    Case Spec_LiftMem_ComboBox.Name '號機
                                        For Each item In Spec_LiftMem_ComboBox.Items
                                            .Items.Add(item)
                                        Next
                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_LiftMem_ComboBox, ConNum_cb, nLift_isCopy)
                                        Else
                                            .SelectedIndex = 0
                                        End If
                                        .TabIndex = 1 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                    Case Spec_Control_ComboBox.Name '操作方式
                                        For Each item In Spec_Control_ComboBox.Items
                                            .Items.Add(item)
                                        Next
                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_Control_ComboBox, ConNum_cb, nLift_isCopy)
                                        Else
                                            .SelectedIndex = 0
                                        End If
                                        .TabIndex = 2 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                    Case Spec_TopFL_ComboBox.Name '最高樓層的實際名稱
                                        For fl As Integer = 7 To 1 Step -1 'B7~B1 FL
                                            .Items.Add($"B{fl}")
                                        Next
                                        .Items.Add("G") 'G FL
                                        For fl As Integer = 1 To 32 '1~32 FL
                                            .Items.Add(fl)
                                        Next
                                        .Items.Add("R")
                                        .Items.Add("R1")
                                        .Items.Add("L")

                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_TopFL_ComboBox, ConNum_cb, nLift_isCopy)
                                        Else
                                            '.SelectedIndex = 7
                                            .Text = "8"
                                        End If
                                        .TabIndex = 3 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                    Case Spec_TopFL_Real_ComboBox.Name '最高樓層的制御階
                                        For fl As Integer = 1 To 39
                                            .Items.Add($"({fl})")
                                        Next
                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_TopFL_Real_ComboBox, ConNum_cb, nLift_isCopy)
                                        Else
                                            '.SelectedIndex = 7
                                            .Text = "(8)"
                                        End If
                                        .TabIndex = 4 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                    Case Spec_BtmFL_ComboBox.Name '最低樓層的實際名稱
                                        For fl As Integer = 7 To 1 Step -1 'B7~B1 FL
                                            .Items.Add($"B{fl}")
                                        Next
                                        .Items.Add("G") 'G FL
                                        For fl As Integer = 1 To 32 '1~32 FL
                                            .Items.Add(fl)
                                        Next
                                        .Items.Add("R")
                                        .Items.Add("R1")
                                        .Items.Add("L")
                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_BtmFL_ComboBox, ConNum_cb, nLift_isCopy)
                                        Else
                                            '.SelectedIndex = 1
                                            .Text = "1"
                                        End If
                                        .TabIndex = 5 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                    Case Spec_BtmFL_Real_ComboBox.Name '最低樓層的制御階
                                        For fl As Integer = 1 To 39
                                            .Items.Add($"({fl})")
                                        Next
                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_BtmFL_Real_ComboBox, ConNum_cb, nLift_isCopy)
                                        Else
                                            '.SelectedIndex = 1
                                            .Text = "(1)"
                                        End If
                                        .TabIndex = 6 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                    Case Spec_StopFL_ComboBox.Name '停止階
                                        For fl As Integer = 1 To 32
                                            .Items.Add(fl)
                                        Next
                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_StopFL_ComboBox, ConNum_cb, nLift_isCopy)
                                        Else
                                            .SelectedIndex = 7
                                        End If
                                        .TabIndex = 7 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                    Case Spec_Speed_ComboBox.Name '速度
                                        For Each item In Spec_Speed_ComboBox.Items
                                            .Items.Add(item)
                                        Next
                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_Speed_ComboBox, ConNum_cb, nLift_isCopy)
                                        Else
                                            .SelectedIndex = 0
                                        End If
                                        .TabIndex = 8 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                End Select

                            End With
                            SpecBasic_LiftItem_Dynamic_Panel.Controls.Add(ConNum_cb)
                        Case GetType(TextBox)
                            With ConNum_tb
                                .Width = ctrlName.Width
                                .Left = ctrlName.Left
                                .Top = ctrlName.Top + (i - 1) * 30
                                .Font = New System.Drawing.Font("微軟正黑體",
                                                                9.0!,
                                                                System.Drawing.FontStyle.Regular,
                                                                System.Drawing.GraphicsUnit.Point,
                                                                CType(136, Byte))
                                .Name = $"{ctrlName.Name}_{i}"

                                Select Case ctrlName.Name
                                    Case Spec_LiftName_TextBox.Name
                                        .Text = $"#{i}"
                                        .TabIndex = 0 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                    Case Spec_FLName_TextBox.Name
                                        If whetherCopy = True Then '複製nLift_isCopy號機
                                            .Text = Spec_LiftCopyInfo(Spec_FLName_TextBox, ConNum_tb, nLift_isCopy)
                                        Else
                                            .Text = "1-8"
                                        End If
                                        .TabIndex = 9 + (i - 1) * SpecBasic_LiftItem_Panel.Controls.Count
                                End Select
                            End With
                            SpecBasic_LiftItem_Dynamic_Panel.Controls.Add(ConNum_tb)
                    End Select
                Next
            Next i
            '---------------------------------- 增加 
        End If
    End Sub

    ''' <summary>
    ''' 檢查'生成號機名的內容' 與 '複製#n號機名資訊' 內文字是否相同，並回傳第N號機
    ''' </summary>
    ''' <returns></returns>
    Private Function Spec_Check_CopyLiftNum() As Integer
        Dim selectLiftNum As Integer
        For Each item1 As Control In SpecBasic_LiftItem_Dynamic_Panel.Controls
            For j As Integer = 1 To LiftNum
                If item1.Name = $"{Spec_LiftName_TextBox.Name}_{j}" And
                       item1.Text = Spec_LiftCopyInfo_TextBox.Text Then
                    selectLiftNum = j
                End If
            Next
        Next

        Return selectLiftNum
    End Function

    ''' <summary>
    ''' 複製第#n號機的該行指定combobox資訊
    ''' </summary>
    ''' <param name="base_ctrl">被當Base的combobox</param>
    ''' <param name="new_ctrl">新生成的combobox</param>
    ''' <param name="liftNum">第n號機</param>
    ''' <returns></returns>
    Private Function Spec_LiftCopyInfo(base_ctrl As Control, new_ctrl As Control, liftNum As Integer) As String
        Dim ctrl_text As String = ""
        For Each item2 As Control In SpecBasic_LiftItem_Dynamic_Panel.Controls
            If item2.Name = $"{base_ctrl.Name}_{liftNum}" Then
                ctrl_text = item2.Text
            End If
        Next
        Return ctrl_text
    End Function

    ''' <summary>
    ''' [仕樣 > TW台灣 > 車廂上到著鈴 > 車廂上]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_CarGong_Top_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_Top_CheckBox.CheckedChanged
        If Spec_CarGong_Top_CheckBox.Checked Then
            'Spec_CarGong_Top_TextBox.Enabled = True
            Spec_CarGong_Top_Only_CheckBox.Enabled = True
            If Spec_CarGong_Top_Only_CheckBox.CheckState = CheckState.Checked Then
                Spec_CarGong_Top_Only_TextBox.Enabled = True
            Else
                Spec_CarGong_Top_Only_TextBox.Enabled = False
            End If
        Else
            'Spec_CarGong_Top_TextBox.Enabled = False
            Spec_CarGong_Top_Only_CheckBox.Enabled = False
            Spec_CarGong_Top_Only_TextBox.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' [仕樣 > TW台灣 > 車廂上到著鈴 > 車廂上下]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_CarGong_TopBtm_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_TopBtm_CheckBox.CheckedChanged
        If Spec_CarGong_TopBtm_CheckBox.Checked Then
            Spec_CarGong_TopBtm_Only_CheckBox.Enabled = True
            If Spec_CarGong_TopBtm_Only_CheckBox.CheckState = CheckState.Checked Then
                Spec_CarGong_TopBtm_Only_TextBox.Enabled = True
            Else
                Spec_CarGong_TopBtm_Only_TextBox.Enabled = False
            End If
        Else
            Spec_CarGong_TopBtm_Only_CheckBox.Enabled = False
            Spec_CarGong_TopBtm_Only_TextBox.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' [仕樣 > TW台灣 > 車廂上到著鈴 > COB]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_CarGong_COB_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_COB_CheckBox.CheckedChanged
        If Spec_CarGong_COB_CheckBox.Checked Then
            Spec_CarGong_COB_Only_CheckBox.Enabled = True
            If Spec_CarGong_COB_Only_CheckBox.CheckState = CheckState.Checked Then
                Spec_CarGong_COB_Only_TextBox.Enabled = True
            Else
                Spec_CarGong_COB_Only_TextBox.Enabled = False
            End If
        Else
            Spec_CarGong_COB_Only_CheckBox.Enabled = False
            Spec_CarGong_COB_Only_TextBox.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' [仕樣 > TW台灣 > 車廂上到著鈴 > VONIC]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_CarGong_VONIC_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_VONIC_CheckBox.CheckedChanged
        If Spec_CarGong_VONIC_CheckBox.Checked Then
            Spec_CarGong_VONIC_Only_CheckBox.Enabled = True
            If Spec_CarGong_VONIC_Only_CheckBox.CheckState = CheckState.Checked Then
                Spec_CarGong_VONIC_Only_TextBox.Enabled = True
            Else
                Spec_CarGong_VONIC_Only_TextBox.Enabled = False
            End If
        Else
            Spec_CarGong_VONIC_Only_CheckBox.Enabled = False
            Spec_CarGong_VONIC_Only_TextBox.Enabled = False
        End If
    End Sub


    ''' <summary>
    ''' [仕樣 > Basic All > 基本功能Panel > Scroll功能]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_LiftItem_Panel_Scroll(sender As Object, e As ScrollEventArgs) Handles SpecBasic_LiftItem_Panel.Scroll
        Dim Spec_LiftItem_ScrollPanel_X As Long
        Spec_LiftItem_ScrollPanel_X = SpecBasic_LiftItem_Panel.AutoScrollPosition.X
        Dim p As Point = New Point(Math.Abs(Spec_LiftItem_ScrollPanel_X), 0)

        SpecBasic_LiftItem_Dynamic_Panel.AutoScrollPosition = p
    End Sub
    ''' <summary>
    ''' [仕樣 > Basic All > 自動生成Panel > Scroll功能]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_LiftItem_Dynamic_Panel_Scroll(sender As Object, e As ScrollEventArgs) Handles SpecBasic_LiftItem_Dynamic_Panel.Scroll
        Dim Spec_LiftItem_Dynamic_ScrollPanel_X As Long
        Spec_LiftItem_Dynamic_ScrollPanel_X = SpecBasic_LiftItem_Dynamic_Panel.AutoScrollPosition.X
        Dim p As Point = New Point(Math.Abs(Spec_LiftItem_Dynamic_ScrollPanel_X), 0)

        SpecBasic_LiftItem_Panel.AutoScrollPosition = p
    End Sub
    '-------------------------------------------------------------------------------------------------------------------- 仕樣 










    '送狀 ---------------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' [送狀 > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_prk_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_prk_CheckBox.CheckedChanged
        '送狀是否啟用?
        'use_DWG_chkbox_clickTimes += 1

        'If Use_prk_CheckBox.Checked Then
        '    DWG_GroupBox.Enabled = True

        '    If use_DWG_chkbox_clickTimes = 1 Then
        '        With DWG_PrkName_ComboBox
        '            get_nameManager.read_DbmsData(get_nameManager.PRK_Name,
        '                                          get_nameManager.SQLite_tableName_Basic,
        '                                          DWG_PrkName_ComboBox,
        '                                          get_nameManager.SQLite_connectionPath_Tool,
        '                                          get_nameManager.SQLite_ToolDBMS_Name)
        '        End With
        '    End If
        'Else
        '    DWG_GroupBox.Enabled = False
        'End If
    End Sub


    Private Function catalogPage_OUTPUT(a As Integer) As String
        '取得送狀名字(未使用)
        Dim catalogPageText_array As String()

        ReDim catalogPageText_array(clp_count)
        For i = 0 To clp_count - 1
            catalogPageText_array(i) = DWG_Page_CheckedListBox.Items(i).ToString
        Next
        Return catalogPageText_array(a)
    End Function

    ''' <summary>
    ''' [送狀 > 新增+ > Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub catalogPage_addButton_Click(sender As Object, e As EventArgs) Handles DWG_Page_AddButton.Click
        '送狀新增
        Dim PageNum As Integer
        Dim left_string As String
        left_string = Microsoft.VisualBasic.Left(Basic_JobNoNew_TextBox.Text, 7)
        Try
            PageNum = DWG_PageNum_TextBox.Text
            If PageNum <> Nothing And DWG_PrkName_ComboBox.Text <> Nothing Then
                For i As Integer = 1 To PageNum
                    DWG_Page_CheckedListBox.Items.Add($"{DWG_PrkName_ComboBox.Text}{i}/{PageNum}",
                                                      False)
                    DWG_Construction_CheckedListBox.Items.Add("", False)
                    DWG_Produce_CheckedListBox.Items.Add("", False)
                Next i
            End If
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.catalogPage_addButton_Click")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' [送狀 > 刪除- > Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub catalogPage_subButton_Click(sender As Object, e As EventArgs) Handles DWG_Page_SubButton.Click
        '送狀刪除
        With DWG_Page_CheckedListBox
            If .CheckedItems.Count > 0 Then
                For checked As Integer = .CheckedItems.Count - 1 To 0 Step -1
                    .Items.Remove(.CheckedItems(checked))
                Next
            End If
        End With
        With DWG_Construction_CheckedListBox
            If .CheckedItems.Count > 0 Then
                For checked As Integer = .CheckedItems.Count - 1 To 0 Step -1
                    .Items.Remove(.CheckedItems(checked))
                Next
            End If
        End With
        With DWG_Produce_CheckedListBox
            If .CheckedItems.Count > 0 Then
                For checked As Integer = .CheckedItems.Count - 1 To 0 Step -1
                    .Items.Remove(.CheckedItems(checked))
                Next
            End If
        End With
    End Sub

    ''' <summary>
    ''' [送狀 > 基本版型套用 > Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DWG_StdPage_Button_Click(sender As Object, e As EventArgs) Handles DWG_StdPage_Button.Click
        '基本版型套用
        DWG_Page_CheckedListBox.Items.Clear()
        DWG_Construction_CheckedListBox.Items.Clear()
        DWG_Produce_CheckedListBox.Items.Clear()

        If DWG_VonicStd_ComboBox.Text = get_nameManager.TB_X Then
            get_nameManager.read_DbmsData_catalogPage(get_nameManager.DWG_StdPage,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  DWG_Page_CheckedListBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
        ElseIf DWG_VonicStd_ComboBox.Text = get_nameManager.TB_O Then
            get_nameManager.read_DbmsData_catalogPage(get_nameManager.DWG_StdPage_withoutVonic,
                                                  get_nameManager.SQLite_tableName_Basic,
                                                  DWG_Page_CheckedListBox,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
        End If

        Try
            For i As Integer = 1 To DWG_Page_CheckedListBox.Items.Count
                DWG_Construction_CheckedListBox.Items.Add("", False)
                DWG_Produce_CheckedListBox.Items.Add("", False)
            Next i
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.DWG_StdPage_Button_Click")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' [送狀 > Check All > Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub catalogPage_ChkAllButton_Click(sender As Object, e As EventArgs) Handles DWG_Page_ChkAllButton.Click
        '送狀全打勾
        For i As Integer = 0 To DWG_Page_CheckedListBox.Items.Count - 1
            With DWG_Page_CheckedListBox
                .SetItemCheckState(i, CheckState.Checked)
            End With
        Next i
        For j As Integer = 0 To DWG_Construction_CheckedListBox.Items.Count - 1
            With DWG_Construction_CheckedListBox
                .SetItemCheckState(j, CheckState.Checked)
            End With
        Next j
        For k As Integer = 0 To DWG_Produce_CheckedListBox.Items.Count - 1
            With DWG_Produce_CheckedListBox
                .SetItemCheckState(k, CheckState.Checked)
            End With
        Next k
    End Sub

    ''' <summary>
    ''' [送狀 > Uncheck All > Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub catalogPage_unChkAllButton_Click(sender As Object, e As EventArgs) Handles DWG_Page_unChkAllButton.Click
        '送狀全不打勾
        For i As Integer = 0 To DWG_Page_CheckedListBox.Items.Count - 1
            With DWG_Page_CheckedListBox
                .SetItemCheckState(i, CheckState.Unchecked)
            End With
        Next i
        For j As Integer = 0 To DWG_Construction_CheckedListBox.Items.Count - 1
            With DWG_Construction_CheckedListBox
                .SetItemCheckState(j, CheckState.Unchecked)
            End With
        Next j
        For k As Integer = 0 To DWG_Produce_CheckedListBox.Items.Count - 1
            With DWG_Produce_CheckedListBox
                .SetItemCheckState(k, CheckState.Unchecked)
            End With
        Next k
    End Sub
    ''' <summary>
    ''' [送狀 >  輸出必要項目打勾 ]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DWG_Page_CheckedListBox_Click(sender As Object, e As EventArgs) Handles DWG_Page_CheckedListBox.Click
        Try
            If DWG_Page_CheckedListBox.Items.Count > 0 Then
                Dim index As Integer
                index = DWG_Page_CheckedListBox.SelectedIndex.ToString

                If DWG_Construction_CheckedListBox.Items.Count > 0 And
                   DWG_Produce_CheckedListBox.Items.Count > 0 Then
                    'MsgBox(index & "," & DWG_Page_CheckedListBox.GetItemCheckState(index))
                    If DWG_Page_CheckedListBox.GetItemCheckState(index) = 0 Then
                        DWG_Construction_CheckedListBox.SetItemCheckState(index,
                                                                         CheckState.Checked)
                        DWG_Produce_CheckedListBox.SetItemCheckState(index,
                                                                     CheckState.Checked)
                    Else
                        DWG_Construction_CheckedListBox.SetItemCheckState(index,
                                                                         CheckState.Unchecked)
                        DWG_Produce_CheckedListBox.SetItemCheckState(index,
                                                                     CheckState.Unchecked)
                    End If
                End If
            End If
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.DWG_Page_CheckedListBox_Click")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
        End Try
    End Sub
    '--------------------------------------------------------------------------------------------------------------------- 送狀 








    '重要設定 -------------------------------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' [重要設定 > Use_ImpIDU_CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_Important_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_Imp_CheckBox.CheckedChanged
        '重要設定是否啟用
        '------------ Hall Indicator HLL自動產生 -----------------------
        If Use_Imp_CheckBox.Checked Then
            If Use_SpecBasic_CheckBox.Checked And Spec_LiftNum_NumericUpDown.Value <> 0 Then '確認<基本仕樣>和<電梯總數>是否使用
                Dim lift_i, stopFL_i As Integer
                ReDim arr_liftName(LiftNum - 1) 'HIN中自動產生-<樓層名稱>
                ReDim arr_liftStopFL(LiftNum - 1) 'HIN中自動產生-<樓層停止數數量>
                'ReDim arr_liftTopFL(LiftNum - 1) 'HIN中自動產生-<樓層頂樓數量>

                Dim dyCtrlName As DynamicControlName = New DynamicControlName
                HallIndicator_FlowLayoutPanel.Controls.Clear() '每啟用就清除表單內容

                '讀取電梯的<樓層名稱>、<樓層停止數>等資訊 並 暫時儲存 ---------------------------------------------------
                For Each tempCtrl As Control In SpecBasic_LiftItem_Dynamic_Panel.Controls
                    For lift_i = 1 To LiftNum
                        '儲存目前自動產生的<樓層名稱> -----------------------
                        If tempCtrl.Name = $"{Spec_LiftName_TextBox.Name}_{lift_i}" Then
                            arr_liftName(lift_i - 1) = tempCtrl.Text
                        End If
                        '----------------------- 儲存目前自動產生的<樓層名稱> 

                        Try
                            '儲存目前自動產生的<樓層停止數> -----------------
                            If tempCtrl.Name = $"{Spec_StopFL_ComboBox.Name}_{lift_i}" Then
                                arr_liftStopFL(lift_i - 1) = CInt(tempCtrl.Text)
                            End If
                            '----------------- 儲存目前自動產生的<樓層停止數> 
                        Catch ex As Exception
                            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.Use_Important_CheckBox_CheckedChanged")
                            errorInfo.writeInfoError_InfoTxt(ex.Message)
                            MsgBox($"電梯停止數:{tempCtrl.Name},第{lift_i}號機 內容非數字")
                            ResultFailOutput_TextBox.Text += $"電梯停止數:{tempCtrl.Name},第{lift_i}號機 內容非數字"
                        End Try

                    Next
                Next
                '--------------------------------------------------- 讀取電梯的<樓層名稱>、<樓層停止數>等資訊 並 暫時儲存 

                '新建立<樓層名稱>、<樓層停止數>等資訊的控制項 ----------------------------------------------------------
                For lift_i = 1 To LiftNum
                    'FowLayoutPanel -------------------------------------------
                    Dim flowPanel As FlowLayoutPanel = New FlowLayoutPanel()
                    With flowPanel
                        .Width = 180
                        .Height = 380
                        .AutoScroll = True
                        .BorderStyle = BorderStyle.FixedSingle
                        .FlowDirection = FlowDirection.TopDown
                        .Name = $"{dyCtrlName.JobMaker_HIN_FlowPanel}_{lift_i}"
                        HallIndicator_FlowLayoutPanel.Controls.Add(flowPanel)
                    End With
                    '-------------------------------------------- FowLayoutPanel

                    '號機名 -------------------------------------------
                    Dim lb As Label = New Label()
                    lb.Text = arr_liftName(lift_i - 1)
                    flowPanel.Controls.Add(lb)
                    '-------------------------------------------號機名 

                    'All Check(勾選全樓層) -------------------------------------------
                    Dim AllFL_chkbox As CheckBox = New CheckBox()
                    With AllFL_chkbox
                        .Text = "全樓層都打勾"
                        .Name = $"{dyCtrlName.JobMaker_HIN_AllFL_ChkB}_{lift_i}"
                        AddHandler .CheckedChanged, AddressOf HIN_AllFL_CheckBox_SelectedIndexChanged
                        flowPanel.Controls.Add(AllFL_chkbox)
                    End With
                    '------------------------------------------- All Check(勾選全樓層)

                    '自動填入with/without.... ---------------------------------------------------------------------
                    Dim cho_chkbox As CheckBox = New CheckBox()
                    With cho_chkbox
                        .Name = $"{dyCtrlName.JobMaker_HIN_ChoAuto_ChkB}_{lift_i}"
                        .Text = ($"自動填入")
                        flowPanel.Controls.Add(cho_chkbox)
                    End With

                    Dim cho_cmbbox As ComboBox = New ComboBox()
                    With cho_cmbbox
                        .Name = $"{dyCtrlName.JobMaker_HIN_ChoAuto_CmbB}_{lift_i}"
                        ResultOutput_TextBox.Text += $"HIN各樓層名稱:{ .Name}{vbCrLf}"

                        get_nameManager.read_DbmsData(get_nameManager.IMP_HIN_FL_Content,
                                                      get_nameManager.SQLite_tableName_Basic,
                                                      cho_cmbbox,
                                                      get_nameManager.SQLite_connectionPath_Tool,
                                                      get_nameManager.SQLite_ToolDBMS_Name)

                        AddHandler .SelectedIndexChanged, AddressOf HIN_choAutoInsert_ComboBox_SelectedIndexChanged
                        flowPanel.Controls.Add(cho_cmbbox)

                        '輸出會用到 ---------------------------------------
                        ReDim arr_liftStopFl_StdContent(cho_cmbbox.Items.Count - 1)
                        ReDim arr_liftStopFl_EachContent(cho_cmbbox.Items.Count - 1, LiftNum)
                        For cnt_i = 1 To cho_cmbbox.Items.Count
                            arr_liftStopFl_StdContent(cnt_i - 1) = cho_cmbbox.Items(cnt_i - 1).ToString
                            arr_liftStopFl_EachContent(cnt_i - 1, 0) = cho_cmbbox.Items(cnt_i - 1).ToString
                        Next
                        '--------------------------------------- 輸出會用到 
                    End With
                    '--------------------------------------------------------------------- 自動填入with/without.... 

                    '分隔線 ProgressBar ------------------------------------------
                    Dim separate_proBar As ProgressBar = New ProgressBar()
                    With separate_proBar
                        .Width = 120
                        .Height = 10
                    End With
                    flowPanel.Controls.Add(separate_proBar)
                    '------------------------------------------ 分隔線 ProgressBar

                    '樓層CheckBox / With Combobox ---------------------------------------------------------
                    For stopFL_i = 1 To CInt(arr_liftStopFL(lift_i - 1))
                        Dim chkbox As CheckBox = New CheckBox()
                        Dim cmbBox As ComboBox = New ComboBox()

                        With chkbox
                            .AutoSize = True
                            .Text = stopFL_i & "FL(制御階)"
                            .Name = $"{stopFL_i}{dyCtrlName.JobMaker_HIN_FL_ChkB}_{lift_i}"
                        End With

                        With cmbBox
                            get_nameManager.read_DbmsData(get_nameManager.IMP_HIN_FL_Content,
                                                          get_nameManager.SQLite_tableName_Basic,
                                                          cmbBox,
                                                          get_nameManager.SQLite_connectionPath_Tool,
                                                          get_nameManager.SQLite_ToolDBMS_Name)
                            .Name = $"{stopFL_i}{dyCtrlName.JobMaker_HIN_FL_CmbB}_{lift_i}"
                        End With
                        flowPanel.Controls.Add(chkbox)
                        flowPanel.Controls.Add(cmbBox)
                    Next
                    '------------------------------------------------------ 樓層CheckBox / With Combobox 
                Next
                '---------------------------------------------------------- 新建立<樓層名稱>、<樓層停止數>等資訊的控制項 
            End If
            '--------- 確認<基本仕樣>和<電梯總數>是否使用

            Imp_OverBalance_ComboBox.SelectedIndex = 0

            ImpSetting_GroupBox.Enabled = True

        Else
            ImpSetting_GroupBox.Enabled = False

        End If
    End Sub
    ''' <summary>
    ''' [重要設定 > HIN > 將with/without...值填入每一個樓層的combobox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub HIN_choAutoInsert_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' 將HIN中自動產生的with/without combobox填入每一個樓層的combobox 的event -------------------------------
        Dim HIN_choAutoInsert_Text As String
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        Dim Lift_i As Integer = 1
        If LiftNum < 10 Then
            Lift_i = CInt(Strings.Right(sender.name, 1))
        Else
            Lift_i = CInt(Strings.Right(sender.name, 2))
        End If
        If Use_Imp_CheckBox.CheckState = CheckState.Checked Then
            For Each flp In HallIndicator_FlowLayoutPanel.Controls.OfType(Of FlowLayoutPanel)
                If flp.Name = $"{dyCtrlName.JobMaker_HIN_FlowPanel}_{Lift_i}" Then
                    For Each chkb In flp.Controls.OfType(Of CheckBox)
                        'For Lift_i = 1 To LiftNum
                        For stop_i = 1 To CInt(arr_liftStopFL(Lift_i - 1))
                            If chkb.Name = $"{dyCtrlName.JobMaker_HIN_ChoAuto_ChkB}_{Lift_i}" And chkb.Checked Then
                                For Each cb In flp.Controls.OfType(Of ComboBox)
                                    If cb.Name = $"{dyCtrlName.JobMaker_HIN_ChoAuto_CmbB}_{Lift_i}" Then
                                        HIN_choAutoInsert_Text = cb.Text
                                    ElseIf cb.Name = $"{stop_i}{dyCtrlName.JobMaker_HIN_FL_CmbB}_{Lift_i}" Then
                                        cb.Text = HIN_choAutoInsert_Text
                                    End If
                                Next
                            End If
                        Next 'stop_i
                        'Next 'lift_i
                    Next 'chkb.
                End If 'flp.name
            Next 'flp
        End If
        '------------------------------- 將HIN中自動產生的with/without combobox填入每一個樓層的combobox 的event 
    End Sub

    ''' <summary>
    ''' [重要設定 > HIN > 可以將全樓層CheckBox都一次打勾 ]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub HIN_AllFL_CheckBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' HIN中自動產生的<全樓層打勾>CheckBox 的event -------------------------------
        Dim HIN_AllFl_bool As Boolean
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        Dim Lift_i As Integer = 1
        If LiftNum < 10 Then
            Lift_i = CInt(Strings.Right(sender.name, 1))
        Else
            Lift_i = CInt(Strings.Right(sender.name, 2))
        End If
        If Use_Imp_CheckBox.CheckState = CheckState.Checked Then
            For Each flp In HallIndicator_FlowLayoutPanel.Controls.OfType(Of FlowLayoutPanel)
                If flp.Name = $"{dyCtrlName.JobMaker_HIN_FlowPanel}_{Lift_i}" Then
                    For Each chkb In flp.Controls.OfType(Of CheckBox)
                        'For Lift_i = 1 To LiftNum
                        For stop_i = 1 To CInt(arr_liftStopFL(Lift_i - 1))
                            '<全樓層都打勾> 動作時跳出迴圈避免資源浪費 ----------------------------------------------
                            If chkb.Name = $"{dyCtrlName.JobMaker_HIN_AllFL_ChkB}_{Lift_i}" Then
                                If chkb.Checked Then
                                    HIN_AllFl_bool = True
                                    Exit For
                                ElseIf chkb.Checked = False Then
                                    HIN_AllFl_bool = False
                                    Exit For
                                End If
                            End If
                            '---------------------------------------------- <全樓層都打勾> 動作時跳出迴圈避免資源浪費 

                            If chkb.Name = $"{stop_i}{dyCtrlName.JobMaker_HIN_FL_ChkB}_{Lift_i}" Then
                                If HIN_AllFl_bool Then
                                    chkb.Checked = True
                                Else
                                    chkb.Checked = False
                                End If
                            End If
                        Next 'stop_i

                        '<全樓層都打勾> 動作時跳出迴圈避免資源浪費 ----------------------------------------------
                        If chkb.Name = $"{dyCtrlName.JobMaker_HIN_AllFL_ChkB}_{Lift_i}" Then
                            If chkb.Checked Then
                                'Exit For
                            Else
                                'Exit For
                            End If
                        End If
                        '---------------------------------------------- <全樓層都打勾> 動作時跳出迴圈避免資源浪費 
                        'Next 'lift_i
                    Next 'chkb
                End If ' flp.Name
            Next 'flp
        End If
        '------------------------------- HIN中自動產生的<全樓層打勾>CheckBox 的event 
    End Sub
    '------------------------------------------------------------------------------------------------------------------------- 重要設定 



    ' MMIC -------------------------------------------------------------------------------------------------------------------------
    'Private Sub MMIC_VD10_Base_TextBox_KeyPress(sender As Object, e As EventArgs) Handles MMIC_VD10_Base_TextBox.KeyPress
    '    ChkList_5_nstd_Content_TextBox.Text = MMIC_VD10_Base_TextBox.Text
    'End Sub
    ''' <summary>
    ''' [MMIC > Use_mmic_CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_mmic_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_mmic_CheckBox.CheckedChanged
        'MMIC是否啟用
        use_mmic_chkbox_clickTimes += 1

        If Use_mmic_CheckBox.Checked Then

            MMIC_GroupBox.Enabled = True

            If use_mmic_chkbox_clickTimes = 1 Then
                '寫入機種,N幾百,eeprom data預設名稱 ----------------------------------------
                With MMIC_MR_CP43x_ComboBox
                    .Items.Add(get_nameManager.TB_WITH)
                    .Items.Add(get_nameManager.TB_WITHOUT)
                End With

                '[MMIC > 機種 Combobox]
                get_nameManager.read_DbmsData(get_nameManager.AllMachineType,
                                              get_nameManager.SQLite_tableName_Basic,
                                              MMIC_MachineType_ComboBox,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)
                '[MMIC > FLEX-N幾百 Combobox]
                get_nameManager.read_DbmsData(get_nameManager.FLEX,
                                              get_nameManager.SQLite_tableName_Basic,
                                              MMIC_FLEX_N_ComboBox,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)

                '[MMIC > EEPROM DATA > Base Combobox]
                get_nameManager.read_DbmsData(get_nameManager.mmicEEPROM_Base,
                                              get_nameManager.SQLite_tableName_Basic,
                                              MMIC_MR_EBase_ComboBox,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)
                '[MMIC > EEPROM DATA > TW Combobox]
                get_nameManager.read_DbmsData(get_nameManager.mmicEEPROM_DataName,
                                              get_nameManager.SQLite_tableName_Basic,
                                              MMIC_MR_ECarObj_ComboBox,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)
                MMIC_MR_ECarObj_ComboBox.Items.Add($"{Strings.Left(Basic_JobNoNew_TextBox.Text, 7)} MRA")

                '[SV > EEPROM DATA > Base Combobox]
                get_nameManager.read_DbmsData(get_nameManager.gspEEPROM_Base,
                                              get_nameManager.SQLite_tableName_Basic,
                                              MMIC_SV_EBase_ComboBox,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)
                '[SV > EEPROM DATA > TW Combobox]
                get_nameManager.read_DbmsData(get_nameManager.gspEEPROM_DataName,
                                              get_nameManager.SQLite_tableName_Basic,
                                              MMIC_SV_ECarObj_ComboBox,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)
                MMIC_SV_ECarObj_ComboBox.Items.Add($"{Strings.Left(Basic_JobNoNew_TextBox.Text, 7)} GSPA")

                '[SV > Flash Rom > Type Combobox]
                get_nameManager.read_DbmsData(get_nameManager.gspTypeName_Array,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              MMIC_SV_Type_ComboBox,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)

                '[VD10 > Base Combobox]
                get_nameManager.read_DbmsData(get_nameManager.VD10TypeName_Array,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              MMIC_VD10_Type_ComboBox,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)
                '---------------------------------------- 寫入機種,N幾百,eeprom data預設名稱 
            Else
                '如果有修改工番名字時，與MMIC的EEPROM DATA預設值不同的話則會修正 --------------------------------
                If $"{Strings.Left(Basic_JobNoNew_TextBox.Text, 7)} MRA" <>
                     MMIC_MR_ECarObj_ComboBox.Items(MMIC_MR_ECarObj_ComboBox.Items.Count - 1).ToString Then
                    MMIC_MR_ECarObj_ComboBox.Items.RemoveAt(MMIC_MR_ECarObj_ComboBox.Items.Count - 1)
                    MMIC_MR_ECarObj_ComboBox.Items.Add($"{Strings.Left(Basic_JobNoNew_TextBox.Text, 7)} MRA")
                End If

                If $"{Strings.Left(Basic_JobNoNew_TextBox.Text, 7)} GSPA" <>
                     MMIC_SV_ECarObj_ComboBox.Items(MMIC_SV_ECarObj_ComboBox.Items.Count - 1).ToString Then
                    MMIC_SV_ECarObj_ComboBox.Items.RemoveAt(MMIC_SV_ECarObj_ComboBox.Items.Count - 1)
                    MMIC_SV_ECarObj_ComboBox.Items.Add($"{Strings.Left(Basic_JobNoNew_TextBox.Text, 7)} GSPA")
                End If
                '-------------------------------- 如果有修改工番名字時，與MMIC的EEPROM DATA預設值不同的話則會修正 
            End If


            Select Case localSelect
                Case Taiwan
                    With MMIC_MR_ECarObj_ComboBox
                        If .Items.Count <> 0 Then
                            .SelectedIndex = MMIC_MR_ECarObj_ComboBox.Items.Count - 1
                        End If
                    End With
                    With MMIC_VD10_Type_ComboBox
                        If .Items.Count <> 0 Then
                            .SelectedIndex = 1
                        End If
                    End With
                    With MMIC_VD10_ROM_ComboBox
                        If .Items.Count <> 0 Then
                            .SelectedIndex = 1
                        End If
                    End With
                    With MMIC_VD10_Quantity_ComboBox
                        If .Items.Count <> 0 Then
                            .SelectedIndex = 0
                        End If
                    End With
            End Select
        Else
            MMIC_GroupBox.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' [MMIC > 機種 ComboBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub mmicType_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MMIC_MachineType_ComboBox.SelectedIndexChanged
        '機種選定後，底下的M-MIC和EEPROM DATA開啟

        If MMIC_MachineType_ComboBox.Text <> "" Then
            MMIC_Panel.Enabled = True

            If MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_IDU_ZT_TW,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If IDU(ZEXIA-T/TW) Then TJAMB61K
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_IDU_ZT_TW,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_IDU_RT_TW,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If IDU(REXIA-T/TW) Then ABM7143C
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_IDU_RT_TW,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_FP17_ZR_TW,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF FP-17(ZR/TW) THEN TJAMG11C
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_FP17_ZR_TW,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

                'IF FP-17(ZR/TW) THEN 必是(PC9)
                MMIC_FLEX_N_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX100_PC9,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_FP17_ZR_TW_FrontRearDoor,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF FP-17(TW)_正背門 THEN TJAMG12A
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_FP17_ZR_TW_FrontRearDoor,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_FP17_ZR_HK,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF FP-17(ZR/HK) THEN TJDMG94F                                       
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_FP17_ZR_HK,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_GLVF_HK_Hallbus,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF GLVF-MOD(HK)_HALLBUS通信 THEN TJDM201C
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_GLVF_HK_Hallbus,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_GLVF_HK_SelcomDoor,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF GLVF-MOD(HK)_SELCOM_DOOR THEN TJDM203A
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_GLVF_HK_SelcomDoor,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_GLVF_E_SP,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF GLVF-E-C_LVF THEN TJEMC63H
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_GLVF_E_SP,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_REXIAa_TW,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF REXIAa(TW) THEN TJAMA51A
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_REXIAa_TW,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_TP09_TW,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF TP-09(TW) THEN TJAME61A
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_TP09_TW,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_XIOR_TW,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF XIOR(TW) THEN TJAMF21A
                MMIC_MR_Base_TextBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmic_XIOR_TW,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_GLVF_HK_Millnet,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF GLVF-MOD(HK)_MILLNET通信 THEN TJDM202A
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_GLVF_HK_Millnet,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_MachineType_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.mmicN_GLVF_D_SP,
                                              get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'IF GLVF-Da_HLV THEN TJEMD63B
                MMIC_MR_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.mmic_GLVF_D_SP,
                                                  get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
            End If
        Else
            MMIC_Panel.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' [MMIC > FLEX N XX ComboBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FLEX_N_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MMIC_FLEX_N_ComboBox.TextChanged
        If MMIC_FLEX_N_ComboBox.Text <> "" Then
            Select Case localSelect
                Case Taiwan
                    With MMIC_SV_ECarObj_ComboBox
                        If .Items.Count <> 0 Then
                            .SelectedIndex = MMIC_SV_ECarObj_ComboBox.Items.Count - 1
                        End If
                    End With
            End Select

            MMIC_SV_GroupBox.Enabled = True
            MMIC_SV_E_GroupBox.Enabled = True

            If MMIC_FLEX_N_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX100_PC8,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name) Then
                'if NX100-PC8 then F7702202
                MMIC_SV_EBase_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX100_PC8_FileName,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                'if NX100-PC8 THEN WITH CP43X
                MMIC_MR_CP43x_ComboBox.Text = get_nameManager.TB_WITH


                Select Case MMIC_MachineType_ComboBox.Text
                    Case get_nameManager.read_DbmsData(get_nameManager.mmicN_FP17_ZR_TW,
                                                       get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                                       get_nameManager.SQLite_connectionPath_Tool,
                                                       get_nameManager.SQLite_ToolDBMS_Name)
                        '[提醒]if NX100-PC8 且 台灣FP-17時，目前通常使用PC9 -----------------------------------------
                        MsgBox("目前台灣FP-17通常使用PC9", MsgBoxStyle.Exclamation, "提醒")
                        MMIC_MR_Base_TextBox.Text =
                            get_nameManager.read_DbmsData(get_nameManager.mmic_FP17_ZR_TW_PC8,
                                                          get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                          get_nameManager.SQLite_connectionPath_Tool,
                                                          get_nameManager.SQLite_ToolDBMS_Name)
                        '-----------------------------------------[提醒]if NX100-PC8 且 台灣FP-17時，目前通常使用PC9 
                    Case get_nameManager.read_DbmsData(get_nameManager.mmicN_FP17_ZR_TW_FrontRearDoor,
                                                       get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                                       get_nameManager.SQLite_connectionPath_Tool,
                                                       get_nameManager.SQLite_ToolDBMS_Name)
                        '[提醒]if NX100-PC8 且 台灣FP-17時，目前通常使用PC9 -----------------------------------------
                        MsgBox("目前台灣FP-17通常使用PC9", MsgBoxStyle.Exclamation, "提醒")
                        '-----------------------------------------[提醒]if NX100-PC8 且 台灣FP-17時，目前通常使用PC9 
                    Case get_nameManager.read_DbmsData(get_nameManager.mmicN_IDU_ZT_TW,
                                                       get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                                       get_nameManager.SQLite_connectionPath_Tool,
                                                       get_nameManager.SQLite_ToolDBMS_Name)
                        'if NX100-PC8 且 台灣ZEXIA-T時 ------------------------------------------------
                        MMIC_MR_Base_TextBox.Text =
                            get_nameManager.read_DbmsData(get_nameManager.mmic_IDU_ZT_TW_PC8,
                                                          get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                          get_nameManager.SQLite_connectionPath_Tool,
                                                          get_nameManager.SQLite_ToolDBMS_Name)
                        '------------------------------------------------ if NX100-PC8 且 台灣ZEXIA-T時 
                End Select


            ElseIf MMIC_FLEX_N_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX100_PC9,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name) Then
                'If NX100-PC9 then F7702302
                MMIC_SV_EBase_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX100_PC9_FileName,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
                'if NX100-PC9 THEN WITHOUT CP43X
                MMIC_MR_CP43x_ComboBox.Text = get_nameManager.TB_WITHOUT


                Select Case MMIC_MachineType_ComboBox.Text
                    Case get_nameManager.read_DbmsData(get_nameManager.mmicN_FP17_ZR_TW,
                                                       get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                                       get_nameManager.SQLite_connectionPath_Tool,
                                                       get_nameManager.SQLite_ToolDBMS_Name)
                        'if NX100-PC9 且 台灣FP-17時 -----------------------------------------------------
                        MMIC_MR_Base_TextBox.Text =
                        get_nameManager.read_DbmsData(get_nameManager.mmic_FP17_ZR_TW,
                                                      get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                      get_nameManager.SQLite_connectionPath_Tool,
                                                      get_nameManager.SQLite_ToolDBMS_Name)
                        '----------------------------------------------------- if NX100-PC9 且 台灣FP-17時 
                    Case get_nameManager.read_DbmsData(get_nameManager.mmicN_FP17_ZR_TW_FrontRearDoor,
                                                       get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                                       get_nameManager.SQLite_connectionPath_Tool,
                                                       get_nameManager.SQLite_ToolDBMS_Name)
                        'if NX100-PC8 且 台灣FP-17 正背門時 -----------------------------------------
                        MMIC_MR_Base_TextBox.Text =
                            get_nameManager.read_DbmsData(get_nameManager.mmic_FP17_ZR_TW_FrontRearDoor,
                                                          get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                          get_nameManager.SQLite_connectionPath_Tool,
                                                          get_nameManager.SQLite_ToolDBMS_Name)
                        '-----------------------------------------If NX100 - PC8 Then 且 台灣FP - 17 正背門時 
                    Case get_nameManager.read_DbmsData(get_nameManager.mmicN_IDU_ZT_TW,
                                                       get_nameManager.SQLite_tableName_MMIC_ProgramTypeName,
                                                       get_nameManager.SQLite_connectionPath_Tool,
                                                       get_nameManager.SQLite_ToolDBMS_Name)
                        'if NX100-PC8 且 台灣ZEXIA-T時 ------------------------------------------------
                        MMIC_MR_Base_TextBox.Text =
                            get_nameManager.read_DbmsData(get_nameManager.mmic_IDU_ZT_TW,
                                                          get_nameManager.SQLite_tableName_MMIC_ProgramType,
                                                          get_nameManager.SQLite_connectionPath_Tool,
                                                          get_nameManager.SQLite_ToolDBMS_Name)
                        '------------------------------------------------ if NX100-PC8 且 台灣ZEXIA-T時 
                End Select


            ElseIf MMIC_FLEX_N_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX200,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name) Then
                'If NX200 then F9702202
                MMIC_SV_EBase_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX200_FileName,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_FLEX_N_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX300,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name) Then
                'If NX300 then F9702202
                MMIC_SV_EBase_ComboBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.FLEX_NX300_FileName,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
            End If
        ElseIf MMIC_FLEX_N_ComboBox.Text = "" Then
            MMIC_SV_GroupBox.Enabled = False
            MMIC_SV_E_GroupBox.Enabled = False
        End If
    End Sub


    ''' <summary>
    ''' [MMIC > MR > NumericUpDown]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MMIC_MR_CarObj_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles MMIC_MR_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_MMICInfo()
        AddSub_Object_Sub(MMIC_MR_NumericUpDown,
                          MMIC_MR_Panel,
                          mmicType1_CarNo_TextBox,
                          mmicType1_ObjName_TextBox,
                          mmicType1_ObjNameBase_TextBox,
                          dyCtrlName.JobMaker_MMIC_MrBase_InfoName_Array.Count,
                          dyCtrlName.JobMaker_MMIC_MrBase_InfoName_Array,
                          MMIC_MR_Base_TextBox.Text)
    End Sub
    ''' <summary>
    ''' [MMIC > MR EEPROM > NumericUpDown]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MMIC_MR_ECarNo_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles MMIC_MR_E_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_MMICInfo()
        AddSub_Object_Sub(MMIC_MR_E_NumericUpDown,
                          MMIC_MR_E_Panel,
                          mmic_CarNo_TextBox,
                          mmic_ObjName_TextBox,
                          dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array.Count,
                          dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array,
                          MMIC_MR_ECarObj_ComboBox.Text)
    End Sub
    ''' <summary>
    ''' [MMIC > SV > NumericUpDown]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MMIC_SV_CarObj_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles MMIC_SV_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_MMICInfo()
        AddSub_Object_Sub(MMIC_SV_NumericUpDown,
                          MMIC_SV_Panel,
                          mmicType1_CarNo_TextBox,
                          mmicType1_ObjName_TextBox,
                          mmicType1_ObjNameBase_TextBox,
                          dyCtrlName.JobMaker_MMIC_SvBase_InfoName_Array.Count,
                          dyCtrlName.JobMaker_MMIC_SvBase_InfoName_Array,
                          MMIC_SV_Base_TextBox.Text)
    End Sub
    ''' <summary>
    ''' [MMIC > SV EEPROM > NumericUpDown]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JM_SV_EEPROM_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles MMIC_SV_E_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_MMICInfo()
        AddSub_Object_Sub(MMIC_SV_E_NumericUpDown,
                          MMIC_SV_E_Panel,
                          mmic_CarNo_TextBox,
                          mmic_ObjName_TextBox,
                          dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array.Count,
                          dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array,
                          MMIC_SV_ECarObj_ComboBox.Text)
    End Sub
    ''' <summary>
    ''' [MMIC > VD10 > NumericUpDown]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MMIC_VD10_ObjCar_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles MMIC_VD10_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        dyCtrlName.JobMaker_MMICInfo()
        AddSub_Object_Sub(MMIC_VD10_NumericUpDown,
                          MMIC_VD10_Panel,
                          mmic_CarNo_TextBox,
                          mmic_ObjName_TextBox,
                          dyCtrlName.JobMaker_MMIC_VD10Base_InfoName_Array.Count,
                          dyCtrlName.JobMaker_MMIC_VD10Base_InfoName_Array,
                          MMIC_VD10_Base_TextBox.Text)
    End Sub

    ''' <summary>
    ''' [MMIC > NumericUpDown > 自動填入或刪除Panel中的TextBox (2個控制項)]
    ''' </summary>
    ''' <param name="mNumericUpDown"></param>
    ''' <param name="mpanel"></param>
    ''' <param name="tb_lift">自動生成的Car No.</param>
    ''' <param name="tb_objName">自動生成的Object Name.</param>
    ''' <param name="dyCtrl_ArrayCount">自動生成控制項的總數</param>
    ''' <param name="dyCrtl_Array">自動生成控制項的名稱Name</param>
    ''' <param name="mBaseName">自動生成控制項Object Name的文字Text</param>
    Overloads Sub AddSub_Object_Sub(mNumericUpDown As NumericUpDown, mpanel As Panel,
                                    tb_lift As Control, tb_objName As Control,
                                    dyCtrl_ArrayCount As Integer, dyCrtl_Array As Array,
                                    mBaseName As String)

        Dim ObjNum As Integer
        '嘗試得到mNumericUpDown的數量
        Try
            ObjNum = mNumericUpDown.Value
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.AddSub_Object_Sub")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
        End Try

        'Dim dyCtrlName As DynamicControlName = New DynamicControlName
        'dyCtrlName.JobMaker_MMICInfo()

        Dim LiftNum_Panel_count, i_start As Integer
        LiftNum_Panel_count = mpanel.Controls.Count
        If LiftNum_Panel_count = 0 Then
            i_start = 1
        Else
            i_start = LiftNum_Panel_count / dyCtrl_ArrayCount + 1
        End If


        Dim TextBoxWidth, TextBox_XPosition, TextBox_YPosition As Integer()

        TextBoxWidth = {tb_lift.Width, tb_objName.Width}
        TextBox_XPosition = {tb_lift.Left, tb_objName.Left}
        TextBox_YPosition = {tb_lift.Top, tb_objName.Top}

        If i_start > ObjNum Then
            '如果 Panel中數量比 mNumericUpDown數量多
            '就 刪除 ----------------------------------
            For decrease_j As Integer = 1 To dyCtrl_ArrayCount
                For Each CtrlName As Control In mpanel.Controls
                    If CtrlName.Name = $"{dyCrtl_Array(decrease_j - 1)}_{i_start - 1}" Then
                        mpanel.Controls.Remove(CtrlName)
                    End If
                Next
            Next
            '---------------------------------- 刪除 
        Else
            '如果 Panel中數量比 mNumericUpDown數量少
            '就 增加 ----------------------------------
            Dim ConNum_tb As TextBox

            For Lift_i As Integer = i_start To ObjNum
                For Obj_j As Integer = 1 To dyCtrl_ArrayCount
                    ConNum_tb = New TextBox()

                    With ConNum_tb
                        If Obj_j = 1 Then
                            .Text = "L#" & ObjNum
                        Else
                            .Text = mBaseName
                        End If
                        .Width = TextBoxWidth(Obj_j - 1)
                        .Left = TextBox_XPosition(Obj_j - 1)
                        .Top = TextBox_YPosition(Obj_j - 1) + (ObjNum - 1) * 30
                        .Visible = True
                        .TextAlign = HorizontalAlignment.Center '文字至中
                        .Font =
                            New System.Drawing.Font("微軟正黑體",
                                                    9.0!,
                                                    System.Drawing.FontStyle.Regular,
                                                    System.Drawing.GraphicsUnit.Point,
                                                    CType(136, Byte))
                        .Name = ($"{dyCrtl_Array(Obj_j - 1)}_{Lift_i}")
                        .TabIndex = Lift_i

                        mpanel.Controls.Add(ConNum_tb)
                    End With
                Next Obj_j
            Next Lift_i
            '---------------------------------- 增加 
        End If
    End Sub

    ''' <summary>
    ''' [MMIC > NumericUpDown > 自動填入或刪除Panel中的TextBox (3個控制項)]
    ''' </summary>
    ''' <param name="mNumericUpDown"></param>
    ''' <param name="mpanel"></param>
    ''' <param name="tb_lift"></param>
    ''' <param name="tb_objName"></param>
    ''' <param name="tb_base"></param>
    ''' <param name="dyCtrl_ArrayCount"></param>
    ''' <param name="dyCrtl_Array"></param>
    ''' <param name="mBaseName"></param>
    Overloads Sub AddSub_Object_Sub(mNumericUpDown As NumericUpDown, mpanel As Panel,
                                    tb_lift As Control, tb_objName As Control, tb_base As Control,
                                    dyCtrl_ArrayCount As Integer, dyCrtl_Array As Array,
                                    mBaseName As String)

        Dim ObjNum As Integer
        '嘗試得到mNumericUpDown的數量
        Try
            ObjNum = mNumericUpDown.Value
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.AddSub_Object_Sub")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
        End Try

        Dim LiftNum_Panel_count, i_start As Integer
        LiftNum_Panel_count = mpanel.Controls.Count
        If LiftNum_Panel_count = 0 Then
            i_start = 1
        Else
            i_start = LiftNum_Panel_count / dyCtrl_ArrayCount + 1
        End If


        Dim TextBoxWidth, TextBox_XPosition, TextBox_YPosition As Integer()

        TextBoxWidth = {tb_lift.Width, tb_objName.Width, tb_base.Width}
        TextBox_XPosition = {tb_lift.Left, tb_objName.Left, tb_base.Left}
        TextBox_YPosition = {tb_lift.Top, tb_objName.Top, tb_base.Top}

        If i_start > ObjNum Then
            '如果 Panel中數量比 mNumericUpDown數量多
            '就 刪除 ----------------------------------
            For decrease_j As Integer = 1 To dyCtrl_ArrayCount
                For Each CtrlName As Control In mpanel.Controls
                    If CtrlName.Name = $"{dyCrtl_Array(decrease_j - 1)}_{i_start - 1}" Then
                        mpanel.Controls.Remove(CtrlName)
                    End If
                Next
            Next
            '---------------------------------- 刪除 
        Else
            '如果 Panel中數量比 mNumericUpDown數量少
            '就 增加 ----------------------------------
            Dim ConNum_tb As TextBox

            For Lift_i As Integer = i_start To ObjNum
                For Obj_j As Integer = 1 To dyCtrl_ArrayCount
                    ConNum_tb = New TextBox()

                    With ConNum_tb
                        If Obj_j = 1 Then
                            .Text = "L#" & ObjNum
                        Else
                            .Text = mBaseName
                        End If
                        .Width = TextBoxWidth(Obj_j - 1)
                        .Left = TextBox_XPosition(Obj_j - 1)
                        .Top = TextBox_YPosition(Obj_j - 1) + (ObjNum - 1) * 30
                        .Visible = True
                        .TextAlign = HorizontalAlignment.Center '文字至中
                        .Font =
                            New System.Drawing.Font("微軟正黑體",
                                                    9.0!,
                                                    System.Drawing.FontStyle.Regular,
                                                    System.Drawing.GraphicsUnit.Point,
                                                    CType(136, Byte))
                        .Name = ($"{dyCrtl_Array(Obj_j - 1)}_{Lift_i}")
                        .TabIndex = Lift_i

                        mpanel.Controls.Add(ConNum_tb)
                    End With
                Next Obj_j
            Next Lift_i
            '---------------------------------- 增加 
        End If
    End Sub


    Overloads Sub AddSub_Object_Sub(mNumericUpDown As NumericUpDown, mpanel As Panel,
                                    ctrl() As Control,
                                    dyCtrl_ArrayCount As Integer, dyCrtl_Array As Array,
                                    mSql_tableName_Array As Array,
                                    mSpecName_Array As Array)

        Dim ObjNum As Integer
        '嘗試得到mNumericUpDown的數量
        Try
            ObjNum = mNumericUpDown.Value
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.AddSub_Object_Sub")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
        End Try

        Dim LiftNum_Panel_count, i_start As Integer
        LiftNum_Panel_count = mpanel.Controls.Count
        If LiftNum_Panel_count = 0 Then
            i_start = 1
        Else
            i_start = LiftNum_Panel_count / dyCtrl_ArrayCount + 1
        End If

        Dim ctrl_count As Integer
        ctrl_count = ctrl.Count

        If i_start > ObjNum Then
            '如果 Panel中數量比 mNumericUpDown數量多
            '就 刪除 ----------------------------------
            For decrease_j As Integer = 1 To dyCtrl_ArrayCount
                For Each CtrlName As Control In mpanel.Controls
                    If CtrlName.Name = $"{dyCrtl_Array(decrease_j - 1)}_{i_start - 1}" Then
                        mpanel.Controls.Remove(CtrlName)
                    End If
                Next
            Next
            '---------------------------------- 刪除 
        Else
            '如果 Panel中數量比 mNumericUpDown數量少
            Dim ConNum
            '就 增加 ----------------------------------


            For Lift_i As Integer = i_start To ObjNum
                For Obj_j As Integer = 1 To dyCtrl_ArrayCount
                    'ConNum_cb = New ComboBox()
                    If TypeOf ctrl(0) Is ComboBox Then
                        ConNum = New ComboBox()
                    ElseIf TypeOf ctrl(0) Is TextBox Then
                        ConNum = New TextBox()
                    Else
                    End If
                    With ConNum
                        .Width = ctrl(Obj_j - 1).Width
                        .Left = ctrl(Obj_j - 1).Left
                        .Top = ctrl(Obj_j - 1).Top + (ObjNum - 1) * 30
                        .Visible = True
                        .Font =
                            New System.Drawing.Font("微軟正黑體",
                                                    9.0!,
                                                    System.Drawing.FontStyle.Regular,
                                                    System.Drawing.GraphicsUnit.Point,
                                                    CType(136, Byte))
                        If TypeOf (ConNum) Is TextBox Then
                            '.text = controlerText
                            .Text = get_nameManager.read_DbmsData(mSpecName_Array(Obj_j - 1),
                                                                  mSql_tableName_Array(Obj_j - 1),
                                                                  ConNum,
                                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                                  get_nameManager.SQLite_ToolDBMS_Name)
                            .TextAlign = HorizontalAlignment.Center '文字至中
                        ElseIf TypeOf (ConNum) Is ComboBox Then
                            get_nameManager.read_DbmsData(mSpecName_Array(Obj_j - 1),
                                                          mSql_tableName_Array(Obj_j - 1),
                                                          ConNum,
                                                          get_nameManager.SQLite_connectionPath_Tool,
                                                          get_nameManager.SQLite_ToolDBMS_Name)
                            .SelectedIndex = 0
                        End If
                        .Name = ($"{dyCrtl_Array(Obj_j - 1)}_{Lift_i}")
                        .TabIndex = Lift_i
                        mpanel.Controls.Add(ConNum)
                    End With
                Next Obj_j
            Next Lift_i
            '---------------------------------- 增加 
        End If
    End Sub
    ''' <summary>
    ''' [MMIC > VD10 > Rom Device Combobox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MMIC_VD10_ROM_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MMIC_VD10_ROM_ComboBox.SelectedIndexChanged
        If MMIC_VD10_ROM_ComboBox.Text = "4Mb" Then
            MMIC_VD10_Quantity_ComboBox.Text = "2"
        ElseIf MMIC_VD10_ROM_ComboBox.Text = "8Mb" Then
            MMIC_VD10_Quantity_ComboBox.Text = "1"
        End If
    End Sub
    ''' <summary>
    ''' [MMIC > VD10 > TYPE Combobox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JM_VD10_Type_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MMIC_VD10_Type_ComboBox.SelectedIndexChanged
        If MMIC_VD10_Type_ComboBox.Text <> "" Then
            If MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_TW_STD_LOWER,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 台灣標準低樓層 Then P3F00L81
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_TW_STD_LOWER,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_TW_STD_HIGHER,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 台灣標準高樓層 Then P3F00M81
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_TW_STD_HIGHER,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_SP_STD_STOREY,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 新加坡Storey發音 Then P3F00H62
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_SP_STD_STOREY,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_SP_STD_FLOOR,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 新加坡Floor發音 Then P3F00J62
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_SP_STD_FLOOR,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_TW_NSTD_Lobby_R,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 台灣非標準 Lobby_R Then 
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_TW_NSTD_Lobby_R,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_TW_NSTD_1M_2M,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 台灣非標準 1M 2M Then 
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_TW_NSTD_1M_2M,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_TW_NSTD_Taiwanese,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 台灣非標準 國+台 Then 
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_TW_NSTD_Taiwanese,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_TW_NSTD_Taiwanese_B,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 台灣非標準 國+台 有B樓 Then 
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_TW_NSTD_Taiwanese_B,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_HK_NSTD_B_G,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 香港非標準 有B G樓 Then 
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_HK_NSTD_B_G,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_SP_NSTD_M,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 新加坡非標準 有M樓 Then P3F00J62
                MMIC_VD10_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.VD10_SP_NSTD_M,
                                                  get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_VD10_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.VD10_Other_Path,
                                              get_nameManager.SQLite_tableName_VD10_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 其他 Then 打開資料夾
                Dim type_result As MsgBoxResult = MsgBox("是否打開其他VD10工直仕樣?", vbYesNo, "提醒(開啟Excel)")
                If type_result = MsgBoxResult.Yes Then
                    msExcel_app = New Excel.Application
                    msExcel_workbook =
                        msExcel_app.Workbooks.Open(get_nameManager.read_DbmsData(get_nameManager.VD10_Other_Path,
                                                                                 get_nameManager.SQLite_tableName_VD10_ProgramType,
                                                                                 get_nameManager.SQLite_connectionPath_Tool,
                                                                                 get_nameManager.SQLite_ToolDBMS_Name))
                    msExcel_app.Visible = True
                End If
            End If

            'ChkList_5_std_Content_TextBox.Text = MMIC_VD10_Base_TextBox.Text
        End If
    End Sub

    ''' <summary>
    ''' [MMIC > SV > TYPE Combobox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JM_SV_Type_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MMIC_SV_Type_ComboBox.TextChanged
        If MMIC_SV_Type_ComboBox.Text <> "" Then
            If MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_N100_PC8,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If N100 PC8 Then F91029ZA
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_N100_PC8,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_N100_PC9,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If N100 PC9 Then F91029ZA
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_N100_PC9,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_OverN200,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If Over N200 Then F91029ZA
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_OverN200,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_ELVIC_TW,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 台灣Elvic Then TJAGB91A
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_ELVIC_TW,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_EOP,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If EOP Then TJZGB9BA
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_EOP,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_EvaucationOpe_SP,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 新加坡救出運轉 Then TJZGB9DA
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_EvaucationOpe_SP,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_GsoTo1Car,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 群管理切1Car Then TJZGB9AA
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_GsoTo1Car,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_IndependentPowerOpe,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If 專用電源運轉 Then TJZGB9CA
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_IndepPowerOpe,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)

            ElseIf MMIC_SV_Type_ComboBox.Text =
                get_nameManager.read_DbmsData(get_nameManager.gspName_Double2Car,
                                              get_nameManager.SQLite_tableName_GSP_ProgramTypeName,
                                              get_nameManager.SQLite_connectionPath_Tool,
                                              get_nameManager.SQLite_ToolDBMS_Name) Then
                'If Double2Car Then TJDGBAEA
                MMIC_SV_Base_TextBox.Text =
                    get_nameManager.read_DbmsData(get_nameManager.gsp_Double2Car,
                                                  get_nameManager.SQLite_tableName_GSP_ProgramType,
                                                  get_nameManager.SQLite_connectionPath_Tool,
                                                  get_nameManager.SQLite_ToolDBMS_Name)
            End If
        End If
    End Sub

    '------------------------------------------------------------------------------------------------------------------------- MMIC 

    'EepData --------------------------------------------------------------------------------------------------------------------
    Private Sub Use_EepData_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_EepData_CheckBox.CheckedChanged
        If Use_EepData_CheckBox.Checked Then
            use_EepData_chkbox_clickTimes += 1
            EepData_TabControl.Enabled = True
            If use_EepData_chkbox_clickTimes = 1 Then

                For Each tabPage As Control In EepData_TabControl.Controls
                    For Each grp As Control In tabPage.Controls
                        For Each Ctrl As Control In grp.Controls
                            If TypeOf (Ctrl) Is TextBox Then
                                With Ctrl
                                    AddHandler .MouseEnter, AddressOf TextBox_ResizeHeight_MouseEnter
                                    AddHandler .MouseLeave, AddressOf TextBox_ResizeHeight_MouseLeave
                                End With
                            End If
                        Next
                    Next
                Next
            End If
        Else
            EepData_TabControl.Enabled = False
        End If
    End Sub
    Private Sub TextBox_ResizeHeight_MouseEnter(sender As Object, e As EventArgs)
        Dim tb = DirectCast(sender, TextBox)
        With tb
            .Height = 60
            .BringToFront()
        End With
    End Sub
    Private Sub TextBox_ResizeHeight_MouseLeave(sender As Object, e As EventArgs)
        Dim tb = DirectCast(sender, TextBox)
        tb.Height = 20
    End Sub
    '--------------------------------------------------------------------------------------------------------------------EepData 

    'G值------------------------------------------------------------------------------------------------------------------------- 
    Private Sub GWeb_Button_Click(sender As Object, e As EventArgs) Handles GWeb_Button.Click
        'Dim wb = New WebBrowser
        'wb.Navigate("http://10.213.2.42/web/WebForm1")
        Shell("explorer http://10.213.2.42/web/WebForm1")
    End Sub
    ''' <summary>
    ''' [G值 > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Use_G_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_FinalCheck_CheckBox.CheckedChanged
        If Use_FinalCheck_CheckBox.Checked Then
            FinalCheck_GroupBox.Enabled = True
        Else
            FinalCheck_GroupBox.Enabled = False
        End If
    End Sub
    '------------------------------------------------------------------------------------------------------------------------- G值


    '其他事件 -----------------------------------------------------------------------------------------------------------------------
    Private Sub JobMaker_Timer_Tick(sender As Object, e As EventArgs) Handles JobMaker_Timer.Tick
        If NumericUpDown1.Value > 0 Then
            JobMaker_Timer.Interval = NumericUpDown1.Value '事件發生間隔透過數值調整設定
            ReminderMarquee_Label.Left = ReminderMarquee_Label.Left - 1
            ReminderMarquee2_Label.Left = ReminderMarquee2_Label.Left - 1
            If ReminderMarquee_Label.Left < 0 - ReminderMarquee_Label.Width / 5 Then
                ReminderMarquee_Label.Left = ReminderMarquee_Label.Width
            End If
            If ReminderMarquee2_Label.Left < 0 - ReminderMarquee2_Label.Width / 5 Then
                ReminderMarquee2_Label.Left = ReminderMarquee2_Label.Width
            End If
        End If
    End Sub

    Private Sub SaveFile_Button_Click(sender As Object, e As EventArgs) Handles SaveFile_Button.Click
        saveFile_toSQLite(False)
    End Sub

    Private Sub JobMaker_Close_Button_Click(sender As Object, e As EventArgs) Handles JobMaker_Close_Button.Click
        saveFile_toSQLite(True)
    End Sub

    ''' <summary>
    ''' [存檔 > SQLite]
    ''' </summary>
    ''' <param name="isClosed">True關閉 / False不關閉 Form</param>
    Private Sub saveFile_toSQLite(isClosed As Boolean)
        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        Dim checkFlie_IfExists As Boolean

        If SQLite_FixBug_Button_ClickCount = 0 Then
            Dim reminder As MsgBoxResult = MsgBox($"仕樣 > Basic All > Page2 {vbCrLf} 載入按鈕未使用你確定要繼續存檔? {vbCrLf} (是:繼續/否:離開)",
                                                  vbYesNo + vbExclamation, "提醒")
            If reminder = MsgBoxResult.Yes Then
            Else
                Exit Sub
            End If
        End If

        Dim Stored_result As MsgBoxResult = MsgBox($"是否儲存你輸入的工番資料? {vbCr} (是:繼續/否:離開)", vbYesNo, "提醒")
        Dim Stored_Input

        Try
            If Stored_result = MsgBoxResult.Yes Then
                Do
                    Dim jobNo_from_user As String
                    If Basic_JobNoNew_TextBox.Text <> "" Then
                        jobNo_from_user = Basic_JobNoNew_TextBox.Text
                    Else
                        jobNo_from_user = Replace(Load_SQLite_JobSearch_ComboBox.Text, ".sqlite", "")
                    End If
                    Stored_Input = InputBox("輸入Job Name(範例:TW-9453-55)", "儲存新檔", jobNo_from_user)

                    If Stored_Input = "" Then
                        MsgBox("未輸入JobName，請重來",, "SQLite存檔訊息")
                    ElseIf Len(Stored_Input) = 0 Then
                        MsgBox("取消",, "SQLite存檔訊息")
                    Else
                        Resize_JMForm(JMForm_size.re_size)
                        '尋找資料夾是否有重複檔案
                        checkFlie_IfExists = File.Exists(spec_stored.SQLite_connectionPath_Job & $"{Stored_Input}.sqlite")

                        If checkFlie_IfExists = True Then
                            Dim checkFile_IfExists_result As MsgBoxResult = MsgBox($"{Stored_Input}已存在，是否覆蓋檔案?",
                                                                                   vbYesNo + vbExclamation, "提醒")
                            If checkFile_IfExists_result = MsgBoxResult.Yes Then
                                My.Computer.FileSystem.DeleteFile(spec_stored.SQLite_connectionPath_Job & $"{Stored_Input}.sqlite")
                                My.Computer.FileSystem.CopyFile(spec_stored.SQLite_connectionPath_Tool & spec_stored.SQLite_StdJobDataDBMS_Name,
                                                            spec_stored.SQLite_connectionPath_Job & $"{Stored_Input}.sqlite")
                                'spec_stored.SQLiteUpdate_Stored($"{Stored_Input}.sqlite", checkFlie_IfExists)
                                spec_stored.SQLiteUpdate_Stored($"{Stored_Input}.sqlite")
                                MsgBox($"{Stored_Input}已覆蓋",, "SQLite存檔訊息")

                                If isClosed Then
                                    Me.Close()
                                End If
                            Else
                                MsgBox($"{Stored_Input}未覆蓋",, "SQLite存檔訊息")
                            End If
                        Else
                            My.Computer.FileSystem.CopyFile(spec_stored.SQLite_connectionPath_Tool & spec_stored.SQLite_StdJobDataDBMS_Name,
                                                            spec_stored.SQLite_connectionPath_Job & $"{Stored_Input}.sqlite")
                            'spec_stored.SQLiteUpdate_Stored($"{Stored_Input}.sqlite", checkFlie_IfExists)
                            spec_stored.SQLiteUpdate_Stored($"{Stored_Input}.sqlite")
                            MsgBox($"JobName:{Stored_Input}已存檔",, "SQLite存檔訊息")

                            If isClosed Then
                                Me.Close()
                            End If
                        End If
                    End If
                Loop Until Stored_Input <> "" Or Len(Stored_Input) = 0
            ElseIf Stored_result = MsgBoxResult.No Then
                If isClosed Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.JobMaker_Close_Button_Click")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox($"關閉時儲存SQLite錯誤{vbCrLf}{ex.Message}",, "SQLite存檔訊息")
        End Try
    End Sub

    ''' <summary>
    ''' [JobMaker > 關閉Debug視窗]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ResultCheck_Button_Click(sender As Object, e As EventArgs) Handles ResultClose_Button.Click
        With ResultOutput_TextBox
            .Clear()
            '.Visible = False
        End With
        With ResultFailOutput_TextBox
            .Clear()
            '.Visible = False
        End With
        With ResultClose_Button
            '.Visible = False
        End With

        Resize_JMForm(JMForm_size.ini_size)
    End Sub

    ''' <summary>
    ''' 測試用
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles HIN_TestButton.Click
        '---------------------------------------------- 求最高樓層 


        If HallIndicator_FlowLayoutPanel.Controls.Count <> 0 Then
            Resize_JMForm(JMForm_size.re_size)

            Dim HinLiftDiff_bool, HinFLDiff_bool As Boolean
            Dim lift_i, stop_i As Integer

            '求最高樓層 ----------------------------------------------
            Dim stopFL_MAX, stopFL_MIN As Integer 'HIN中最高樓層
            For lift_i = 1 To LiftNum
                For stop_i = 1 To arr_liftStopFL(lift_i - 1)
                    If stop_i > stopFL_MAX Then
                        stopFL_MAX = stop_i
                    Else
                        stopFL_MIN = stop_i
                    End If
                Next
            Next

            Dim arr_liftStopFL_userContent(LiftNum - 1, stopFL_MAX - 1) As String
            Dim dyCtrlName As DynamicControlName = New DynamicControlName


            'ResultOutput_TextBox.Text += $"最高樓層數:{stopFL_MAX} 目前陣列數 {arr_liftStopFL_userContent.Length} {vbCrLf}"
            '儲存使用者值得內容 ----------------------------------------------------------------
            For Each flp In HallIndicator_FlowLayoutPanel.Controls.OfType(Of FlowLayoutPanel)
                For Each cb In flp.Controls.OfType(Of CheckBox)
                    For lift_i = 1 To LiftNum
                        For stop_i = 1 To arr_liftStopFL(lift_i - 1)
                            If cb.Name = $"{stop_i}{dyCtrlName.JobMaker_HIN_FL_ChkB}_{lift_i}" Then
                                'ResultOutput_TextBox.Text += $"{cb.Name}:{cb.CheckState}{vbCrLf}"
                                For Each cmbbox In flp.Controls.OfType(Of ComboBox)
                                    If cmbbox.Name = $"{stop_i}{dyCtrlName.JobMaker_HIN_FL_CmbB}_{lift_i}" Then
                                        If cb.Checked Then
                                            arr_liftStopFL_userContent(lift_i - 1, stop_i - 1) = cmbbox.Text
                                        Else
                                            arr_liftStopFL_userContent(lift_i - 1, stop_i - 1) = "Nothing"
                                        End If
                                        'ResultOutput_TextBox.Text += ($"{lift_i - 1}:{stop_i - 1}->{arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}")
                                    End If
                                Next

                            End If
                        Next
                    Next
                Next
            Next
            '---------------------------------------------------------------- 儲存使用者值得內容 


            '計算 比較個號機有無相同
            Dim HinLiftSame_cnt, HinLiftDiff_cnt As Integer

            '顯示 [...] 字樣
            Dim HinPoint_bool As Boolean

            Dim topFL_End_bool As Boolean
            For stop_i = 1 To stopFL_MAX 'arr_liftStopFL(LiftNum - 1)
                '每次換樓層時清空arr_liftStopFl_EachContent內資料 ----
                HinLiftDiff_bool = False '號機不同
                HinFLDiff_bool = False '樓層不同
                For i = 1 To arr_liftStopFl_StdContent.Count
                    For lift_i = 1 To LiftNum
                        If arr_liftStopFl_EachContent(i - 1, lift_i) <> Nothing Then '共三列，第一列為標準值
                            arr_liftStopFl_EachContent(i - 1, lift_i) = Nothing '將值都清空做後續比對
                        End If
                    Next
                Next
                '---- 每次換樓層時清空arr_liftStopFl_EachContent內資料 


                '每次換樓層時判斷 #1~#N 號機該樓層HIN是否都相同? ---------------------------------
                For lift_i = 1 To LiftNum
                    If lift_i < LiftNum Then
                        'Console.WriteLine($"{stop_i}FL:{arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)},{arr_liftStopFL_userContent(lift_i, stop_i - 1)}")
                        If arr_liftStopFL_userContent(lift_i - 1, stop_i - 1) =
                           arr_liftStopFL_userContent(lift_i, stop_i - 1) Then
                            '號機之間值相同 -------------------
                            HinLiftDiff_bool = False
                            'Hin_LiftSame_bool = True
                            '------------------- 號機之間值相同

                            '上下樓層之間不同 ------------
                            For lift_ii = 1 To LiftNum
                                If stop_i + 1 <= stopFL_MAX Then
                                    If arr_liftStopFL_userContent(lift_ii - 1, stop_i) <>
                                       arr_liftStopFL_userContent(lift_ii - 1, stop_i - 1) Then
                                        HinFLDiff_bool = True
                                        'Hin_FLDiff_bool = True
                                        'HinPoint_bool = False
                                    End If
                                End If
                            Next
                            '------------ 上下樓層之間不同 
                        Else
                            '號機之間值不相同 -----------------
                            HinLiftDiff_bool = True
                            'Hin_LiftDiff_bool = True
                            HinLiftDiff_cnt = HinLiftSame_cnt + 1
                            '----------------- 號機之間值不相同
                            Exit For
                        End If
                    End If
                Next
                lift_i = 0




                If HinLiftDiff_bool Then '表示同樓層的號機之間值都不相同
                    For lift_i = 1 To LiftNum
                        '當使用者輸入的HIN為空時 ----------------------------------------------
                        If arr_liftStopFL_userContent(lift_i - 1, stop_i - 1) = "" Then
                            'ResultOutput_TextBox.Text += $"號機#{lift_i} 第{stop_i}樓不相同 : #{lift_i}:None {vbCrLf}"
                        End If
                        '---------------------------------------------- 當使用者輸入的HIN為空時 

                        '如果使用者輸入與標準值相同時就先儲存在EachContent陣列中 ----------------------------------------------
                        For i = 1 To arr_liftStopFl_StdContent.Count
                            If arr_liftStopFL_userContent(lift_i - 1, stop_i - 1) = arr_liftStopFl_StdContent(i - 1) Then
                                arr_liftStopFl_EachContent(i - 1, lift_i) = arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)
                                'ResultOutput_TextBox.Text += $"號機#{lift_i} 第{stop_i}樓不相同 : #{lift_i}:{arr_liftStopFl_EachContent(i - 1, lift_i)} {vbCrLf}"
                            End If
                        Next
                        '---------------------------------------------- 如果使用者輸入與標準值相同時就先儲存在EachContent陣列中 
                    Next
                    lift_i = 0

                    '輸出以下值 e.g #1,2:without/#3:with 字樣 -------------------------------------------------
                    If HinLiftDiff_cnt = stopFL_MAX Then
                        ResultOutput_TextBox.Text += $"Hall Indicator {stop_i - 1} FL : {arr_liftStopFL_userContent(lift_i, stop_i - 1)}{vbCrLf}"
                    End If

                    ResultOutput_TextBox.Text += $"Hall Indicator {stop_i} FL : Only 號機  "
                    Dim EachContent_Bool As Boolean
                    For i = 1 To arr_liftStopFl_StdContent.Count
                        EachContent_Bool = False
                        For lift_i = 1 To LiftNum
                            If arr_liftStopFl_EachContent(i - 1, lift_i) <> "" Then
                                ResultOutput_TextBox.Text += $"#{lift_i},"
                                EachContent_Bool = True
                            End If
                        Next
                        If EachContent_Bool And arr_liftStopFl_EachContent(i - 1, 0) <> "" Then
                            ResultOutput_TextBox.Text += $":{arr_liftStopFl_EachContent(i - 1, 0)}/"
                        End If
                    Next

                    If HinLiftDiff_cnt = stopFL_MAX Then
                        topFL_End_bool = True
                    Else
                        topFL_End_bool = False
                    End If
                    ResultOutput_TextBox.Text += $"{vbCrLf}"
                    '------------------------------------------------- 輸出以下值 e.g #1,2:without/#3:with 字樣 

                ElseIf HinLiftDiff_bool = False Then '表示同樓層號機之間值都相同

                    lift_i = 1
                    HinLiftSame_cnt += 1
                    If HinLiftSame_cnt = 1 Then
                        If stop_i = 1 Then '最底樓層
                            ResultOutput_TextBox.Text +=
                            $"Hall Indicator BOTTOM FL : {arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}"
                        Else '當其他樓層從HinLiftSame_cnt = 1開始
                            ResultOutput_TextBox.Text +=
                            $"Hall Indicator {stop_i} FL : {arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}"
                        End If
                    ElseIf HinLiftSame_cnt = 2 Then
                        If HinFLDiff_bool Then
                            'HinLiftSame_cnt = 0
                            ResultOutput_TextBox.Text +=
                                $"Hall Indicator {stop_i} FL : {arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}"
                        End If
                    ElseIf HinLiftSame_cnt > 2 Then
                        If HinPoint_bool = False Then
                            ResultOutput_TextBox.Text += $".........{vbCrLf}"
                            HinPoint_bool = True
                        End If
                        If HinFLDiff_bool Then
                            'HinLiftSame_cnt = 0
                            ResultOutput_TextBox.Text +=
                                $"Hall Indicator {stop_i} FL : {arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}"
                        End If
                    End If

                    If HinFLDiff_bool Then
                        HinLiftSame_cnt = 0
                    End If
                End If
                '--------------------------------- 每次換樓層時判斷 #1~#N 號機該樓層是否都相同?

            Next

            Dim test As Integer
            If lift_i = 1 Then
                test = 1
            Else
                test = 2
            End If

            If topFL_End_bool = False And HinFLDiff_bool = False Then
                ResultOutput_TextBox.Text +=
                $"Hall Indicator TOP FL : {arr_liftStopFL_userContent(lift_i - test, stop_i - 2)}{vbCrLf}"
            End If
        End If
    End Sub

    Private Sub FinalCheck_Button_Click(sender As Object, e As EventArgs) Handles FinalCheck_Button.Click
        'ResultFailOutput_TextBox.Text = ""
        'ResultOutput_TextBox.Text = ""

        Resize_JMForm(JMForm_size.re_size)
        Try
            errorInfo.writeInfo_toTextBox_focusOnBelow(ResultOutput_TextBox,
                                                       $"「最後檢查」 開始 =========================")
            errorInfo.writeInfo_toTextBox_focusOnBelow(ResultFailOutput_TextBox,
                                                       $"「最後檢查」 開始 =========================")

            '基本
            If Use_Basic_CheckBox.Checked And Basic_GroupBox.Enabled Then
                Check_cb_tb_are_empty_in_mCtrl(Basic_GroupBox, Basic_TabPage)
                errorInfo.writeInfo_toTextBox_focusOnBelow(ResultOutput_TextBox,
                                                           $"工番號 : {Basic_JobNoNew_TextBox.Text}")
                errorInfo.writeInfo_toTextBox_focusOnBelow(ResultOutput_TextBox,
                                                           $"工番名 : {Basic_JobName_TextBox.Text}")
            End If

            'CheckList ----------------------------------------------------------------------
            If Use_ChkList_CheckBox.Checked Then
                Dim chkList_radioBtn() As RadioButton
                chkList_radioBtn = {ChkList_1_no_RadioButton, ChkList_2_no_RadioButton,
                                    ChkList_3_no_RadioButton, ChkList_5_no_RadioButton,
                                    ChkList_6_no_RadioButton, ChkList_7_no_RadioButton}
                Dim chkList_ctrl() As Control
                chkList_ctrl = {ChkList_1_Panel, ChkList_2_Panel,
                                ChkList_3_Panel, ChkList_5_Panel,
                                ChkList_6_Panel, ChkList_7_Panel}

                For chk_i As Integer = 1 To (chkList_radioBtn).Count
                    If chkList_radioBtn(chk_i - 1).Checked = False And chkList_ctrl(chk_i - 1).Enabled Then
                        Check_cb_tb_are_empty_in_mCtrl(chkList_ctrl(chk_i - 1), CheckList2_TabPage)
                    End If
                Next
            End If
            '---------------------------------------------------------------------- CheckList 

            '仕樣-Basic all ----------------------------------------------------------------------
            If Use_SpecBasic_CheckBox.Checked Then
                Dim spec_ctrl() As Control
                spec_ctrl = {SpecBasic_LiftItem_Dynamic_Panel,
                             Spec_MachineType_Panel,
                             Spec_ControlWay_Panel,
                             Spec_Purpose_Panel}
                For Each sc In spec_ctrl
                    If sc.Enabled Then
                        Check_cb_tb_are_empty_in_mCtrl(sc, Spec_BasicAll_TabPage)
                    End If
                Next
            End If
            '---------------------------------------------------------------------- 仕樣-Basic all 

            '仕樣-TW ----------------------------------------------------------------------
            If Use_SpecTWFP17_CheckBox.Checked Or Use_SpecTWIDU_CheckBox.Checked Then
                errorInfo.writeInfo_toTextBox_focusOnBelow(ResultOutput_TextBox,
                                                           $"<仕樣確認>")
                Dim spec_item As Spec_Item = New Spec_Item
                spec_item.ini_specTW_AllControler()
                Dim replaceName_Label, replace_ComboBox As String
                For Each mPanel As Control In spec_item.specTW_panel
                    If mPanel.Enabled Then
                        replace_ComboBox = spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mPanel, "Panel", "ComboBox")
                        replaceName_Label = spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mPanel, "Panel", "Label")
                        Check_cb_tb_are_empty_in_mCtrl_if_mCmbbox_is_O(mPanel,
                                                                       "Panel",
                                                                       Spec_TW_TabPage)


                        If spec_item.getRelace_NameText_onPanel(replace_ComboBox, mPanel) = get_nameManager.TB_O Then
                            errorInfo.writeInfo_toTextBox_focusOnBelow(ResultOutput_TextBox,
                                $"    {spec_item.getRelace_NameText_onPanel(replaceName_Label, mPanel)} : {spec_item.getRelace_NameText_onPanel(replace_ComboBox, mPanel)}")
                        End If
                    End If
                Next


            End If
            '---------------------------------------------------------------------- 仕樣-TW 

            '重要設定 ----------------------------------------------------------------------
            If Use_Imp_CheckBox.Checked And ImpSetting_GroupBox.Enabled Then
                Check_cb_tb_are_empty_in_mCtrl(ImpSetting_GroupBox, Important_TabPage)
            End If
            '---------------------------------------------------------------------- 重要設定 

            'MMIC ----------------------------------------------------------------------
            If Use_Imp_CheckBox.Checked Then
                Dim mmic_ctrl() As Control
                mmic_ctrl = {MMIC_GroupBox,
                             MMIC_MR_GroupBox, MMIC_MR_Panel,
                             MMIC_MR_E_GroupBox, MMIC_MR_E_Panel,
                             MMIC_SV_GroupBox, MMIC_SV_Panel,
                             MMIC_SV_E_GroupBox, MMIC_SV_E_Panel,
                             MMIC_VD10_GroupBox, MMIC_VD10_Panel}

                For Each sc In mmic_ctrl
                    If sc.Enabled Then
                        Check_cb_tb_are_empty_in_mCtrl(sc, MMIC_TabPage)
                    End If
                Next
            End If
            '---------------------------------------------------------------------- MMIC 

            errorInfo.writeInfo_toTextBox_focusOnBelow(ResultOutput_TextBox,
                                                       $"========================= 「最後檢查」 結束")
            errorInfo.writeInfo_toTextBox_focusOnBelow(ResultFailOutput_TextBox,
                                                       $"========================= 「最後檢查」 結束")

            If Load_Job_JobSelect_RadioButton.Checked Then
                All_OutputButton.Enabled = True
                Spec_OutputButton.Enabled = True
                CheckList_OutputButton.Enabled = True
            End If
            If Load_Job_ChkListSelect_RadioButton.Checked Then
                CheckList_OutputButton.Enabled = True
            End If


            MsgBox($"檢查完成{vbCrLf}空值以紅底顯示，右側對話視窗可參考")
        Catch ex As Exception
            errorInfo.writeTitleIntoError_InfoTxt("JobMaker.FinalCheck_Button_Click")
            errorInfo.writeInfoError_InfoTxt(ex.Message)
            MsgBox("檢查錯誤")
        End Try
    End Sub

    ''' <summary>
    ''' 檢查mCrtl內的ComBoBox&TextBox是否為空? 如果為空則print tabPage 的分頁名稱並 print 哪個沒填
    ''' </summary>
    ''' <param name="use_chkBox"></param>
    ''' <param name="mCtrl"></param>
    ''' <param name="mTabPage"></param>
    Private Sub Check_cb_tb_are_empty_in_mCtrl(mCtrl As Control, mTabPage As TabPage)
        Dim outputTabPage_Bool As Boolean

        For Each ctrl As Control In mCtrl.Controls
            If TypeOf (ctrl) Is TextBox Or TypeOf (ctrl) Is ComboBox Then
                'ctrl.BackColor = SystemColors.Window
                If ctrl.Text = "" Then
                    If outputTabPage_Bool = False Then
                        outputTabPage_Bool = True
                        errorInfo.writeInfo_toTextBox_focusOnBelow(ResultFailOutput_TextBox,
                                                                   $"<{mTabPage.Text}分頁>")
                        'ResultFailOutput_TextBox.Text += $"<{mTabPage.Text}分頁> {vbCrLf}"

                    End If
                    ctrl.BackColor = Color.Red
                    errorInfo.writeInfo_toTextBox_focusOnBelow(ResultFailOutput_TextBox,
                                                               $"      {ctrl.Name} 沒填 {vbCrLf}")
                    'ResultFailOutput_TextBox.Text += $"      {ctrl.Name} 沒填 {vbCrLf}"
                Else
                    ctrl.BackColor = SystemColors.Window
                End If
            End If
        Next

    End Sub

    ''' <summary>
    ''' 如果mCmbBox是O或空的話，檢查mCrtl內的ComBoBox&TextBox是否為空? 如果為空則輸出tabPage 的分頁名稱並輸出哪個沒填
    ''' </summary>
    ''' <param name="mCmbBox"></param>
    ''' <param name="mCtrl"></param>
    ''' <param name="mTabPage"></param>
    Private Sub Check_cb_tb_are_empty_in_mCtrl_if_mCmbbox_is_O(mCtrl As Control,
                                                               mCtrlType As String,
                                                               mTabPage As TabPage)
        Dim outputTabPage_Bool As Boolean
        Dim spec_item As Spec_Item = New Spec_Item
        'Dim replace_TextBox, replace_ComboBox, replace_Panel As String

        For Each ctrl As Control In mCtrl.Controls
            If TypeOf (ctrl) Is TextBox Or TypeOf (ctrl) Is ComboBox Then
                If ctrl.Text = "" Then
                    If outputTabPage_Bool = False Then
                        '只輸出一次
                        outputTabPage_Bool = True
                        ResultFailOutput_TextBox.Text += $"<{mTabPage.Text}分頁> {vbCrLf}"
                    End If
                    ctrl.BackColor = Color.Red
                    ResultFailOutput_TextBox.Text += $"      {ctrl.Name} 沒填 {vbCrLf}"
                Else
                    ctrl.BackColor = SystemColors.Window
                End If

            End If
        Next

    End Sub


    Private Sub Output_select_spec_to_resultTextbox(mTitle As String, mContent As String)
        ResultOutput_TextBox.Text += $"{mTitle}:{mContent}{vbCrLf}"
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs)
        'Dim string1 As String = "Hello World"
        'MsgBox(string1.Equals("Hello World"))
        With EepData_Speed_TextBox
            AddHandler .MouseEnter, AddressOf TextBox_ResizeHeight_MouseEnter
            AddHandler .MouseLeave, AddressOf TextBox_ResizeHeight_MouseLeave
        End With
    End Sub

    Private Sub Spec_MachineType_Label_MouseEnter(sender As Object, e As EventArgs) Handles Spec_MachineType_Label.MouseEnter


        'For Each pic As Control In SpecBasic_GroupBox2.Controls
        '    If pic.Name = "Spec_MachineType_PictureBox" Then
        '        Dim picBox As New PictureBox
        '        With picBox
        '            .Name = "Spec_MachineType_PictureBox"
        '            .SizeMode = PictureBoxSizeMode.Zoom
        '            .Width = 350
        '            .Height = 150
        '            .Left = 50 'Cursor.Position.X \ 10
        '            .Top = 53 'Cursor.Position.Y \ 10
        '            '.Image = New Bitmap("M:\DESIGN\BACK UP\yc_tian\Tool Application\Tool main program\pic\Spec_MachineType.jpg")
        '            .ImageLocation = "M:\DESIGN\BACK UP\yc_tian\Tool Application\Tool main program\pic\Spec_MachineType.jpg"
        '            SpecBasic_GroupBox2.Controls.Add(picBox)
        '            .BringToFront()
        '        End With
        '    End If
        'Next
        'MsgBox($"{Cursor.Position.X},{Cursor.Position.Y}")
    End Sub





    Private Sub Imp_DoorType_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Imp_DoorType_CheckBox.CheckedChanged
        If Imp_DoorType_CheckBox.Checked Then
            Imp_DoorType_TextBox.Enabled = True
        Else
            Imp_DoorType_TextBox.Enabled = False
        End If
    End Sub


    Dim Spec_EscapeFL_TextBox_height, Spec_Fire_Panel_height As Integer
    Private Sub Spec_EscapeFL_TextBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_EscapeFL_TextBox.TextChanged
        Textbox_AutoSize_withPanel(Spec_EscapeFL_TextBox, Spec_EscapeFL_TextBox_height,
                                   Spec_Fire_Panel, Spec_Fire_Panel_height)
    End Sub

    Dim Spec_MFLReturn_FL_TextBox_height, Spec_MFLReturn_Panel_height As Integer
    Private Sub Spec_MFLReturn_FL_TextBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_MFLReturn_FL_TextBox.TextChanged
        Textbox_AutoSize_withPanel(Spec_MFLReturn_FL_TextBox, Spec_MFLReturn_FL_TextBox_height,
                                   Spec_MFLReturn_Panel, Spec_MFLReturn_Panel_height)
        If Spec_MFLReturn_FL_TextBox.Text <> "" Then
            Spec_MFLReturn_FL_Only_CheckBox.Enabled = True
        Else
            Spec_MFLReturn_FL_Only_CheckBox.Enabled = False
        End If
    End Sub

    Dim Spec_Flood_FL_TextBox_height, Spec_Flood_Panel_height As Integer
    Private Sub Spec_Flood_FL_TextBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Flood_FL_TextBox.TextChanged
        Textbox_AutoSize_withPanel(Spec_Flood_FL_TextBox, Spec_Flood_FL_TextBox_height,
                                   Spec_Flood_Panel, Spec_Flood_Panel_height)
    End Sub
    Dim Spec_Parking_FL_TextBox_height, Spec_Parking_Panel_height As Integer
    Private Sub Spec_Parking_FL_TextBox_TextChanged(sender As Object, e As EventArgs) Handles Spec_Parking_FL_TextBox.TextChanged
        Textbox_AutoSize_withPanel(Spec_Parking_FL_TextBox, Spec_Parking_FL_TextBox_height,
                                   Spec_Parking_Panel, Spec_Parking_Panel_height)
    End Sub




    ''' <summary>
    ''' textbox增加行列時會autosize，panel也會同步更改height
    ''' </summary>
    ''' <param name="textbox"></param>
    ''' <param name="textbox_height"></param>
    ''' <param name="panel"></param>
    ''' <param name="panel_height"></param>
    Private Sub Textbox_AutoSize_withPanel(textbox As TextBox, textbox_height As Integer, panel As Panel, panel_height As Integer)
        With textbox
            If textbox_height <> 0 And panel_height <> 0 Then
                If .Lines.Length >= 1 Then
                    .Height = textbox_height * .Lines.Length
                    panel.Height = panel_height + textbox_height * (.Lines.Length - 1)
                End If
            End If
        End With
    End Sub


    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If Spec_LiftNum_NumericUpDown.Value >= 1 Then
            Dim mdir As New Dictionary(Of String, String)
            For Each items As Control In SpecBasic_LiftItem_Dynamic_Panel.Controls
                For i As Integer = 1 To LiftNum
                    If items.Name = $"{Spec_Control_ComboBox.Name}_{i}" Then
                        If mdir.ContainsKey(items.Text) Then
                            For Each items_2 As Control In SpecBasic_LiftItem_Dynamic_Panel.Controls
                                If items_2.Name = $"{Spec_LiftName_TextBox.Name}_{i}" Then
                                    mdir(items.Text) = mdir.Item(items.Text) & $",{items_2.Text}"
                                    Exit For
                                End If
                            Next
                        Else
                            For Each items_2 As Control In SpecBasic_LiftItem_Dynamic_Panel.Controls
                                If items_2.Name = $"{Spec_LiftName_TextBox.Name}_{i}" Then
                                    mdir.Add(items.Text, items_2.Text)
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Next
            Next
            Dim outputText = ""
            Dim pair As KeyValuePair(Of String, String)
            For Each pair In mdir
                Console.WriteLine("Key={0}, Vale={1}", pair.Key, pair.Value)
                outputText += $"{pair.Value}:{pair.Key}{vbCrLf}"
            Next
            MsgBox(outputText)
        End If
    End Sub

    Private Sub 問題回報ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 問題回報ToolStripMenuItem.Click
        MagicTool.open_DirectPath("M:\DESIGN\BACK UP\yc_tian\Tool Application\Tool update folder\Job_Problem_Report.xlsx")
    End Sub
    Private Sub 查看錯誤回報ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 查看錯誤回報ToolStripMenuItem.Click
        MagicTool.open_DirectPath($"{Application.StartupPath}\{ProgramAllName.fileName_ErrorInfo}")
    End Sub
    '----------------------------------------------------------------------------------------------------------------------- 其他事件 
End Class