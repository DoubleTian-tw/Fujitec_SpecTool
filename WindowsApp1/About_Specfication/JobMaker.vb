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
    Dim finalCheck_Btm_clickTimes As Integer

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

    'EXCEL use
    Dim msExcel_app As Excel.Application
    Dim msExcel_workbook As Excel.Workbook
    Dim msExcel_worksheet As Excel.Worksheet

    '--- 仕樣書 ----------------------------------------------------------------------------------------------------

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
    Public arr_liftTopFL() As Integer  '
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

    'Dim conNum_tb_temp(,) As TextBox '暫存


    ''' <summary>
    ''' 原始或變更後表單大小
    ''' </summary>
    Enum JMForm_size
        ''' <summary>
        ''' 原始大小
        ''' </summary>
        ini_size
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
                ResultCheck_Button.Visible = False

            Case mysize.re_size
                Me.Width = reForm_width
                JobMaker_Close_Button.Location = New Point(reCloseBtn_X, iniCloseBtn_Y)
                ResultOutput_TextBox.Visible = True
                ResultCheck_Button.Visible = True
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
                currentEmployee_Number = "TWN2100" Or
                currentEmployee_Number = "TWN2100" Then
                Button1.Visible = True
                Button2.Visible = True
            End If
            '----------------------------------------------------------------------------------- 判斷工號

            ' SQLite 遺失 --------------------------------
            Dim fileExitPath As String
            fileExitPath = get_nameManager.SQLite_connectionPath_Tool & get_nameManager.SQLite_ToolDBMS_Name
            If Not File.Exists(fileExitPath) Then
                MsgBox($"未取得Sqlite檔案請確認路徑: {fileExitPath} 是否正確?")
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

        '初始化 Load > 仕樣書 分頁 ------------------------
        With JMFileCho_Spec_TextBox
            .Text = Load_info_txt
            .ForeColor = Color.Gray
        End With

        With JM_DefaultPath_Spec_Label
            .Text = chalink.ChgLink_DefaultPath_Spec_TextBox.Text
        End With
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
        With JMFileCho_SQLite_TextBox
            .Text = Load_info_txt
            .ForeColor = Color.Gray
        End With
        With JM_DefaultPath_SQLite_Label
            .Text = get_nameManager.SQLite_connectionPath_Tool
        End With
        '---------------------初始化 Load > 載入SQLite 分頁


        '---------------------------------- 初始化 Load 分頁 結束


        '初始化 基本 分頁 開始 -----------------------------------
        ReminderMarquee_Label.Text = $"{currentEmployee_ChineseName} 舉起你的雙手(｢･ω･)｢...."
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
            .Controls.SetChildIndex(use_ProgramChg_Panel5, 3)
        End With
        '----------------------------------- 初始化 程式變更表 分頁 結束

        '初始化 送狀 分頁 開始 -----------------------------------
        DWG_PrkName_ComboBox.Items.Clear()
        '----------------------------------- 初始化 送狀 分頁 結束

        '初始化 仕樣 分頁 開始 -----------------------------------
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
        Spec_LoadCell_ComboBox.Text = Spec_LoadCell_ComboBox.Items(0)                   'Load Cell
        Spec_install_ope_ComboBox.Text = Spec_install_ope_ComboBox.Items(0)             '拒付運轉
        Spec_FireSignal_ComboBox.Text = Spec_FireSignal_ComboBox.Items(0)               '火災運轉訊號
        Spec_ParkingFL_DR_ComboBox.Text = Spec_ParkingFL_DR_ComboBox.Items(1)           'Parking休止開關門
        Spec_CRDSpec_ComboBox.Text = Spec_CRDSpec_ComboBox.Items(0)                     '刷卡機仕樣有無
        Spec_CRDReg_ComboBox.Text = Spec_CRDReg_ComboBox.Items(1)                       '刷卡機自動登陸有無
        Spec_CRDCancell_ComboBox.Text = Spec_CRDCancell_ComboBox.Items(1)               '刷卡機逆向呼叫無效
        Spec_CRDNuisance_ComboBox.Text = Spec_CRDNuisance_ComboBox.Items(0)             '刷卡機防嬉戲
        '----------------------------------- 初始化 仕樣 分頁 結束 
    End Sub


    'LOAD ------------------------------------------------------------------------------------------------------------
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
            finalCheck_Btm_clickTimes = 0
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
        For Each path In file
            If System.IO.File.Exists(path) Then
                JMFileCho_AutoLoad_TextBox.Text = path
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
    Private Sub JM_SQlite_JobSelect_TextBox_TextChanged(sender As Object, e As EventArgs) Handles JM_JobSelect_SQLite_TextBox.TextChanged
        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        JobSelect_type_into_textBox({"*.sqlite"},
                                    spec_stored.SQLite_connectionPath_Job,
                                    JM_JobSelect_SQLite_ComboBox, JM_JobSelect_SQLite_TextBox)
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

    Private Sub JM_SQlite_JobSelect_ComboBox_TextChanged(sender As Object, e As EventArgs) Handles JM_JobSelect_SQLite_ComboBox.TextChanged
        'Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        JobSelect_add_into_comboBox_and_textBox(JM_DefaultPath_SQLite_Label.Text,
                                                JM_JobSelect_SQLite_ComboBox,
                                                JMFileCho_SQLite_TextBox)
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
    Private Sub JMFileCho_SQLite_Button_DragDrop(sender As Object, e As DragEventArgs) Handles JMFileCho_SQLite_TextBox.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In file
            If System.IO.File.Exists(path) Then
                JMFileCho_SQLite_TextBox.Text = path
                JMFileCho_SQLite_TextBox.ForeColor = Color.Black
            End If
        Next
    End Sub
    ''' <summary>
    ''' [DragEnter][Load > 載入SQLite > 路徑]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileCho_SQLite_Button_DragEnter(sender As Object, e As DragEventArgs) Handles JMFileCho_SQLite_TextBox.DragEnter
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
    Private Sub JMFileCho_SQLite_TextBox_TextChanged(sender As Object, e As EventArgs) Handles JMFileCho_SQLite_TextBox.TextChanged
        If JMFileCho_SQLite_TextBox.Text <> Load_info_txt Then
            If JMFileCho_SQLite_TextBox.Text <> "" Then
                JMFileConfirm_SQLite_Button.Enabled = True
                Check_direction_file_is_needed_type({"sqlite"}, JMFileCho_SQLite_TextBox)
            End If
        Else
                JMFileConfirm_SQLite_Button.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' [Load > 載入SQLite > CheckBox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JobMaker_LOAD_SQLite_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles JobMaker_LOAD_SQLite_CheckBox.CheckedChanged
        If JobMaker_LOAD_SQLite_CheckBox.Checked Then
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

        ChangeLink.OpenFile_event(JMFileCho_SQLite_TextBox,
                                  ChangeLink.OpenFileType.mOther,
                                  mpath)

        If JMFileCho_SQLite_TextBox.Text <> "" Then
            JMFileConfirm_SQLite_Button.Enabled = True
        End If
    End Sub
    ''' <summary>
    ''' [Load > 載入SQLite > 確認Button]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JMFileConfirm_SQLite_Button_Click(sender As Object, e As EventArgs) Handles JMFileConfirm_SQLite_Button.Click
        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        Resize_JMForm(JMForm_size.re_size)
        spec_stored.Load_Stored(Path.GetFileName(JMFileCho_SQLite_TextBox.Text))
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
            Output_new_excel_and_open_from_textbox(JMFileCho_Spec_TextBox.Text)
            'msExcel_app.Visible = True

            Resize_JMForm(JMForm_size.re_size) '重新變大小
            output_ToSpec.Spec_FinalCheck(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_Spec_Std(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_Basic(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_TW(LiftNum, ContainNum, msExcel_workbook, msExcel_app)

            Output_open_excel_folder_and_save_when_done(JMFileCho_Spec_TextBox)
        Catch ex As Exception
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

            Output_open_excel_folder_and_save_when_done(JMFileCho_ChkList_TextBox)
        Catch ex As Exception
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
            Output_new_excel_and_open_from_textbox(JMFileCho_Spec_TextBox.Text)
            'msExcel_app.Visible = True

            Resize_JMForm(JMForm_size.re_size) '重新變大小
            output_ToSpec.Spec_FinalCheck(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_Spec_Std(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_Basic(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_SPEC_TW(LiftNum, ContainNum, msExcel_workbook, msExcel_app)
            'output_ToSpec.Spec_DWG(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_Important(msExcel_workbook, msExcel_app)
            output_ToSpec.Spec_MMIC(msExcel_workbook, msExcel_app)


            Output_open_excel_folder_and_save_when_done(JMFileCho_Spec_TextBox)
        Catch ex As Exception
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
        Try
            Output_new_excel_and_open_from_textbox(JMFileCho_Spec_TextBox.Text)
            msExcel_app.Visible = True

            'Dim excelMath As getMath_onExcel = New getMath_onExcel
            'excelMath.setValue_to_Cells_onWorksht(msExcel_workbook, "JOBNO", EepData_MachineRoom_TextBox.Text)

            '取得 名稱管理員specName Range的頭例如A4的A


            'getMath_onExcel.convertColumn_fromIntToString(startRange_Col)
            '取得該合併儲存格的數量
            'merge_num =
            'msExcel_workbook.Worksheets(startWorksheet_name).range(startRange_Row & startRange_Col).MergeArea.Rows.Count
            'msExcel_workbook.Save()
            'output_ToSpec.Spec_FinalCheck(msExcel_workbook, msExcel_app)
            'output_ToSpec.Spec_SPEC_Basic(msExcel_workbook, msExcel_app)
            'output_ToSpec.Spec_SPEC_TW(LiftNum, ContainNum, msExcel_workbook, msExcel_app)

            'Output_open_excel_folder_and_save_when_done(JMFileCho_Spec_TextBox)
        Catch ex As Exception
        Finally

            'generation = GC.GetGeneration(msExcel_app)
            'If (msExcel_workbook IsNot Nothing) Then
            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(msExcel_workbook)
            '    msExcel_workbook = Nothing
            'End If
            'If msExcel_app IsNot Nothing Then
            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(msExcel_app)
            '    msExcel_app = Nothing
            'End If
            'GC.Collect()
            'GC.WaitForPendingFinalizers()

            'Output_kill_excel_when_done()
        End Try
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

            If use_chkList_chkbox_clickTimes = 1 Then
                ChkList_Confirm_CheckBox.Checked = True     '確認圖ChkBox
                ChkList_1_no_RadioButton.Checked = True     '1 不清楚仕樣
                ChkList_2_no_RadioButton.Checked = True     '2 法規、安全
                ChkList_3_no_RadioButton.Checked = True     '3 迴路圖面是否不清楚
                ChkList_5_no_RadioButton.Checked = True     '5 VONIC
                ChkList_6_no_RadioButton.Checked = True     '6 確認式樣動作
                ChkList_7_no_RadioButton.Checked = True     '7 參考資料
                ChkList_8_yes_RadioButton.Checked = True    '8 最後確認
                ChkList_8Item_RadioButton.Checked = True    '8 滿足特記事項
                ChkList_9_yes_RadioButton.Checked = True    '9 自我檢查表
            End If
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
    ''' <summary>
    ''' [CheckList > 3.電器不清楚 > 有，討論者]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_3_yes_Man_TextBox_TextChanged(sender As Object, e As EventArgs) Handles ChkList_3_yes_Man_TextBox.TextChanged
        If ChkList_3_yes_RadioButton.Checked = False Then
            If ChkList_3_yes_Content_TextBox.Text <> "" Or ChkList_3_yes_Result_TextBox.Text <> "" Or ChkList_3_yes_Man_TextBox.Text Then
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
    ''' [CheckList > 5.VONIC > 標準]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_5_std_Content_TextBox_TextChanged(sender As Object, e As EventArgs)
        If ChkList_5_std_RadioButton.Checked = False Then
            If ChkList_5_std_Content_TextBox.Text <> "" Then
                ChkList_5_std_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 5.VONIC > 工直]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_5_nstd_Content_TextBox_TextChanged(sender As Object, e As EventArgs)
        If ChkList_5_nstd_RadioButton.Checked = False Then
            If ChkList_5_nstd_Content_TextBox.Text <> "" Then
                ChkList_5_nstd_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 6.確認 > 有，檢驗項目]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_6_yes_Content_TextBox_TextChanged(sender As Object, e As EventArgs)
        If ChkList_6_yes_RadioButton.Checked = False Then
            If ChkList_6_yes_Content_TextBox.Text <> "" Then
                ChkList_6_yes_RadioButton.Checked = True
                ChkList_6_yesItem_RadioButton.Checked = True
            End If
        End If
    End Sub
    ''' <summary>
    ''' [CheckList > 7.參考資料 > 有，文書]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChkList_7_yes1_content_TextBox_TextChanged(sender As Object, e As EventArgs)
        If ChkList_7_yes_RadioButton.Checked = False Then
            If ChkList_7_yes1_content_TextBox.Text <> "" Then
                ChkList_7_yes_RadioButton.Checked = True
            End If
        End If
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
            ProgramChange_FlowLayoutPanel.Enabled = True

            If use_program_chkbox_clickTimes = 1 Then
                PrmList_2_test_CheckBox.Checked = True     '測試裝置
                PrmList_3_debug_CheckBox.Checked = True    'DEBUG
                PrmList_3_confirm_CheckBox.Checked = True  '一般動作確認
                PrmList_3_excute_CheckBox.Checked = True   '確認程式執行
                PrmList_4_yes1_RadioButton.Checked = True  '4-1 手動全自動
                PrmList_4_yes2_RadioButton.Checked = True  '4-2 入出力點一致
                PrmList_4_yes3_RadioButton.Checked = True  '4-3 變數初始化
                PrmList_4_yes4_RadioButton.Checked = True  '4-4 OTHER的CASE
                PrmList_4_yes5_RadioButton.Checked = True  '4-5 ELSE IF
                PrmList_4_yes6_RadioButton.Checked = True  '4-6 LOOP
                PrmList_4_yes7_RadioButton.Checked = True  '4-7 範圍內
                PrmList_4_no8_RadioButton.Checked = True   '4-8 CASTING
                PrmList_4_no9_RadioButton.Checked = True   '4-9 0除
                PrmList_4_yes10_RadioButton.Checked = True '4-10 運算子
                PrmList_4_yes11_RadioButton.Checked = True '4-11 ADDRESS
                PrmList_4_yes12_RadioButton.Checked = True '4-12 要求仕樣
            End If
        Else
            ProgramChange_FlowLayoutPanel.Enabled = False
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
    '--------------------------------------------------------------------------------------------------------------------程式變更 


    '仕樣 -------------------------------------------------------------------------------------------------------------------- 
    Private Sub Spec_MachineType_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Spec_MachineType_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        'dyCtrlName.JobMaker_SpecInfo()
        '機種
        AddSub_Object_Sub(Spec_MachineType_NumericUpDown,
                          Spec_MachineType_Panel,
                          {Spec_Base_ComboBox},
                          {dyCtrlName.Spec_MachineType_ComboBox}.Count,
                          {dyCtrlName.Spec_MachineType_ComboBox},
                          {get_nameManager.SQLite_tableName_Basic},
                          {get_nameManager.Spec_MachineType})
        '控制方式
        AddSub_Object_Sub(Spec_MachineType_NumericUpDown,
                          Spec_ControlWay_Panel,
                          {Spec_Base_ComboBox},
                          {dyCtrlName.Spec_ControlWay_ComboBox}.Count,
                          {dyCtrlName.Spec_ControlWay_ComboBox},
                          {get_nameManager.SQLite_tableName_Basic},
                          {get_nameManager.Spec_ControlWay})
    End Sub

    Private Sub Spec_Purpose_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Spec_Purpose_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName

        AddSub_Object_Sub(Spec_Purpose_NumericUpDown,
                          Spec_Purpose_Panel,
                          {Spec_Base_ComboBox},
                          {dyCtrlName.Spec_Purpose_ComboBox}.Count,
                          {dyCtrlName.Spec_Purpose_ComboBox},
                          {get_nameManager.SQLite_tableName_Basic},
                          {get_nameManager.Spec_Purpose})
    End Sub
    ''' <summary>
    ''' [仕樣 > TW > NumericUpDown >自家發 ]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_EmerNum_NumericUpDown_ValueChanged(sender As Object, e As EventArgs)
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
        Catch
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
                Spec_Base_ComboBox.Enabled = True '機種

                Use_SpecTWIDU_CheckBox.CheckState = CheckState.Unchecked

                'Spec_OutputButton.Enabled = True
                Spec_IF79x_Panel.Enabled = False    'IF79入出力位置
                Spec_EachStop_Panel.Enabled = False '各停開關
                Spec_WTB_Panel.Enabled = False      'WTB
                Spec_LoadCell_Panel.Enabled = False 'Load Cell

            Else
                If Use_SpecTWIDU_CheckBox.CheckState = CheckState.Unchecked Then
                    Spec_TW_FlowLayoutPanel1.Enabled = False
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
    ''' <summary>
    ''' [仕樣 > TW > 避難階]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_EscapeFL_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Fireman_Only_CheckBox.CheckedChanged
        If Spec_Fireman_Only_CheckBox.Checked Then
            Spec_Fireman_Only_TextBox.Enabled = True
        Else
            Spec_Fireman_Only_TextBox.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' [仕樣 > Basic All > 電梯總數]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_LiftNum_NumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles Spec_LiftNum_NumericUpDown.ValueChanged
        Dim dyCtrlName As DynamicControlName = New DynamicControlName

        Dim TextBoxWidth, TextBox_XPosition, TextBox_YPosition As Integer()


        'Dim StartPos_x As Integer = Spec_LiftName_TextBox.Location.X '10 '起始第一格left寬度
        'Dim StartPos_y As Integer = 10 '起始top高度
        Dim ConNum_tb As TextBox
        Dim ConNum_cb As ComboBox
        '顯示內容數量及文字

        TextBoxWidth = {Spec_LiftName_TextBox.Width, Spec_LiftMem_TextBox.Width, Spec_Control_ComboBox.Width,
                        Spec_TopFL_TextBox.Width, Spec_BtmFL_TextBox.Width, Spec_StopFL_TextBox.Width,
                        Spec_Speed_TextBox.Width, Spec_FLName_TextBox.Width}
        TextBox_XPosition = {Spec_LiftName_TextBox.Left, Spec_LiftMem_TextBox.Left, Spec_Control_ComboBox.Left,
                             Spec_TopFL_TextBox.Left, Spec_BtmFL_TextBox.Left, Spec_StopFL_TextBox.Left,
                             Spec_Speed_TextBox.Left, Spec_FLName_TextBox.Left}
        TextBox_YPosition = {Spec_LiftName_TextBox.Top, Spec_LiftMem_TextBox.Top, Spec_Control_ComboBox.Top,
                             Spec_TopFL_TextBox.Top, Spec_BtmFL_TextBox.Top, Spec_StopFL_TextBox.Top,
                             Spec_Speed_TextBox.Top, Spec_FLName_TextBox.Top}
        '嘗試得到電梯輸入之總數
        Try
            LiftNum = Spec_LiftNum_NumericUpDown.Value
            'LiftNum_Panel.Controls.Clear()
        Catch
        End Try

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
                    'For j As Integer = 1 To dyCtrlName.JobMaker_LiftInfoName_Array.Count
                    ConNum_tb = New TextBox()
                    ConNum_cb = New ComboBox()
                    Select Case ctrlName.GetType
                        Case GetType(ComboBox)
                            With ConNum_cb
                                .Width = ctrlName.Width 'TextBoxWidth(j - 1)
                                .Left = ctrlName.Left 'TextBox_XPosition(j - 1)
                                .Top = ctrlName.Top + (i - 1) * 30 'TextBox_YPosition(j - 1) + (i - 1) * 30
                                .Font = New System.Drawing.Font("微軟正黑體",
                                                                9.0!,
                                                                System.Drawing.FontStyle.Regular,
                                                                System.Drawing.GraphicsUnit.Point,
                                                                CType(136, Byte))
                                .Name = $"{ctrlName.Name}_{i}" '($"{dyCtrlName.JobMaker_LiftInfoName_Array(j - 1)}_{i}")

                                get_nameManager.read_DbmsData(get_nameManager.OperationType,
                                                              get_nameManager.SQLite_tableName_Basic,
                                                              ConNum_cb,
                                                              get_nameManager.SQLite_connectionPath_Tool,
                                                              get_nameManager.SQLite_ToolDBMS_Name)
                            End With
                            SpecBasic_LiftItem_Dynamic_Panel.Controls.Add(ConNum_cb)
                        Case GetType(TextBox)
                            With ConNum_tb

                                .Width = ctrlName.Width 'TextBoxWidth(j - 1)
                                .Left = ctrlName.Left 'TextBox_XPosition(j - 1)
                                .Top = ctrlName.Top + (i - 1) * 30 'TextBox_YPosition(j - 1) + (i - 1) * 30
                                .Font = New System.Drawing.Font("微軟正黑體",
                                                                9.0!,
                                                                System.Drawing.FontStyle.Regular,
                                                                System.Drawing.GraphicsUnit.Point,
                                                                CType(136, Byte))
                                .Name = $"{ctrlName.Name}_{i}" '($"{dyCtrlName.JobMaker_LiftInfoName_Array(j - 1)}_{i}")
                                If ctrlName.Name = Spec_TopFL_Real_TextBox.Name Or
                                   ctrlName.Name = Spec_BtmFL_Real_TextBox.Name Then
                                    .Text = "(  )"
                                End If
                            End With
                            SpecBasic_LiftItem_Dynamic_Panel.Controls.Add(ConNum_tb)
                    End Select
                Next
            Next i
            '---------------------------------- 增加 
        End If
    End Sub

    ''' <summary>
    ''' [仕樣 > TW台灣 > 車廂上到著鈴 > 車廂上]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Spec_CarGong_Top_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_Top_CheckBox.CheckedChanged
        If Spec_CarGong_Top_CheckBox.Checked Then
            'Spec_CarGong_Top_TextBox.Enabled = True
            Spec_CarGong_Top_Only_CheckBox.Enabled = True
            Spec_CarGong_Top_Only_TextBox.Enabled = True
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
            'Spec_CarGong_TopBtm_TextBox.Enabled = True
            Spec_CarGong_TopBtm_Only_CheckBox.Enabled = True
            Spec_CarGong_TopBtm_Only_TextBox.Enabled = True
        Else
            'Spec_CarGong_TopBtm_TextBox.Enabled = False
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
            'Spec_CarGong_COB_TextBox.Enabled = True
            Spec_CarGong_COB_Only_CheckBox.Enabled = True
            Spec_CarGong_COB_Only_TextBox.Enabled = True
        Else
            'Spec_CarGong_COB_TextBox.Enabled = False
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
            'Spec_CarGong_VONIC_TextBox.Enabled = True
            Spec_CarGong_VONIC_Only_CheckBox.Enabled = True
            Spec_CarGong_VONIC_Only_TextBox.Enabled = True
        Else
            'Spec_CarGong_VONIC_TextBox.Enabled = False
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
        Catch

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
        Catch

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

            Dim lift_i, stopFL_i As Integer
            ReDim arr_liftName(LiftNum - 1) 'HIN中自動產生-<樓層名稱>
            ReDim arr_liftStopFL(LiftNum - 1) 'HIN中自動產生-<樓層停止數數量>
            ReDim arr_liftTopFL(LiftNum - 1) 'HIN中自動產生-<樓層頂樓數量>

            Dim dyCtrlName As DynamicControlName = New DynamicControlName
            HallIndicator_FlowLayoutPanel.Controls.Clear() '每啟用就清除表單內容

            If Use_SpecBasic_CheckBox.Checked And Spec_LiftNum_NumericUpDown.Value <> 0 Then '確認<基本仕樣>和<電梯總數>是否使用
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
                            If tempCtrl.Name = $"{Spec_StopFL_TextBox.Name}_{lift_i}" Then
                                arr_liftStopFL(lift_i - 1) = CInt(tempCtrl.Text)
                            End If
                            '----------------- 儲存目前自動產生的<樓層停止數> 
                        Catch ex As Exception
                            MsgBox($"電梯停止數:{tempCtrl.Name},第{lift_i}號機 內容非數字")
                            ResultFailOutput_TextBox.Text += $"電梯停止數:{tempCtrl.Name},第{lift_i}號機 內容非數字"
                        End Try

                        Try
                            '儲存目前自動產生的<樓層頂樓> -----------------
                            If tempCtrl.Name = $"{Spec_TopFL_TextBox.Name}_{lift_i}" Then
                                arr_liftTopFL(lift_i - 1) = tempCtrl.Text
                            End If
                            '----------------- 儲存目前自動產生的<樓層頂樓>
                        Catch ex As Exception
                            MsgBox($"電梯最高樓層:{tempCtrl.Name},第{lift_i}號機 內容非數字")
                            ResultFailOutput_TextBox.Text += $"電梯最高樓層:{tempCtrl.Name},第{lift_i}號機 內容非數字"
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
                        Dim cbox As CheckBox = New CheckBox()
                        Dim cmbBox As ComboBox = New ComboBox()

                        With cbox
                            .AutoSize = True
                            .Text = stopFL_i & "FL"
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
                        flowPanel.Controls.Add(cbox)
                        flowPanel.Controls.Add(cmbBox)
                    Next
                    '------------------------------------------------------ 樓層CheckBox / With Combobox 
                Next
                '---------------------------------------------------------- 新建立<樓層名稱>、<樓層停止數>等資訊的控制項 
            End If
            '--------- 確認<基本仕樣>和<電梯總數>是否使用

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

        If Use_Imp_CheckBox.CheckState = CheckState.Checked Then
            For Each flp In HallIndicator_FlowLayoutPanel.Controls.OfType(Of FlowLayoutPanel)
                For Each chkb In flp.Controls.OfType(Of CheckBox)
                    For Lift_i = 1 To LiftNum
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
                    Next 'lift_i
                Next 'chkb
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
        'dyCtrlName.JobMaker_HINInfoName_Array()
        If Use_Imp_CheckBox.CheckState = CheckState.Checked Then
            For Each flp In HallIndicator_FlowLayoutPanel.Controls.OfType(Of FlowLayoutPanel)
                For Each chkb In flp.Controls.OfType(Of CheckBox)
                    For Lift_i = 1 To LiftNum
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
                                Exit For
                            Else
                                Exit For
                            End If
                        End If
                        '---------------------------------------------- <全樓層都打勾> 動作時跳出迴圈避免資源浪費 
                    Next 'lift_i
                Next 'chkb
            Next 'flp
        End If
        '------------------------------- HIN中自動產生的<全樓層打勾>CheckBox 的event 
    End Sub
    '------------------------------------------------------------------------------------------------------------------------- 重要設定 



    ' MMIC -------------------------------------------------------------------------------------------------------------------------
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
            End If
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
                    get_nameManager.read_DbmsData(get_nameManager.mmicN_FP17_ZR_TW,
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
    Private Sub FLEX_N_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MMIC_FLEX_N_ComboBox.SelectedIndexChanged
        If MMIC_FLEX_N_ComboBox.Text <> "" Then
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
                'if NX100-PC8 THEN WITH CP43X
                MMIC_MR_CP43x_ComboBox.Text = get_nameManager.TB_WITHOUT
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
        Catch
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
        Catch
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
        Catch
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
        'Dim TextBoxWidth, TextBox_XPosition, TextBox_YPosition As Integer()

        'TextBoxWidth = {tb_lift.Width, tb_objName.Width}
        'TextBox_XPosition = {tb_lift.Left, tb_objName.Left}
        'TextBox_YPosition = {tb_lift.Top, tb_objName.Top}

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
                        ConNum = New ComboBox
                    ElseIf TypeOf ctrl(0) Is TextBox Then
                        ConNum = New TextBox
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
                        End If
                        .Name = ($"{dyCrtl_Array(Obj_j - 1)}_{Lift_i}")

                        mpanel.Controls.Add(ConNum)
                    End With
                Next Obj_j
            Next Lift_i
            '---------------------------------- 增加 
        End If
    End Sub

    ''' <summary>
    ''' [MMIC > VD10 > TYPE Combobox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JM_VD10_Type_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
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
        End If
    End Sub

    ''' <summary>
    ''' [MMIC > SV > TYPE Combobox]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub JM_SV_Type_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
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
    Private Sub Use_G_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Use_G_CheckBox.CheckedChanged
        If Use_G_CheckBox.Checked Then
            GWeb_GroupBox.Enabled = True
        Else
            GWeb_GroupBox.Enabled = False
        End If
    End Sub
    '------------------------------------------------------------------------------------------------------------------------- G值


    '其他事件 -----------------------------------------------------------------------------------------------------------------------
    Private Sub JobMaker_Timer_Tick(sender As Object, e As EventArgs) Handles JobMaker_Timer.Tick
        If NumericUpDown1.Value > 0 Then
            JobMaker_Timer.Interval = NumericUpDown1.Value '事件發生間隔透過數值調整設定
            ReminderMarquee_Label.Left = ReminderMarquee_Label.Left - 1
            If ReminderMarquee_Label.Left < 0 - ReminderMarquee_Label.Width / 5 Then
                ReminderMarquee_Label.Left = ReminderMarquee_Label.Width
            End If
        End If
    End Sub

    Private Sub JobMaker_Close_Button_Click(sender As Object, e As EventArgs) Handles JobMaker_Close_Button.Click
        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
        Dim checkFlie_IfExists As Boolean
        'Dim checkFlie_IfExists_String As String

        Dim Stored_result As MsgBoxResult = MsgBox("是否儲存你輸入的工番資料?", vbYesNoCancel, "提醒")
        Dim Stored_Input

        Try
            If Stored_result = MsgBoxResult.Yes Then
                Do
                    Dim jobNo_from_user As String
                    If Basic_JobNoNew_TextBox.Text <> "" Then
                        jobNo_from_user = Basic_JobNoNew_TextBox.Text
                    Else
                        jobNo_from_user = Replace(JM_JobSelect_SQLite_ComboBox.Text, ".sqlite", "")
                    End If
                    Stored_Input = InputBox("輸入Job Name(範例:TW-9453-55)", "儲存新檔", jobNo_from_user)

                    If Stored_Input = "" Then
                        MsgBox("未輸入JobName，請重來")
                    ElseIf Len(Stored_Input) = 0 Then
                        MsgBox("取消")
                    Else
                        '尋找資料夾是否有重複檔案
                        checkFlie_IfExists = File.Exists(spec_stored.SQLite_connectionPath_Job & $"{Stored_Input}.sqlite")
                        If checkFlie_IfExists = False Then
                            'checkFlie_IfExists = False
                            My.Computer.FileSystem.CopyFile(spec_stored.SQLite_connectionPath_Tool & spec_stored.SQLite_StdJobDataDBMS_Name,
                                                            spec_stored.SQLite_connectionPath_Job & $"{Stored_Input}.sqlite")
                            spec_stored.Update_Stored($"{Stored_Input}.sqlite", checkFlie_IfExists)
                            MsgBox($"JobName:{Stored_Input}已存檔")

                            Me.Close()
                        Else
                            Dim checkFile_IfExists_result As MsgBoxResult = MsgBox($"{Stored_Input}已存在，是否覆蓋檔案?", vbYesNo, "提醒")
                            If checkFile_IfExists_result = MsgBoxResult.Yes Then
                                spec_stored.Update_Stored($"{Stored_Input}.sqlite", checkFlie_IfExists)
                                MsgBox($"{Stored_Input}已覆蓋")
                                Me.Close()
                            Else
                                MsgBox($"{Stored_Input}未覆蓋")
                            End If
                        End If
                    End If
                Loop Until Stored_Input <> "" Or Len(Stored_Input) = 0
            ElseIf Stored_result = MsgBoxResult.No Then
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox($"關閉時儲存SQLite錯誤{vbCrLf}{ex}")
        End Try
    End Sub

    ''' <summary>
    ''' [JobMaker > 關閉Debug視窗]
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ResultCheck_Button_Click(sender As Object, e As EventArgs) Handles ResultCheck_Button.Click
        With ResultOutput_TextBox
            .Clear()
            .Visible = False
        End With
        With ResultFailOutput_TextBox
            .Clear()
            .Visible = False
        End With
        With ResultCheck_Button
            .Visible = False
        End With

        Resize_JMForm(JMForm_size.ini_size)
    End Sub

    ''' <summary>
    ''' 測試用
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles HIN_TestButton.Click
        Resize_JMForm(JMForm_size.re_size)

        Dim HinLiftDiff_bool, HinFLDiff_bool As Boolean
        'Dim Hin_LiftDiff_bool, Hin_FLDiff_bool As Boolean
        'Dim Hin_LiftSame_bool, Hin_FLSame_bool As Boolean
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

        'Console.WriteLine($"所有電梯中最大樓層數為:{stopFL_MAX} / 最小樓層數為:{stopFL_MIN}")
        Dim arr_liftStopFL_userContent(LiftNum - 1, stopFL_MAX - 1) As String
        'ResultOutput_TextBox.Text += $"最高樓層數:{stopFL_MAX} 目前陣列數 {arr_liftStopFL_userContent.Length} {vbCrLf}"
        '---------------------------------------------- 求最高樓層 

        Dim dyCtrlName As DynamicControlName = New DynamicControlName
        If HallIndicator_FlowLayoutPanel.Controls.Count <> 0 Then

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
        ResultFailOutput_TextBox.Text = ""
        ResultOutput_TextBox.Text = ""
        finalCheck_Btm_clickTimes += 1

        Resize_JMForm(JMForm_size.re_size)
        Try
            '基本
            If Use_Basic_CheckBox.Checked And Basic_GroupBox.Enabled Then
                Check_cb_tb_are_empty_in_mCtrl(Basic_GroupBox, Basic_TabPage)
                Output_select_spec_to_resultTextbox("工番號:", Basic_JobNoNew_TextBox.Text)
                Output_select_spec_to_resultTextbox("工番名:", Basic_JobName_TextBox.Text)
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
                        Check_cb_tb_are_empty_in_mCtrl(chkList_ctrl(chk_i - 1), CheckList_TabPage)
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

                Output_select_spec_to_resultTextbox("<仕樣確認>", "")
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
                            Output_select_spec_to_resultTextbox(spec_item.getRelace_NameText_onPanel(replaceName_Label, mPanel),
                                                                spec_item.getRelace_NameText_onPanel(replace_ComboBox, mPanel))
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


            'If finalCheck_Btm_clickTimes = 1 Then
            If JobMaker_LOAD_Spec_CheckBox.Checked And JMFileCho_Spec_TextBox.Text <> Load_info_txt Then
                All_OutputButton.Enabled = True
                Spec_OutputButton.Enabled = True
            End If
            If JobMaker_LOAD_ChkList_CheckBox.Checked Then
                CheckList_OutputButton.Enabled = True
            End If
            'End If
            MsgBox($"檢查完成{vbCrLf}空值以紅底顯示，右側對話視窗可參考")
        Catch ex As Exception
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
                        ResultFailOutput_TextBox.Text += $"<{mTabPage.Text}分頁> {vbCrLf}"
                    End If
                    ctrl.BackColor = Color.Red
                    ResultFailOutput_TextBox.Text += $"      {ctrl.Name} 沒填 {vbCrLf}"
                Else
                    ctrl.BackColor = SystemColors.Window
                End If
            End If
        Next
        'JobMaker_TabControl.SelectedTab = Load_TabPage
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
                'If ctrl.Text = get_nameManager.TB_O Or ctrl.Text = "" Then
                '    'If mCmbBox.Text = get_nameManager.TB_O Or mCmbBox.Text = "" Then
                '    If ctrl.Text = "" Then
                '        If outputTabPage_Bool = False Then
                '            '只輸出一次
                '            outputTabPage_Bool = True
                '            ResultFailOutput_TextBox.Text += $"<{mTabPage.Text}分頁> {vbCrLf}"
                '        End If
                '        ctrl.BackColor = Color.Red
                '        ResultFailOutput_TextBox.Text += $"      {ctrl.Name} 沒填 {vbCrLf}"
                '    Else
                '        ctrl.BackColor = SystemColors.Window
                '    End If
                'ElseIf ctrl.Text = get_nameManager.TB_X Or ctrl.Text <> "" Then
                '    'ElseIf mCmbBox.Text = get_nameManager.TB_X Then
                '    ctrl.BackColor = SystemColors.Window
                'End If
                'End If
            End If
        Next
        'JobMaker_TabControl.SelectedTab = Load_TabPage
    End Sub


    Private Sub Output_select_spec_to_resultTextbox(mTitle As String, mContent As String)
        ResultOutput_TextBox.Text += $"{mTitle}:{mContent}{vbCrLf}"
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        'Dim string1 As String = "Hello World"
        'MsgBox(string1.Equals("Hello World"))
        With EepData_Speed_TextBox
            AddHandler .MouseEnter, AddressOf TextBox_ResizeHeight_MouseEnter
            AddHandler .MouseLeave, AddressOf TextBox_ResizeHeight_MouseLeave
        End With
    End Sub


    Private Sub Spec_Fire_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Fire_Only_CheckBox.CheckedChanged
        If Spec_Fire_Only_CheckBox.Checked Then
            Spec_Fire_Only_TextBox.Enabled = True
        Else
            Spec_Fire_Only_TextBox.Enabled = False
        End If
    End Sub


    Private Sub Spec_CpiOLT_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CpiOLT_Only_CheckBox.CheckedChanged
        If Spec_CpiOLT_Only_CheckBox.Checked Then
            Spec_CpiOLT_Only_TextBox.Enabled = True
        Else
            Spec_CpiOLT_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_Seismic_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Seismic_Only_CheckBox.CheckedChanged
        If Spec_Seismic_Only_CheckBox.Checked Then
            Spec_Seismic_Only_TextBox.Enabled = True
        Else
            Spec_Seismic_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_SeismicSensor_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_SeismicSensor_Only_CheckBox.CheckedChanged
        If Spec_SeismicSensor_Only_CheckBox.Checked Then
            Spec_SeismicSensor_Only_TextBox.Enabled = True
        Else
            Spec_SeismicSensor_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_SeismicSW_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_SeismicSW_Only_CheckBox.CheckedChanged
        If Spec_SeismicSW_Only_CheckBox.Checked Then
            Spec_SeismicSW_Only_TextBox.Enabled = True
        Else
            Spec_SeismicSW_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_Parking_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_Parking_Only_CheckBox.CheckedChanged
        If Spec_Parking_Only_CheckBox.Checked Then
            Spec_Parking_Only_TextBox.Enabled = True
        Else
            Spec_Parking_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_CarGong_TopBtm_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_TopBtm_Only_CheckBox.CheckedChanged
        If Spec_CarGong_TopBtm_Only_CheckBox.Checked Then
            Spec_CarGong_TopBtm_Only_TextBox.Enabled = True
        Else
            Spec_CarGong_TopBtm_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_CarGong_Top_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_Top_Only_CheckBox.CheckedChanged
        If Spec_CarGong_Top_Only_CheckBox.Checked Then
            Spec_CarGong_Top_Only_TextBox.Enabled = True
        Else
            Spec_CarGong_Top_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_CarGong_COB_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_COB_Only_CheckBox.CheckedChanged
        If Spec_CarGong_COB_Only_CheckBox.Checked Then
            Spec_CarGong_COB_Only_TextBox.Enabled = True
        Else
            Spec_CarGong_COB_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_CarGong_VONIC_Only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_CarGong_VONIC_Only_CheckBox.CheckedChanged
        If Spec_CarGong_VONIC_Only_CheckBox.Checked Then
            Spec_CarGong_VONIC_Only_TextBox.Enabled = True
        Else
            Spec_CarGong_VONIC_Only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_WSCOB_only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_WSCOB_only_CheckBox.CheckedChanged
        If Spec_WSCOB_only_CheckBox.Checked Then
            Spec_WSCOB_only_TextBox.Enabled = True
        Else
            Spec_WSCOB_only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Spec_WCOB_only_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Spec_WCOB_only_CheckBox.CheckedChanged
        If Spec_WCOB_only_CheckBox.Checked Then
            Spec_WCOB_only_TextBox.Enabled = True
        Else
            Spec_WCOB_only_TextBox.Enabled = False
        End If
    End Sub

    Private Sub Label92_Click(sender As Object, e As EventArgs) Handles Label92.Click

    End Sub




    '----------------------------------------------------------------------------------------------------------------------- 其他事件 
End Class