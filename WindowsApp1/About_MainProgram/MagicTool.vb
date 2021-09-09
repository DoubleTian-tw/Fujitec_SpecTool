'Imports System.Diagnostics
'imports System.Windows.Forms.SystemInformation
Imports System.Windows.Forms.Application
Imports System.Text
Imports System.Timers
Imports System.IO
Imports System.IO.Directory

Module user32dll_use
    ''' <summary>
    ''' 取得windwo鍵盤按鍵功能
    ''' </summary>
    ''' <param name="vkey"></param>
    ''' <returns></returns>
    Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vkey As Integer) As Integer
End Module

Public Class MagicTool

    Dim sKeyValue, sKeyValue2 As New StringBuilder(512)
    Dim nSize As UInt32 = Convert.ToUInt32(1024)
    Dim sinifilename As String = StartupPath & "\ini\SetFile.ini"
    Dim note_dat As String = StartupPath & "\dat\Note.dat"

    Dim chalink As ChangeLink = New ChangeLink()
    Dim ProAllPath As ProgramAllPath = New ProgramAllPath()
    'Dim Form_Icon As String = Application.StartupPath & "\Pokeball_Icon.ico"
    Dim ballPathOri As String = StartupPath & "\ico\ball_1.ico"
    Dim ballPath As String = StartupPath & "\ico\ballC_2.ico"

    ''' <summary>
    ''' 提示ICON顯示之秒數
    ''' </summary>
    Const NotifyIconTimer As Integer = 10

    ''' <summary>
    ''' 目前子資料夾的總數
    ''' </summary>
    Public Const childForlder_sum As Integer = 6

    ''' <summary>
    ''' [MagicTool > 日曆 > 每日記事 檔案名稱]
    ''' </summary>
    Dim selectDateName_toDat As String

    ''' <summary>
    ''' [MagicTool > 日曆 > 每日記事 路徑]
    ''' </summary>
    Dim New_noteDat_path As String

    ''' <summary>
    ''' 確認path name combobox是否已有值
    ''' </summary>
    Public FolderPath_Name_Bool As Boolean = False
    'Public ComJob_Bool As Boolean = False '確認path name combobox是否已有值

    'Dim dateNote_dat_path As String
    ''' <summary>
    ''' [MagicTool > 日曆 > 全部日曆Dat檔案]
    ''' </summary>
    Dim datFiles As Object
    'Dim myFolderOpenPath_Dialog As New FolderBrowserDialog() '資料夾建立的對話視窗

    ''' <summary>
    ''' 主程式資料夾 名稱
    ''' </summary>
    'Dim mainProgramFileName As String = "Tool update folder"

    ''' <summary>
    ''' 主程式【DAT日記】資料夾 名稱
    ''' </summary>
    'Dim mainProgram_dat_FileName As String = "dat"

    ''' <summary>
    ''' 主要路徑
    ''' </summary>
    'Dim main_path As String = $"M:\DESIGN\BACK UP\yc_tian\Tool Application"

    ''' <summary>
    ''' 主要路徑離職後備份處
    ''' </summary>
    'Dim main_backupPath As String = $"M:\DESIGN BACKUP(LEAVE)\yc_tian\Tool Application"

    ''' <summary>
    ''' 主程式資料夾 路徑
    ''' </summary>
    'Dim checkNew_path As String = $"{main_path}\{mainProgramFileName}"
    'Dim checkNew_path As String = $"\\Yc-tian\共用文件夾\software\{mainProgramFileName}"

    ''' <summary>
    ''' 主程式 名稱
    ''' </summary>
    Dim form_name As String = "PokemonGOGO"

    'Const update_app_name As String = "Update_magicTool"
    ''' <summary>
    ''' 主程式【更新】資料夾 路徑
    ''' </summary>
    Dim updateTool_path As String '= $"{main_path}\{mainProgramFileName}\更新"
    'Dim updateTool_path As String = $"\\Yc-tian\共用文件夾\software\{mainProgramFileName}\更新" 'Application.StartupPath

    Dim chkNewVer_MainProgram As CheckNewVersion
    Dim chkNewVer_UpdateProgram As CheckNewVersion

    'Dim chkNewVer_Win As New CheckNewVersion($"{main_path}\{mainProgramFileName}\更新\ToolVersion.txt",
    '                                         "CheckNewVersion_WindowsApp1") '建立CheckNewVersion類別
    'Dim chkNewVer_Up As New CheckNewVersion($"{main_path}\{mainProgramFileName}\更新\ToolVersion.txt",
    '                                        "CheckNewVersion_Update_magicTool") '建立CheckNewVersion類別
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim thisAppVersion As FileVersionInfo
        thisAppVersion = FileVersionInfo.GetVersionInfo($"{StartupPath}\{ProgramAllName.fileName_mainProgram}.exe")
        MsgBox($"{CStr(thisAppVersion.FileVersion)},{CStr(thisAppVersion.FileMajorPart)}.{CStr(thisAppVersion.FileMinorPart)}.{CStr(thisAppVersion.FileBuildPart)}.")

        Dim specificAppVersion As FileVersionInfo
        specificAppVersion = FileVersionInfo.GetVersionInfo($"M:\DESIGN\BACK UP\yc_tian\Tool Application\Tool update folder\backup\ver109\{ProgramAllName.fileName_mainProgram}.exe")
        MsgBox($"{CStr(specificAppVersion.FileVersion)},{CStr(specificAppVersion.FileMajorPart)}.{CStr(specificAppVersion.FileMinorPart)}.{CStr(specificAppVersion.FileBuildPart)}.")
    End Sub

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            updateTool_path =
                $"{ProgramAllPath.folderName_mainProgram}\{ProgramAllPath.folderName_update}\{ProgramAllPath.folderName_updateChinese}"
            loadIni_form_changLink()
            'Form size
            Me.Width = 460
            Me.Height = 410
            LinkGroup1_SplitContainer.SplitterDistance = 130
            LinkGroup2_SplitContainer.SplitterDistance = 130
            LinkGroup3_SplitContainer.SplitterDistance = 130
            '元件小簡介
            ToolTip1.SetToolTip(Me.note_TextBox, "(๑•̀ω•́)ノ 你可以在這記事情喔喔喔!!!")
            ToolTip1.SetToolTip(Me.note_DateTimePicker, "(｢･ω･)｢　這兒切換筆記本日期")
            ToolTip1.SetToolTip(Me.MagicToll_MenuStrip, "ちゅ─=≡Σ((( つ•̀ω•́)つ")
            ToolTip1.SetToolTip(Me.MagicTool_TabControl, "きらきら(๑•̀ㅂ•́)و✧")
            'fun
            JustForFun_SplitContainer.SplitterDistance = 0
            JustForFun_SplitContainer.SplitterWidth = 11

            '按鈕初始
            btnUI_INI()

            '檢查更新 ---------------------------------------------------------

            Dim thisApp_Version As FileVersionInfo '執行端版本 this main program version，例如:1.2.3
            Dim thisAppVer_First, thisAppVer_Second, thisAppVer_Third As Integer
            Dim thisAppVer_Array As Integer()

            thisApp_Version =
                FileVersionInfo.GetVersionInfo($"{StartupPath}\{ProgramAllName.fileName_mainProgram}.exe")

            thisAppVer_First = thisApp_Version.FileMajorPart  '1.2.3取得版本的1
            thisAppVer_Second = thisApp_Version.FileMinorPart '1.2.3取得版本的2
            thisAppVer_Third = thisApp_Version.FileBuildPart  '1.2.3取得版本的3
            thisAppVer_Array = {thisAppVer_First, thisAppVer_Second, thisAppVer_Third}

            Dim pathApp_Version As FileVersionInfo '更新端路徑的版本
            Dim pathAppVer_First, pathAppVer_Second, pathAppVer_Third As Integer
            Dim pathAppVer_Array As Integer()
            pathApp_Version =
                FileVersionInfo.GetVersionInfo($"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\backup\ver109\{ProgramAllName.fileName_mainProgram}.exe")
            'FileVersionInfo.GetVersionInfo($"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\{ProgramAllPath.folderName_updateChinese}\{ProgramAllName.fileName_mainProgram}.exe")

            pathAppVer_First = pathApp_Version.FileMajorPart  '1.2.3取得版本的1
            pathAppVer_Second = pathApp_Version.FileMinorPart '1.2.3取得版本的2
            pathAppVer_Third = pathApp_Version.FileBuildPart  '1.2.3取得版本的3
            pathAppVer_Array = {pathAppVer_First, pathAppVer_Second, pathAppVer_Third}


            'If thisAppVer_First < pathAppVer_First Then
            '    MsgBox("update first")
            'ElseIf thisAppVer_First = pathAppVer_First Then
            '    If thisAppVer_Second < pathAppVer_Second Then
            '        MsgBox("update second")
            '    ElseIf thisAppVer_Second = pathAppVer_Second Then
            '        If thisAppVer_Third < pathAppVer_Third Then
            '            MsgBox("update third")
            '        End If
            '    End If
            'End If


            chkNewVer_MainProgram = New CheckNewVersion($"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\{ProgramAllPath.folderName_updateChinese}\ToolVersion.txt",
                                                $"CheckNewVersion_{ProgramAllName.fileName_mainProgram}") '建立CheckNewVersion類別
            chkNewVer_UpdateProgram = New CheckNewVersion($"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\{ProgramAllPath.folderName_updateChinese}\ToolVersion.txt",
                                                $"CheckNewVersion_{ProgramAllName.fileName_updateProgram}") '建立CheckNewVersion類別
            chkNewVer_MainProgram.CheckNewVersion()
            chkNewVer_UpdateProgram.CheckNewVersion_Up()
            Select Case chkNewVer_MainProgram.GetCheckConsequence '取得更新結果
                Case 0 'nothing
                    Me.Text = form_name & "目前為最新版本:ver." & chkNewVer_MainProgram.GetMyVersion
                Case 1 '有更新
                    Me.Text = form_name & "目前為舊版本號碼:ver." & chkNewVer_MainProgram.GetMyVersion
                    Dim result As MsgBoxResult
                    result = MsgBox($"有更新版本! 最新版本為:ver.{chkNewVer_MainProgram.GetCheckConsequenceNumber}{vbCrLf}是否自動更新?", vbYesNo, "更新訊息")

                    If result = MsgBoxResult.Yes Then
                        Select Case chkNewVer_UpdateProgram.GetCheckConsequence_up '取得更新結果
                            Case 0 'nothing
                                MsgBox($"{ProgramAllName.fileName_updateProgram}目前為最新版本:ver.{chkNewVer_UpdateProgram.GetMyVersion_up}")
                            Case 1 '有更新
                                MsgBox($"{ProgramAllName.fileName_updateProgram}目前為舊版本:ver.{chkNewVer_UpdateProgram.GetMyVersion_up}{vbLf}直接更新")
                                For Each myFile In Directory.GetFileSystemEntries(updateTool_path) '更新資料夾
                                    If Dir(myFile, vbDirectory) = $"{ProgramAllName.fileName_updateProgram}.exe" Then
                                        FileCopy(myFile, StartupPath & "\" & Path.GetFileName(myFile))
                                    End If
                                Next

                            Case 2 '更新失敗
                                Me.Text = form_name & "更新失敗"
                                MsgBox("更新失敗", MsgBoxStyle.Critical)
                        End Select

                        Dim update_p() As Process

                        Using p As Process = New Process()
                            p.Start($"{StartupPath}\{ProgramAllName.fileName_updateProgram}.exe")
                        End Using
                        update_p = Process.GetProcessesByName($"{ProgramAllName.fileName_updateProgram}")

                        If update_p.Count > 0 Then
                            Me.Close()
                        End If

                    End If


                Case 2 '更新失敗
                    Me.Text = form_name & "更新失敗"
                    MsgBox("更新失敗", MsgBoxStyle.Critical)
            End Select


            '--------------------------------------------------------- 檢查更新 


            '設定DAT檔案的名稱
            selectDateName_toDat =
                $"Note_{note_DateTimePicker.Value.Year}.{note_DateTimePicker.Value.Month}.{note_DateTimePicker.Value.Day}"

            New_noteDat_path = $"{StartupPath}\{ProgramAllPath.folderName_dat}\{selectDateName_toDat}.dat" '日期筆記
            Try
                note_TextBox.Text = IO.File.ReadAllText(note_dat) '一般筆記
            Catch ex As Exception
                'MsgBox("Note.dat > 資料遺失/路徑不正確" & vbCrLf & "請移至下列路徑 : " & note_dat.ToString,, "dat檔案遺失")
                MsgBox($"Note.dat > 資料遺失/路徑不正確{vbCrLf}請移至下列路徑 : {note_dat.ToString} , dat檔案遺失")
                Process.Start($"{StartupPath}\{ProgramAllPath.folderName_dat}")
            End Try
            If IO.File.Exists(New_noteDat_path) Then
                DateNote_TextBox.Text = IO.File.ReadAllText(New_noteDat_path)
            Else
                DateNote_TextBox.Text = note_DateTimePicker.Value.ToShortDateString
            End If


            '檢查文件更新
            ChangeLink.fileUpdateNotice_check()
        Catch ea As Exception
            MsgBox($"ChangLick.Form1_load，訊息{ea.ToString}")
        End Try
    End Sub


    Private Sub MagicTool_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'read
        'LoadIni()

        If IO.File.Exists(note_dat) Then
            note_TextBox.Text = IO.File.ReadAllText(note_dat)
        End If

        If IO.File.Exists(New_noteDat_path) Then
            DateNote_TextBox.Text = IO.File.ReadAllText(New_noteDat_path)
        Else
            DateNote_TextBox.Text = note_DateTimePicker.Value.ToShortDateString
        End If


        'hotkey timer_tick()
        Me.KeyPreview = True
        Timer1.Enabled = True
        Timer1.Interval = 1

    End Sub

    Public Sub New()

        ' 設計工具需要此呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        MagicTool_NotifyIcon.ContextMenuStrip = MagicToll_MenuStrip.ContextMenuStrip
        MagicTool_NotifyIcon.Text = Me.Text

        datFile_load()

    End Sub

    Private Sub datFile_load()
        '取得DAT檔案的數量以及名稱
        DelDateNote_ToolStrip.DropDownItems.Clear()

        datFiles = From chkDAT In Directory.EnumerateFiles(StartupPath & "\dat", "*.dat", SearchOption.AllDirectories)
                   Select New With {.curFile = chkDAT}
        For Each df In datFiles
            DelDateNote_ToolStrip.DropDownItems.Add(df.curFile)
        Next
    End Sub
    Public Sub loadIni_form_changLink()
        Try
            chalink.Initialization_ini()
            chalink.formPositionOnScreen_Setting(Me, chalink.sKeyValueScr.ToString, chalink.sKeyValuePos.ToString)
            chalink.Topmost_setting(Me, False)
            chalink.LinkCB_setting()

            '顏色INI
            Dim sKey_setColor As New StringBuilder(512)
            GetPrivateProfileString("SettingColor", "SetLinkBtn_Change", "", sKey_setColor, nSize, sinifilename)
            If sKey_setColor.ToString = CStr(True) Then
                btnUI_INI()
            End If
        Catch e As Exception
            MsgBox($"ChangLick.LoadIni，訊息{e.ToString}")
        End Try
    End Sub


    '點快捷錯誤時判斷
    Private Sub LinkButton_error(dir_text As String)
        Try
            Process.Start(dir_text)
        Catch ex As Exception
            MsgBox("沒有指定目錄/目錄錯誤",, "路徑錯誤")
        End Try
    End Sub
    Private Sub Manual_ToolStrip_Click(sender As Object, e As EventArgs) Handles Manual_ToolStrip.Click
        'LinkButton_error(StartupPath & "\ppt\Manual.pptx")
        LinkButton_error($"{StartupPath}\{ProgramAllPath.folderName_ppt}\{ProgramAllName.fileName_Manualpptx}")
    End Sub
    Private Sub Link1_Button_Click(sender As Object, e As EventArgs) Handles Link1_1_Button.Click
        LinkButton_error(chalink.Link1_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_Button_Click(sender As Object, e As EventArgs) Handles Link1_2_Button.Click
        LinkButton_error(chalink.Link2_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_Button_Click(sender As Object, e As EventArgs) Handles Link1_3_Button.Click
        LinkButton_error(chalink.Link3_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_Button_Click(sender As Object, e As EventArgs) Handles Link1_4_Button.Click
        LinkButton_error(chalink.Link4_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_Button_Click(sender As Object, e As EventArgs) Handles Link1_5_Button.Click
        LinkButton_error(chalink.Link5_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_Button_Click(sender As Object, e As EventArgs) Handles Link1_6_Button.Click
        LinkButton_error(chalink.Link6_Dir_TextBox.Text)
    End Sub
    Private Sub Link7_Button_Click(sender As Object, e As EventArgs) Handles Link1_7_Button.Click
        LinkButton_error(chalink.Link7_Dir_TextBox.Text)
    End Sub
    Private Sub Link8_Button_Click(sender As Object, e As EventArgs) Handles Link1_8_Button.Click
        LinkButton_error(chalink.Link8_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_1_Button_Click(sender As Object, e As EventArgs) Handles Link2_1_Button.Click
        LinkButton_error(chalink.Link2_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_2_Button_Click(sender As Object, e As EventArgs) Handles Link2_2_Button.Click
        LinkButton_error(chalink.Link2_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_3_Button_Click(sender As Object, e As EventArgs) Handles Link2_3_Button.Click
        LinkButton_error(chalink.Link2_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_4_Button_Click(sender As Object, e As EventArgs) Handles Link2_4_Button.Click
        LinkButton_error(chalink.Link2_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_5_Button_Click(sender As Object, e As EventArgs) Handles Link2_5_Button.Click
        LinkButton_error(chalink.Link2_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_6_Button_Click(sender As Object, e As EventArgs) Handles Link2_6_Button.Click
        LinkButton_error(chalink.Link2_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_7_Button_Click(sender As Object, e As EventArgs) Handles Link2_7_Button.Click
        LinkButton_error(chalink.Link2_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_8_Button_Click(sender As Object, e As EventArgs) Handles Link2_8_Button.Click
        LinkButton_error(chalink.Link2_8_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_1_Button_Click(sender As Object, e As EventArgs) Handles Link3_1_Button.Click
        LinkButton_error(chalink.Link3_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_2_Button_Click(sender As Object, e As EventArgs) Handles Link3_2_Button.Click
        LinkButton_error(chalink.Link3_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_3_Button_Click(sender As Object, e As EventArgs) Handles Link3_3_Button.Click
        LinkButton_error(chalink.Link3_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_4_Button_Click(sender As Object, e As EventArgs) Handles Link3_4_Button.Click
        LinkButton_error(chalink.Link3_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_5_Button_Click(sender As Object, e As EventArgs) Handles Link3_5_Button.Click
        LinkButton_error(chalink.Link3_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_6_Button_Click(sender As Object, e As EventArgs) Handles Link3_6_Button.Click
        LinkButton_error(chalink.Link3_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_7_Button_Click(sender As Object, e As EventArgs) Handles Link3_7_Button.Click
        LinkButton_error(chalink.Link3_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_8_Button_Click(sender As Object, e As EventArgs) Handles Link3_8_Button.Click
        LinkButton_error(chalink.Link3_8_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_1_Button_Click(sender As Object, e As EventArgs) Handles Link4_1_Button.Click
        LinkButton_error(chalink.Link4_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_2_Button_Click(sender As Object, e As EventArgs) Handles Link4_2_Button.Click
        LinkButton_error(chalink.Link4_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_3_Button_Click(sender As Object, e As EventArgs) Handles Link4_3_Button.Click
        LinkButton_error(chalink.Link4_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_4_Button_Click(sender As Object, e As EventArgs) Handles Link4_4_Button.Click
        LinkButton_error(chalink.Link4_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_5_Button_Click(sender As Object, e As EventArgs) Handles Link4_5_Button.Click
        LinkButton_error(chalink.Link4_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_6_Button_Click(sender As Object, e As EventArgs) Handles Link4_6_Button.Click
        LinkButton_error(chalink.Link4_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_7_Button_Click(sender As Object, e As EventArgs) Handles Link4_7_Button.Click
        LinkButton_error(chalink.Link4_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_8_Button_Click(sender As Object, e As EventArgs) Handles Link4_8_Button.Click
        LinkButton_error(chalink.Link4_8_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_1_Button_Click(sender As Object, e As EventArgs) Handles Link5_1_Button.Click
        LinkButton_error(chalink.Link5_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_2_Button_Click(sender As Object, e As EventArgs) Handles Link5_2_Button.Click
        LinkButton_error(chalink.Link5_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_3_Button_Click(sender As Object, e As EventArgs) Handles Link5_3_Button.Click
        LinkButton_error(chalink.Link5_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_4_Button_Click(sender As Object, e As EventArgs) Handles Link5_4_Button.Click
        LinkButton_error(chalink.Link5_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_5_Button_Click(sender As Object, e As EventArgs) Handles Link5_5_Button.Click
        LinkButton_error(chalink.Link5_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_6_Button_Click(sender As Object, e As EventArgs) Handles Link5_6_Button.Click
        LinkButton_error(chalink.Link5_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_7_Button_Click(sender As Object, e As EventArgs) Handles Link5_7_Button.Click
        LinkButton_error(chalink.Link5_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_8_Button_Click(sender As Object, e As EventArgs) Handles Link5_8_Button.Click
        LinkButton_error(chalink.Link5_8_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_1_Button_Click(sender As Object, e As EventArgs) Handles Link6_1_Button.Click
        LinkButton_error(chalink.Link6_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_2_Button_Click(sender As Object, e As EventArgs) Handles Link6_2_Button.Click
        LinkButton_error(chalink.Link6_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_3_Button_Click(sender As Object, e As EventArgs) Handles Link6_3_Button.Click
        LinkButton_error(chalink.Link6_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_4_Button_Click(sender As Object, e As EventArgs) Handles Link6_4_Button.Click
        LinkButton_error(chalink.Link6_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_5_Button_Click(sender As Object, e As EventArgs) Handles Link6_5_Button.Click
        LinkButton_error(chalink.Link6_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_6_Button_Click(sender As Object, e As EventArgs) Handles Link6_6_Button.Click
        LinkButton_error(chalink.Link6_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_7_Button_Click(sender As Object, e As EventArgs) Handles Link6_7_Button.Click
        LinkButton_error(chalink.Link6_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_8_Button_Click(sender As Object, e As EventArgs) Handles Link6_8_Button.Click
        LinkButton_error(chalink.Link6_8_Dir_TextBox.Text)
    End Sub

    '點快捷錯誤時判斷

    Private Sub NotifyIcon_ToolStrip_Click(sender As Object, e As EventArgs) Handles NotifyIcon_ToolStrip.Click
        '關於小工具
        MagicTool_NotifyIcon.BalloonTipTitle = "可以在這使用寶貝球"
        MagicTool_NotifyIcon.BalloonTipText = "點我兩下開啟/收回"
        MagicTool_NotifyIcon.BalloonTipIcon = ToolTipIcon.Info
        MagicTool_NotifyIcon.ShowBalloonTip(NotifyIconTimer)
    End Sub

    Private Sub About_ToolStrip_Click(sender As Object, e As EventArgs) Handles About_ToolStrip.Click
        About.Show()
        chalink.LinkCB_setting()
    End Sub

    Private Sub BasicSetting_ToolStrip_Click(sender As Object, e As EventArgs) Handles BasicSetting_ToolStrip.Click
        'Me.Hide()
        ChangeLink.Show()
    End Sub

    Private Sub MagicTool_NotifyIcon_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles MagicTool_NotifyIcon.MouseDoubleClick

        '點兩下icon打開/縮小視窗
        If Me.WindowState = FormWindowState.Minimized Then
            Me.WindowState = FormWindowState.Normal
            MagicTool_NotifyIcon.Icon = New Icon(ballPath)
        Else
            Me.WindowState = FormWindowState.Minimized
            MagicTool_NotifyIcon.Icon = New Icon(ballPathOri)
        End If

    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim ctrlKey As Boolean
        Dim QKey As Boolean
        ctrlKey = GetAsyncKeyState(Keys.ControlKey)
        QKey = GetAsyncKeyState(Keys.Q)

        If ctrlKey And QKey = True Then
            'Me.Show()
            Me.WindowState = FormWindowState.Normal

            Me.TopMost = True
            If chalink.Topmost_CheckBox.CheckState = CheckState.Checked Then
                'Me.TopMost = True
            Else
                Me.TopMost = False
            End If
        End If
    End Sub


    Private Sub LinkGroup1_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LinkGroup1_ComboBox.SelectedIndexChanged
        If LinkGroup1_ComboBox.SelectedIndex = 0 Then
            LinkGroup1_SplitContainer.SplitterDistance = 130
        ElseIf LinkGroup1_ComboBox.SelectedIndex = 1 Then
            LinkGroup1_SplitContainer.SplitterDistance = 0
        End If
    End Sub
    Private Sub LinkGroup2_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LinkGroup2_ComboBox.SelectedIndexChanged
        If LinkGroup2_ComboBox.SelectedIndex = 0 Then
            LinkGroup2_SplitContainer.SplitterDistance = 130
        ElseIf LinkGroup2_ComboBox.SelectedIndex = 1 Then
            LinkGroup2_SplitContainer.SplitterDistance = 0
        End If
    End Sub
    Private Sub LinkGroup3_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LinkGroup3_ComboBox.SelectedIndexChanged
        If LinkGroup3_ComboBox.SelectedIndex = 0 Then
            LinkGroup3_SplitContainer.SplitterDistance = 130
        ElseIf LinkGroup3_ComboBox.SelectedIndex = 1 Then
            LinkGroup3_SplitContainer.SplitterDistance = 0
        End If
    End Sub

    Private Sub JustForFun_SplitContainer_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles JustForFun_SplitContainer.SplitterMoved
        If JustForFun_SplitContainer.SplitterDistance < 155 And JustForFun_SplitContainer.SplitterDistance > 5 Then
            ToolTip1.SetToolTip(Me.JustForFun_SplitContainer, "(๑•̀ω•́)ノ 想偷看阿!")
        ElseIf JustForFun_SplitContainer.SplitterDistance > 235 And JustForFun_SplitContainer.SplitterDistance < 350 Then
            ToolTip1.SetToolTip(Me.JustForFun_SplitContainer, "(ΦωΦ)喵~")
        Else
            ToolTip1.SetToolTip(Me.JustForFun_SplitContainer, "唉丫~被你發現了甚麼")
        End If
    End Sub

    Private Sub BackOriPos_ToolStrip_Click(sender As Object, e As EventArgs) Handles BackOriPos_ToolStrip.Click
        chalink.formPositionOnScreen_Setting(Me, chalink.sKeyValueScr.ToString, chalink.sKeyValuePos.ToString)
    End Sub

    Private Sub note_TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles note_TextBox.KeyPress

        If IO.File.Exists(note_dat) Then
            File.WriteAllText(note_dat, note_TextBox.Text)
        Else
            MsgBox("寫入之檔案不存在，請重新導入" & vbCrLf & "請將Note.dat檔案移置\dat資料夾底下",, "dat檔案遺失")
            Process.Start($"{StartupPath}\{ProgramAllPath.folderName_dat}")
        End If

    End Sub
    Private Sub DateNote_TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DateNote_TextBox.KeyPress
        IO.File.WriteAllText(New_noteDat_path, DateNote_TextBox.Text)
    End Sub

    Private Sub AboutDelDateNote_ToolStri_Click(sender As Object, e As EventArgs) Handles AboutDelDateNote_ToolStri.Click
        MsgBox("~刪除<日記事項>~" & vbCrLf & "選項內有您建立的資料，如果有不需要之日記，可以來這刪除",, "關於刪除這小事")
    End Sub

    Private Sub note_DateTimePicker_ValueChanged(sender As Object, e As EventArgs) Handles note_DateTimePicker.ValueChanged
        selectDateName_toDat = "Note_" & note_DateTimePicker.Value.Year & "." & note_DateTimePicker.Value.Month & "." & note_DateTimePicker.Value.Day
        'New_noteDat_path = StartupPath & "\dat\" & selectDateName_toDat & ".dat" '日期筆記
        New_noteDat_path = $"{StartupPath}\{ProgramAllPath.folderName_dat}\{selectDateName_toDat}.{ProgramAllPath.folderName_dat}" '日期筆記
        If IO.File.Exists(New_noteDat_path) Then
            DateNote_TextBox.Text = IO.File.ReadAllText(New_noteDat_path)
        Else
            DateNote_TextBox.Text = note_DateTimePicker.Value.ToShortDateString
        End If
    End Sub
    '創新資料夾
    Private Sub FolderOpenPath_Button_Click(sender As Object, e As EventArgs) Handles FolderOpenPath_Button.Click '創建新資料夾的路徑
        ChangeLink.OpenFilePath_event(FolderPath_ComboBox)
    End Sub

    Private Sub FileChoUse_Button_Click(sender As Object, e As EventArgs) Handles FileChoUse_Button.Click 'copy常用工番檔案的路徑
        ChangeLink.OpenFilePath_event(FileChoUse_ComboBox)
    End Sub

    Private Sub FileChoose_Button_Click(sender As Object, e As EventArgs) Handles FileChoose_Button.Click

        FileUse_CheckedListBox.Items.Clear()
        For Each i As Object In FileChoose_CheckedListBox.Items
            If FileChoose_CheckedListBox.GetItemChecked(FileChoose_CheckedListBox.Items.IndexOf(i).ToString) Then
                FileUse_CheckedListBox.Items.Add(i, CheckState.Checked)
            End If
        Next i

    End Sub

    Private Sub CreateFolder_Button_Click(sender As Object, e As EventArgs) Handles CreateFolder_Button.Click
        Dim fatherPath_to_ChildPAth As String '父資料夾的路徑
        Dim sourceDir, PasteDir As String
        'Dim source_fileList As New ArrayList(1)

        fatherPath_to_ChildPAth = FolderPath_ComboBox.Text & "\" & FolderName_ComboBox.Text
        sourceDir = FileChoUse_ComboBox.Text
        PasteDir = fatherPath_to_ChildPAth

        If FolderPath_ComboBox.Text <> "" Then 'if folderpath is null do nothing

            '建立子父資料夾
            Directory.CreateDirectory(fatherPath_to_ChildPAth) '父資料夾

            Dim MA_ChildFolder_Group As TextBox() =
                {MAchildFolder_TextBox1, MAchildFolder_TextBox2, MAchildFolder_TextBox3,
                 MAchildFolder_TextBox4, MAchildFolder_TextBox5, MAchildFolder_TextBox6}
            Dim MA_ChildFolder_Checkbox_Group As CheckBox() =
                {childFolder_CheckBox1, childFolder_CheckBox2, childFolder_CheckBox3,
                childFolder_CheckBox4, childFolder_CheckBox5, childFolder_CheckBox6}
            For i = 0 To childForlder_sum - 1 'create child folder
                If MA_ChildFolder_Checkbox_Group(i).Checked = True Then
                    'Directory.CreateDirectory(fatherPath_to_ChildPAth & "\" & MA_ChildFolder_Group(i).Text)
                    Directory.CreateDirectory($"{fatherPath_to_ChildPAth}\{MA_ChildFolder_Group(i).Text}")
                End If
            Next i

            '常用工番檔案新增
            If FileChoUse_ComboBox.Text <> "" Then
                For Each fileUse In FileUse_CheckedListBox.Items
                    Try
                        File.Copy(Path.Combine(sourceDir, fileUse), Path.Combine(PasteDir, fileUse)) 'copy前方放來源，後方放目的
                    Catch ex As Exception
                        '相同名稱忽略錯誤，繼續執行
                    End Try

                Next
            End If

            'open folder打開資料夾
            LinkButton_error(fatherPath_to_ChildPAth)


        Else
            MsgBox("沒有選擇路徑")
        End If
    End Sub

    Private Sub FileChoUse_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FileChoUse_ComboBox.SelectedIndexChanged
        '常用工番路徑的檔案

        FileChoose_CheckedListBox.Items.Clear() '淨空

        Dim file_Len, allfile_len As Integer
        Dim filter_name() As String
        Dim fileChoUse_path As String
        'app_Len = Len(Application.StartupPath)

        filter_name = {"*.doc", "*.xls"}
        'appStart_path = Application.StartupPath
        fileChoUse_path = FileChoUse_ComboBox.Text

        Try
            For Each myFilter In filter_name
                For Each file In GetFileSystemEntries(fileChoUse_path, myFilter)
                    file_Len = Len(fileChoUse_path)
                    allfile_len = Len(file)
                    FileChoose_CheckedListBox.Items.Add(Strings.Right(file, allfile_len - (file_Len + 1)))
                Next
            Next
        Catch ex As Exception
            MsgBox("指定常用工番路徑已刪除變動，系統找不到相對應資料夾", vbCritical, "ERROR常用工番路徑ERROR")
            'Process.Start(appStart_path & "\jobfile")
        End Try

    End Sub

    Private Sub childAllFolder_ComboBox_Click(sender As Object, e As EventArgs) Handles childAllFolder_ComboBox.Click
        Dim childCheck As CheckBox()
        Dim childTextbox As TextBox()
        Dim count_i As Integer
        count_i = 0
        childCheck = {childFolder_CheckBox1, childFolder_CheckBox2, childFolder_CheckBox3, childFolder_CheckBox4, childFolder_CheckBox5, childFolder_CheckBox6}
        childTextbox = {MAchildFolder_TextBox1, MAchildFolder_TextBox2, MAchildFolder_TextBox3, MAchildFolder_TextBox4, MAchildFolder_TextBox5, MAchildFolder_TextBox6}

        childAllFolder_ComboBox.Items.Clear()
        For Each i In childCheck
            If i.Checked = True Then
                childAllFolder_ComboBox.Items.Add(childTextbox(count_i).Text)
                count_i = count_i + 1
            End If
        Next

    End Sub

    Private Sub JobMaker_Button_Click(sender As Object, e As EventArgs) Handles JobMaker_Button.Click
        JobMaker_Form.Show()
    End Sub

    Private Sub CleanAll_Button_Click(sender As Object, e As EventArgs) Handles CleanAll_Button.Click
        '按此按鈕可清除所有<創建資料夾>的表格內容
        FolderPath_ComboBox.Text = ""
        FolderName_ComboBox.Text = ""
        FileChoUse_ComboBox.Text = ""
        FileChoose_CheckedListBox.Items.Clear()
        FileUse_CheckedListBox.Items.Clear()

        Dim LinkCheckBox As CheckBox() = {childFolder_CheckBox1, childFolder_CheckBox2, childFolder_CheckBox3, childFolder_CheckBox4, childFolder_CheckBox5 _
            , childFolder_CheckBox6}

        For Each i_box In LinkCheckBox
            If i_box.Checked = True Then
                i_box.Checked = False
            End If
        Next

    End Sub

    Private Sub Update_Button_Click(sender As Object, e As EventArgs) Handles Update_Button.Click
        Dim update_result As DialogResult = MessageBox.Show("是否檢查更新?", "即將更新...", MessageBoxButtons.YesNo)
        Dim update_p() As Process

        If update_result = DialogResult.Yes Then

            Select Case chkNewVer_MainProgram.GetCheckConsequence '取得更新結果
                Case 0 'nothing
                    MsgBox("magicTool目前為最新版本:ver." & chkNewVer_MainProgram.GetMyVersion & ",不更新", , "更新資訊")
                Case 1 '有更新
                    MsgBox("magicTool目前為舊版本:ver." & chkNewVer_MainProgram.GetMyVersion & vbLf & "直接更新", , "更新資訊")
                Case 2 '更新失敗
                    Me.Text = form_name & "更新失敗"
                    MsgBox("更新失敗", MsgBoxStyle.Critical)
            End Select

            Select Case chkNewVer_UpdateProgram.GetCheckConsequence_up '取得更新結果
                Case 0 'nothing
                    MsgBox($"{ProgramAllName.fileName_updateProgram}目前為最新版本:ver.{chkNewVer_UpdateProgram.GetMyVersion_up},不更新", , "更新資訊")
                Case 1 '有更新
                    MsgBox($"{ProgramAllName.fileName_updateProgram}目前為舊版本:ver.{chkNewVer_UpdateProgram.GetMyVersion_up}{vbLf}直接更新",, "更新資訊")
                    For Each myFile In Directory.GetFileSystemEntries(updateTool_path) '更新資料夾
                        If Dir(myFile, vbDirectory) = $"{ProgramAllName.fileName_updateProgram}.exe" Then
                            FileCopy(myFile, StartupPath & "\" & Path.GetFileName(myFile))
                        End If
                    Next
                    Process.Start($"{StartupPath}\{ProgramAllName.fileName_updateProgram}.exe")
                    update_p = Process.GetProcessesByName($"{ProgramAllName.fileName_updateProgram}")

                    If update_p.Count > 0 Then
                        Me.Close()
                    End If
                Case 2 '更新失敗
                    Me.Text = form_name & "更新失敗"
                    MsgBox("更新失敗", MsgBoxStyle.Critical, "更新資訊")
            End Select
        End If
    End Sub

    '創新資料夾
    Private Sub DelDateNote_ToolStrip_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles DelDateNote_ToolStrip.DropDownItemClicked
        'use DirectCast turn e.clicked to toolstrip
        Dim MenuItem As ToolStripMenuItem = DirectCast(e.ClickedItem, ToolStripMenuItem)
        del_or_open_DateNote_by_msgbox(dateNote_state.del_dateNote,
                                       MenuItem.ToString,
                                       Mid(CType(MenuItem.ToString, String),
                                           InStr(CType(MenuItem.ToString, String), "N"),
                                           Len(CType(MenuItem.ToString, String))))
        '重新讀取
        datFile_load()
    End Sub
    Enum dateNote_state
        del_dateNote
        open_dateNote
    End Enum
    Public Sub del_or_open_DateNote_by_msgbox(del_or_open As dateNote_state, link As String, name As String)
        Dim msgbxYN As DialogResult

        If del_or_open = dateNote_state.open_dateNote Then
            msgbxYN = MessageBox.Show($"確定打開{name}嗎?", "提醒", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Asterisk)

            If msgbxYN = vbYes Then
                Process.Start(link)
            End If
        ElseIf del_or_open = dateNote_state.del_dateNote Then
            msgbxYN = MessageBox.Show($"確定刪除{name}嗎?", "提醒", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Asterisk)

            If msgbxYN = vbYes Then
                My.Computer.FileSystem.DeleteFile(link)
            End If
        End If
    End Sub
    '定義連結按鈕的狀態
    Private Sub btnUI_INI()
        'Dim splCon_array() As SplitContainer
        'splCon_array = {LinkGroup1_SplitContainer, LinkGroup2_SplitContainer, LinkGroup3_SplitContainer}


        'For Each mSplCon As Control In LinkGroup1_SplitContainer.Controls
        '    For Each mFlow As Control In mSplCon.Controls
        '        For Each mBtn As Control In mFlow.Controls
        '            MsgBox(mBtn.Name)
        '        Next
        '    Next
        'Next
        Dim linkBtn_All() As Button =
            {Link1_1_Button, Link1_2_Button, Link1_3_Button, Link1_4_Button, Link1_5_Button, Link1_6_Button, Link1_7_Button, Link1_8_Button,
             Link2_1_Button, Link2_2_Button, Link2_3_Button, Link2_4_Button, Link2_5_Button, Link2_6_Button, Link2_7_Button, Link2_8_Button,
             Link3_1_Button, Link3_2_Button, Link3_3_Button, Link3_4_Button, Link3_5_Button, Link3_6_Button, Link3_7_Button, Link3_8_Button,
             Link4_1_Button, Link4_2_Button, Link4_3_Button, Link4_4_Button, Link4_5_Button, Link4_6_Button, Link4_7_Button, Link4_8_Button,
             Link5_1_Button, Link5_2_Button, Link5_3_Button, Link5_4_Button, Link5_5_Button, Link5_6_Button, Link5_7_Button, Link5_8_Button,
             Link6_1_Button, Link6_2_Button, Link6_3_Button, Link6_4_Button, Link6_5_Button, Link6_6_Button, Link6_7_Button, Link6_8_Button}
        Dim note_All() As TextBox = {note_TextBox, DateNote_TextBox}

        For Each myBtn In linkBtn_All
            linkBtnUI_state(myBtn)
        Next

        For Each myNote In note_All
            noteUI_state(myNote)
        Next

        If chalink.SetLinkBtn_BgPicture_TextBox.Text <> "" Then
            Links_Group.BackgroundImage = Image.FromFile(chalink.SetLinkBtn_BgPicture_TextBox.Text)
        End If
    End Sub

    Private Sub noteUI_state(note As TextBox)
        If chalink.SetNote_BackColor_Button.Text <> "" Then
            note.BackColor = ColorTranslator.FromHtml(chalink.SetNote_BackColor_Button.Text)
        Else
            note.BackColor = DefaultBackColor
        End If

        If chalink.SetNote_FontColor_Button.Text <> "" Then
            note.ForeColor = ColorTranslator.FromHtml(chalink.SetNote_FontColor_Button.Text)
        Else
            note.ForeColor = DefaultForeColor
        End If
    End Sub

    Private Sub noteAdd_Button_Click(sender As Object, e As EventArgs) Handles noteAdd_Button.Click
        Dim new_GroupBox As GroupBox = New GroupBox
        Dim new_textBox As TextBox = New TextBox
        Dim new_TimerCheckBox As CheckBox = New CheckBox
        Dim new_HourLabel As Label = New Label
        Dim new_TimerLabel As Label = New Label
        Dim new_XLabel As Label = New Label

        With new_GroupBox
            .Height = noteSample_GroupBox.Height
            .Width = noteSample_GroupBox.Width
            .Text = $"工作{note_FlowLayoutPanel.Controls.Count + 1}"
        End With


        'Dim ctrls_i As Integer
        'For Each ctrl As Control In note_FlowLayoutPanel.Controls
        '    ctrls_i += 1
        '    If TypeName(ctrl) = "GroupBox" Then
        '        ctrl.Name = $"note_GroupBox{ctrls_i}"
        '        ctrl.Top = ctrl.Height * ctrls_i
        '    End If
        '    MsgBox(ctrl.Name)
        'Next
        note_FlowLayoutPanel.Controls.Add(new_GroupBox)
    End Sub

    Private Sub note_FlowLayoutPanel_Paint(sender As Object, e As PaintEventArgs) Handles note_FlowLayoutPanel.Paint

    End Sub

    Private Sub linkBtnUI_state(btn As Button)
        btn.FlatStyle = Windows.Forms.FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 1
        btn.FlatAppearance.BorderColor = ColorTranslator.FromHtml(chalink.SetLinkBtn_BorderColor_Button.Text)
        btn.FlatAppearance.MouseDownBackColor = Color.Transparent
        btn.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(chalink.SetLinkBtn_MouseOverColor_Button.Text)
        If chalink.SetLinkBtn_Transparent_Button.Text = "YES" Then
            btn.BackColor = Color.Transparent
        ElseIf chalink.SetLinkBtn_Transparent_Button.Text = "NO" Then
            btn.BackColor = DefaultBackColor
        End If
        btn.ForeColor = ColorTranslator.FromHtml(chalink.SetLinkBtn_FontColor_Button.Text)
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub note_FlowLayoutPanel_MouseWheel(sender As Object, e As MouseEventArgs) Handles note_FlowLayoutPanel.MouseWheel
        MsgBox(My.Computer.Mouse.WheelScrollLines)
        'Dim numberOfTextLinesToMove As Integer = CInt(e.Delta * SystemInformation.MouseWheelScrollLines / 120)
        'Dim numberOfPixelsToMove As Integer = numberOfTextLinesToMove * 20
        'Dim mousePath As New Drawing.Drawing2D.GraphicsPath
        'If numberOfPixelsToMove <> 0 Then
        '    Dim translateMatrix As New System.Drawing.Drawing2D.Matrix()
        '    translateMatrix.Translate(0, numberOfPixelsToMove)
        '    mousePath.Transform(translateMatrix)
        'End If

        'note_FlowLayoutPanel.Invalidate()
    End Sub

    Private Sub Panel1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Panel1.MouseWheel
        MsgBox(My.Computer.Mouse.WheelScrollLines)
        Panel1.FindForm()
    End Sub

    Private Sub note_FlowLayoutPanel_Click(sender As Object, e As EventArgs) Handles note_FlowLayoutPanel.Click
        note_FlowLayoutPanel.FindForm()
    End Sub
    '定義連結按鈕的狀態
End Class



