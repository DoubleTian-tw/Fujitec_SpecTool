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
    ''' <summary>
    ''' [MagicTool > 日曆 > 全部日曆Dat檔案]
    ''' </summary>
    Dim datFiles As Object

    ''' <summary>
    ''' 主程式【更新】資料夾 路徑
    ''' </summary>
    Dim updateTool_path As String

    Dim chkNewVer_MainProgram As CheckNewVersion
    Dim chkNewVer_UpdateProgram As CheckNewVersion


    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            updateTool_path =
                $"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\{ProgramAllPath.folderName_updateChinese}"
            loadIni_form_changLink()
            'Form size
            Me.Width = 460
            Me.Height = 410
            LinkGroup1_SplitContainer.SplitterDistance = 130
            LinkGroup2_SplitContainer.SplitterDistance = 130
            LinkGroup3_SplitContainer.SplitterDistance = 130
            '元件小簡介
            JustForFun_ToolTip.SetToolTip(Me.note_TextBox, "(๑•̀ω•́)ノ 你可以在這記事情喔喔喔!!!")
            JustForFun_ToolTip.SetToolTip(Me.note_DateTimePicker, "(｢･ω･)｢　這兒切換筆記本日期")
            JustForFun_ToolTip.SetToolTip(Me.MagicToll_MenuStrip, "ちゅ─=≡Σ((( つ•̀ω•́)つ")
            JustForFun_ToolTip.SetToolTip(Me.MagicTool_TabControl, "きらきら(๑•̀ㅂ•́)و✧")
            'fun
            JustForFun_SplitContainer.SplitterDistance = 0
            JustForFun_SplitContainer.SplitterWidth = 11

            '按鈕初始
            btnUI_INI()

            '檢查更新 ---------------------------------------------------------
            check_File_Version(isUpdateButton:=False)


            '--------------------------------------------------------- 檢查更新 

            MagicTool_NotifyIcon.ContextMenuStrip = MagicToll_MenuStrip.ContextMenuStrip
            MagicTool_NotifyIcon.Text = Me.Text

            datFile_load()
            '設定DAT檔案的名稱
            selectDateName_toDat =
                $"Note_{note_DateTimePicker.Value.Year}.{note_DateTimePicker.Value.Month}.{note_DateTimePicker.Value.Day}"

            New_noteDat_path = $"{StartupPath}\{ProgramAllPath.folderName_dat}\{selectDateName_toDat}.dat" '日期筆記
            Try
                note_TextBox.Text = IO.File.ReadAllText(note_dat) '一般筆記
            Catch ex As Exception
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


    ''' <summary>
    ''' 回傳當前軟體的組件名稱
    ''' </summary>
    Dim thisApp_fullName As String = ProgramAllName.get_assemblyName
    ''' <summary>
    ''' 本機端-主程式-版本 this main program version，例如:1.2.3
    ''' </summary>
    Dim thisApp_Version As FileVersionInfo '=
    'FileVersionInfo.GetVersionInfo($"{StartupPath}\{thisApp_fullName}.exe")
    ''' <summary>
    ''' 更新端-主程式-版本 this main program version，例如:1.2.3
    ''' </summary>
    Dim updateApp_Version As FileVersionInfo '=
    'FileVersionInfo.GetVersionInfo($"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\更新\{thisApp_fullName}.exe") '執行端版本 this main program version，例如:1.2.3
    ''' <summary>
    ''' 本機端-更新主程式-版本 this main program version，例如:1.2.3
    ''' </summary>
    Dim thisUpdateApp_Version As FileVersionInfo '=
    'FileVersionInfo.GetVersionInfo($"{StartupPath}\{ProgramAllName.fileName_updateProgram}.exe")
    ''' <summary>
    ''' 更新端-更新主程式-版本 this main program version，例如:1.2.3
    ''' </summary>
    Dim update_updateApp_Version As FileVersionInfo '=
    'FileVersionInfo.GetVersionInfo($"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\更新\{ProgramAllName.fileName_updateProgram}.exe") '執行端版本 this main program version，例如:1.2.3


    ''' <summary>
    ''' 檢查是否需要更新
    ''' </summary>
    ''' <param name="isUpdateButton">是否為檢查扭按下?</param>
    Private Sub check_File_Version(isUpdateButton As Boolean)
        Try
            Try
                thisApp_Version =
                    FileVersionInfo.GetVersionInfo($"{StartupPath}\{thisApp_fullName}.exe")
                updateApp_Version =
                    FileVersionInfo.GetVersionInfo($"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\更新\{thisApp_fullName}.exe") '執行端版本 this main program version，例如:1.2.3
                thisUpdateApp_Version =
                    FileVersionInfo.GetVersionInfo($"{StartupPath}\{ProgramAllName.fileName_updateProgram}.exe")
                update_updateApp_Version =
                    FileVersionInfo.GetVersionInfo($"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\更新\{ProgramAllName.fileName_updateProgram}.exe") '執行端版本 this main program version，例如:1.2.3
            Catch ex As Exception
                Me.Text = $"{thisApp_fullName}更新失敗"
                MsgBox($"找不到檔案{vbCrLf}{ex.Message}", MsgBoxStyle.Critical, "錯誤")
                writeTitleIntoError_InfoTxt($"check_File_Version")
                writeInfoError_InfoTxt($"找不到檔案{vbCrLf}{ex.Message}")
                Exit Sub
            End Try

            If compare_FileVersion_haveToUpdate(thisApp_Version, updateApp_Version) Then
                Me.Text = $"{thisApp_fullName}目前為舊版本號碼:ver.{thisApp_Version.FileVersion}"
                Dim result As MsgBoxResult =
                    MsgBox($"有更新版本! 最新版本為:ver.{updateApp_Version.FileVersion}{vbCrLf}是否自動更新?{vbCrLf}更新資訊請至『關於』查看", vbYesNo, "更新訊息")

                If result = MsgBoxResult.Yes Then
                    If compare_FileVersion_haveToUpdate(thisUpdateApp_Version, update_updateApp_Version) Then
                        For Each myFile In Directory.GetFileSystemEntries(updateTool_path) '更新資料夾
                            If Dir(myFile, vbDirectory) = $"{ProgramAllName.fileName_updateProgram}.exe" Then
                                FileCopy(myFile, StartupPath & "\" & Path.GetFileName(myFile))
                            End If
                        Next
                    End If

                    Dim update_p() As Process
                    Using p As Process = New Process()
                        p.Start($"{StartupPath}\{ProgramAllName.fileName_updateProgram}.exe")
                    End Using
                    update_p = Process.GetProcessesByName($"{ProgramAllName.fileName_updateProgram}")

                    If update_p.Count > 0 Then
                        Me.Close()
                    End If
                End If
            Else
                Me.Text = $"{thisApp_fullName}目前為最新版本:ver.{thisApp_Version.FileVersion}"
                If isUpdateButton Then
                    MsgBox($"{thisApp_fullName}目前為最新版本:ver.{thisApp_Version.FileVersion}{vbCrLf}不更新{vbCrLf}更新資訊請至『關於』查看", , "更新資訊")
                End If
            End If
        Catch ex As Exception
            Me.Text = $"{thisApp_fullName}更新失敗"
            MsgBox($"更新失敗{vbCrLf}{ex.Message}", MsgBoxStyle.Critical, "錯誤")
            writeTitleIntoError_InfoTxt($"check_File_Version")
            writeInfoError_InfoTxt($"更新失敗{vbCrLf}{ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 比較版本是否需要更新? Ture要/False否
    ''' </summary>
    ''' <returns></returns>
    Private Function compare_FileVersion_haveToUpdate(thisVer As FileVersionInfo, updateVer As FileVersionInfo) As Boolean

        Dim thisAppVer_First, thisAppVer_Second, thisAppVer_Third As Integer

        thisAppVer_First = thisVer.FileMajorPart  '1.2.3取得版本的1
        thisAppVer_Second = thisVer.FileMinorPart '1.2.3取得版本的2
        thisAppVer_Third = thisVer.FileBuildPart  '1.2.3取得版本的3

        Dim updateAppVer_First, updateAppVer_Second, updateAppVer_Third As Integer

        updateAppVer_First = updateVer.FileMajorPart  '1.2.3取得版本的1
        updateAppVer_Second = updateVer.FileMinorPart '1.2.3取得版本的2
        updateAppVer_Third = updateVer.FileBuildPart  '1.2.3取得版本的3

        If thisAppVer_First < updateAppVer_First Then Return True
        If thisAppVer_First > updateAppVer_First Then Return False

        If thisAppVer_Second < updateAppVer_Second Then Return True
        If thisAppVer_Second > updateAppVer_Second Then Return False

        If thisAppVer_Third < updateAppVer_Third Then Return True
        If thisAppVer_Third > updateAppVer_Third Then Return False

    End Function

    Private Sub MagicTool_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'read
        'LoadIni()

        'If IO.File.Exists(note_dat) Then
        '    note_TextBox.Text = IO.File.ReadAllText(note_dat)
        'End If

        'If IO.File.Exists(New_noteDat_path) Then
        '    DateNote_TextBox.Text = IO.File.ReadAllText(New_noteDat_path)
        'Else
        '    DateNote_TextBox.Text = note_DateTimePicker.Value.ToShortDateString
        'End If


        'hotkey timer_tick()
        Me.KeyPreview = True
        MagicTool_Timer.Enabled = True
        MagicTool_Timer.Interval = 1

    End Sub

    Public Sub New()

        ' 設計工具需要此呼叫。
        InitializeComponent()

        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        'MagicTool_NotifyIcon.ContextMenuStrip = MagicToll_MenuStrip.ContextMenuStrip
        'MagicTool_NotifyIcon.Text = Me.Text

        'datFile_load()

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
            MsgBox($"ChangLick.LoadIni，訊息{e.Message}")
        End Try
    End Sub


    '點快捷錯誤時判斷
    Public Sub open_DirectPath(dir_path As String)
        Try
            Process.Start(dir_path)
        Catch ex As Exception
            MsgBox("沒有指定目錄/目錄錯誤",, "路徑錯誤")
        End Try
    End Sub
    Private Sub Manual_ToolStrip_Click(sender As Object, e As EventArgs) Handles Manual_ToolStrip.Click
        If Directory.Exists($"M:\DESIGN\BACK UP\yc_tian\Tool Application\使用者使用說明書\{ProgramAllName.fileName_Manualpptx}") Then
            open_DirectPath($"M:\DESIGN\BACK UP\yc_tian\Tool Application\使用者使用說明書\{ProgramAllName.fileName_Manualpptx}")
        Else
            open_DirectPath($"{StartupPath}\{ProgramAllPath.folderName_ppt}\{ProgramAllName.fileName_Manualpptx}")
        End If
    End Sub
    Private Sub Link1_Button_Click(sender As Object, e As EventArgs) Handles Link1_1_Button.Click
        open_DirectPath(chalink.Link1_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_Button_Click(sender As Object, e As EventArgs) Handles Link1_2_Button.Click
        open_DirectPath(chalink.Link2_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_Button_Click(sender As Object, e As EventArgs) Handles Link1_3_Button.Click
        open_DirectPath(chalink.Link3_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_Button_Click(sender As Object, e As EventArgs) Handles Link1_4_Button.Click
        open_DirectPath(chalink.Link4_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_Button_Click(sender As Object, e As EventArgs) Handles Link1_5_Button.Click
        open_DirectPath(chalink.Link5_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_Button_Click(sender As Object, e As EventArgs) Handles Link1_6_Button.Click
        open_DirectPath(chalink.Link6_Dir_TextBox.Text)
    End Sub
    Private Sub Link7_Button_Click(sender As Object, e As EventArgs) Handles Link1_7_Button.Click
        open_DirectPath(chalink.Link7_Dir_TextBox.Text)
    End Sub
    Private Sub Link8_Button_Click(sender As Object, e As EventArgs) Handles Link1_8_Button.Click
        open_DirectPath(chalink.Link8_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_1_Button_Click(sender As Object, e As EventArgs) Handles Link2_1_Button.Click
        open_DirectPath(chalink.Link2_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_2_Button_Click(sender As Object, e As EventArgs) Handles Link2_2_Button.Click
        open_DirectPath(chalink.Link2_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_3_Button_Click(sender As Object, e As EventArgs) Handles Link2_3_Button.Click
        open_DirectPath(chalink.Link2_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_4_Button_Click(sender As Object, e As EventArgs) Handles Link2_4_Button.Click
        open_DirectPath(chalink.Link2_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_5_Button_Click(sender As Object, e As EventArgs) Handles Link2_5_Button.Click
        open_DirectPath(chalink.Link2_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_6_Button_Click(sender As Object, e As EventArgs) Handles Link2_6_Button.Click
        open_DirectPath(chalink.Link2_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_7_Button_Click(sender As Object, e As EventArgs) Handles Link2_7_Button.Click
        open_DirectPath(chalink.Link2_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link2_8_Button_Click(sender As Object, e As EventArgs) Handles Link2_8_Button.Click
        open_DirectPath(chalink.Link2_8_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_1_Button_Click(sender As Object, e As EventArgs) Handles Link3_1_Button.Click
        open_DirectPath(chalink.Link3_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_2_Button_Click(sender As Object, e As EventArgs) Handles Link3_2_Button.Click
        open_DirectPath(chalink.Link3_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_3_Button_Click(sender As Object, e As EventArgs) Handles Link3_3_Button.Click
        open_DirectPath(chalink.Link3_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_4_Button_Click(sender As Object, e As EventArgs) Handles Link3_4_Button.Click
        open_DirectPath(chalink.Link3_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_5_Button_Click(sender As Object, e As EventArgs) Handles Link3_5_Button.Click
        open_DirectPath(chalink.Link3_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_6_Button_Click(sender As Object, e As EventArgs) Handles Link3_6_Button.Click
        open_DirectPath(chalink.Link3_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_7_Button_Click(sender As Object, e As EventArgs) Handles Link3_7_Button.Click
        open_DirectPath(chalink.Link3_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link3_8_Button_Click(sender As Object, e As EventArgs) Handles Link3_8_Button.Click
        open_DirectPath(chalink.Link3_8_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_1_Button_Click(sender As Object, e As EventArgs) Handles Link4_1_Button.Click
        open_DirectPath(chalink.Link4_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_2_Button_Click(sender As Object, e As EventArgs) Handles Link4_2_Button.Click
        open_DirectPath(chalink.Link4_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_3_Button_Click(sender As Object, e As EventArgs) Handles Link4_3_Button.Click
        open_DirectPath(chalink.Link4_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_4_Button_Click(sender As Object, e As EventArgs) Handles Link4_4_Button.Click
        open_DirectPath(chalink.Link4_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_5_Button_Click(sender As Object, e As EventArgs) Handles Link4_5_Button.Click
        open_DirectPath(chalink.Link4_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_6_Button_Click(sender As Object, e As EventArgs) Handles Link4_6_Button.Click
        open_DirectPath(chalink.Link4_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_7_Button_Click(sender As Object, e As EventArgs) Handles Link4_7_Button.Click
        open_DirectPath(chalink.Link4_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link4_8_Button_Click(sender As Object, e As EventArgs) Handles Link4_8_Button.Click
        open_DirectPath(chalink.Link4_8_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_1_Button_Click(sender As Object, e As EventArgs) Handles Link5_1_Button.Click
        open_DirectPath(chalink.Link5_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_2_Button_Click(sender As Object, e As EventArgs) Handles Link5_2_Button.Click
        open_DirectPath(chalink.Link5_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_3_Button_Click(sender As Object, e As EventArgs) Handles Link5_3_Button.Click
        open_DirectPath(chalink.Link5_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_4_Button_Click(sender As Object, e As EventArgs) Handles Link5_4_Button.Click
        open_DirectPath(chalink.Link5_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_5_Button_Click(sender As Object, e As EventArgs) Handles Link5_5_Button.Click
        open_DirectPath(chalink.Link5_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_6_Button_Click(sender As Object, e As EventArgs) Handles Link5_6_Button.Click
        open_DirectPath(chalink.Link5_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_7_Button_Click(sender As Object, e As EventArgs) Handles Link5_7_Button.Click
        open_DirectPath(chalink.Link5_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link5_8_Button_Click(sender As Object, e As EventArgs) Handles Link5_8_Button.Click
        open_DirectPath(chalink.Link5_8_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_1_Button_Click(sender As Object, e As EventArgs) Handles Link6_1_Button.Click
        open_DirectPath(chalink.Link6_1_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_2_Button_Click(sender As Object, e As EventArgs) Handles Link6_2_Button.Click
        open_DirectPath(chalink.Link6_2_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_3_Button_Click(sender As Object, e As EventArgs) Handles Link6_3_Button.Click
        open_DirectPath(chalink.Link6_3_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_4_Button_Click(sender As Object, e As EventArgs) Handles Link6_4_Button.Click
        open_DirectPath(chalink.Link6_4_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_5_Button_Click(sender As Object, e As EventArgs) Handles Link6_5_Button.Click
        open_DirectPath(chalink.Link6_5_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_6_Button_Click(sender As Object, e As EventArgs) Handles Link6_6_Button.Click
        open_DirectPath(chalink.Link6_6_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_7_Button_Click(sender As Object, e As EventArgs) Handles Link6_7_Button.Click
        open_DirectPath(chalink.Link6_7_Dir_TextBox.Text)
    End Sub
    Private Sub Link6_8_Button_Click(sender As Object, e As EventArgs) Handles Link6_8_Button.Click
        open_DirectPath(chalink.Link6_8_Dir_TextBox.Text)
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



    Private Sub MagicTool_Timer_Tick(sender As Object, e As EventArgs) Handles MagicTool_Timer.Tick
        Dim ctrlKey As Boolean
        Dim QKey As Boolean
        ctrlKey = GetAsyncKeyState(Keys.ControlKey)
        QKey = GetAsyncKeyState(Keys.Q)

        If ctrlKey And QKey = True Then

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
            JustForFun_ToolTip.SetToolTip(Me.JustForFun_SplitContainer, "(๑•̀ω•́)ノ 想偷看阿!")
        ElseIf JustForFun_SplitContainer.SplitterDistance > 235 And JustForFun_SplitContainer.SplitterDistance < 350 Then
            JustForFun_ToolTip.SetToolTip(Me.JustForFun_SplitContainer, "(ΦωΦ)喵~")
        Else
            JustForFun_ToolTip.SetToolTip(Me.JustForFun_SplitContainer, "唉丫~被你發現了甚麼")
        End If
    End Sub

    Private Sub BackOriPos_ToolStrip_Click(sender As Object, e As EventArgs) Handles BackOriPos_ToolStrip.Click
        chalink.formPositionOnScreen_Setting(Me, chalink.sKeyValueScr.ToString, chalink.sKeyValuePos.ToString)
    End Sub

    Private Sub note_TextBox_KeyUp(sender As Object, e As KeyEventArgs) Handles note_TextBox.KeyUp
        If IO.File.Exists(note_dat) Then
            File.WriteAllText(note_dat, note_TextBox.Text)
        Else
            MsgBox("寫入之檔案不存在，請重新導入" & vbCrLf & "請將Note.dat檔案移置\dat資料夾底下",, "dat檔案遺失")
            Process.Start($"{StartupPath}\{ProgramAllPath.folderName_dat}")
        End If
    End Sub
    Private Sub DateNote_TextBox_KeyUp(sender As Object, e As KeyEventArgs) Handles DateNote_TextBox.KeyUp
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
            open_DirectPath(fatherPath_to_ChildPAth)


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

        filter_name = {"*.doc", "*.xls"}
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
        check_File_Version(isUpdateButton:=True)
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
    '定義連結按鈕的狀態
End Class



