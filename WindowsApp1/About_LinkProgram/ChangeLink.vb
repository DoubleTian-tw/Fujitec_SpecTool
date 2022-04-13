Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.SystemInformation
Imports System.IO

Module modINI
    '* lpAppName：指向包含Section 名稱的字符串地址
    '* lpKeyName：指向包含Key 名稱的字符串地址
    '* lpDefault：如果Key 值沒有找到，缺省返回缺省的字符串
    '* lpReturnedString：用於保存返回字符串的緩衝區
    '* nSize： 緩衝區的長度
    '* lpFileName ：ini 文件的文件名
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpApplicationName As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpDefault As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder, ByVal nSize As UInt32,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As UInt32

    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpApplicationName As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As UInt32

    Public Declare Ansi Function FlushPrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (
                                      ByVal lpApplicationName As Integer,
                                      ByVal lpKeyName As Integer,
                                      ByVal lpString As Integer,
                                      ByVal lpFileName As String) As Integer

End Module

Public Class ChangeLink

    Public sKeyValueNa, sKeyValuePath, sKeyValueCB, sKeyValueScr,
           sKeyValuePos, sKeyValueCB_State, sKeyValue_note, sKeyNewFolder,
           sKeyFolderPath, sKeyFolderName, sKeyComJobName, sKeyComJobPath,
           sKeyAllEmp, sKeyLocal, sKey_getNameManager, sKey_setColor As New StringBuilder(512)

    Dim nSize As UInt32 = Convert.ToUInt32(1024)
    Dim sinifilename As String =
        $"{Application.StartupPath}\{ProgramAllPath.folderName_ini}\{ProgramAllName.fileName_SetFileIni}"

    ''' <summary>
    ''' [Magic Tool > Link Button總共有六欄]
    ''' </summary>
    Dim ini_linkBtnCol As Integer = 6
    ''' <summary>
    ''' [Magic Tool > Link Button 一欄有八列]
    ''' </summary>
    Dim ini_linkBtnRow As Integer = 8 '1-1~1-8 ...6-1~6-8
    Public ini_LCB, ini_Name, ini_LinkPath, ini_NewFolder, ini_FolderPath,
           ini_FolderName, ini_comJobName, ini_comJobPath, ini_allEmp, ini_allDwgPrk As String


    Dim linkBtn_isTransarent As Boolean ' 連結按鈕的透明度

    Dim count_allFolderCBPath As Integer = 0 'count all folder path sum
    Dim count_allFolderCBName As Integer = 0 'count all folder name sum
    Dim count_comJobCBPath As Integer = 0 'count common job path
    Dim count_comJobCBName As Integer = 0 'count common job file name

    Public combobox_state As controlStateOnChangeLink = New controlStateOnChangeLink()
    Private Sub ChangeLink_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Width = 460
        Me.Height = 410

        Initialization_ini()

        Topmost_setting(Me, False)

        formPositionOnScreen_Setting(Me, ScreenChoose_Label.Text, ScreenPos_Label.Text)
        If ScreenChoose_CB.Text = "" Or ScreenPosition_CB.Text = "" Then
            ScreenChoose_CB.Text = ScreenChoose_Label.Text
            ScreenPosition_CB.Text = ScreenPos_Label.Text
        End If

    End Sub

    ''' <summary>
    ''' [Change Link > 初始化寫入ini的值]
    ''' </summary>
    Public Sub Initialization_ini()
        Try

            'Note save always
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_CheckBox_State,
                                    ChangeLink_BasicString.setCont_NoteSave, "", sKeyValueCB_State, nSize, sinifilename)
            note_CheckBox.Checked = sKeyValueCB_State.ToString
            'Topmost
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_CheckBox_State,
                                    ChangeLink_BasicString.setCont_TopmostSet, "", sKeyValueCB_State, nSize, sinifilename)
            Topmost_CheckBox.Checked = sKeyValueCB_State.ToString
            'autoProgram
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_CheckBox_State2,
                                    ChangeLink_BasicString.setCont_autoProgram, "", sKeyValueCB_State, nSize, sinifilename)
            autoProgram_CheckBox.Checked = sKeyValueCB_State.ToString
            'updateINI
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_CheckBox_State2,
                                    ChangeLink_BasicString.setCont_UpdateINI, "", sKeyValueCB_State, nSize, sinifilename)
            UpdateINI_CheckBox.Checked = sKeyValueCB_State.ToString
            'backupNotice
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_CheckBox_State2,
                                    ChangeLink_BasicString.setCont_backupNotice, "", sKeyValueCB_State, nSize, sinifilename)
            Backup_Notice_CheckBox.Checked = sKeyValueCB_State.ToString

            'ScrChoose
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_ScreenChoose,
                                    ChangeLink_BasicString.setCont_Screen, "", sKeyValueScr, nSize, sinifilename)
            ScreenChoose_Label.Text = sKeyValueScr.ToString
            'ScrPos
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_ScreenPos,
                                    ChangeLink_BasicString.setCont_Pos, "", sKeyValuePos, nSize, sinifilename)
            ScreenPos_Label.Text = sKeyValuePos.ToString

            'Link
            Dim LinkCheckBox As CheckBox(,) = {{Link1_1_CheckBox, Link1_2_CheckBox, Link1_3_CheckBox, Link1_4_CheckBox, Link1_5_CheckBox, Link1_6_CheckBox, Link1_7_CheckBox, Link1_8_CheckBox} _
                                            , {Link2_1_CheckBox, Link2_2_CheckBox, Link2_3_CheckBox, Link2_4_CheckBox, Link2_5_CheckBox, Link2_6_CheckBox, Link2_7_CheckBox, Link2_8_CheckBox} _
                                            , {Link3_1_CheckBox, Link3_2_CheckBox, Link3_3_CheckBox, Link3_4_CheckBox, Link3_5_CheckBox, Link3_6_CheckBox, Link3_7_CheckBox, Link3_8_CheckBox} _
                                            , {Link4_1_CheckBox, Link4_2_CheckBox, Link4_3_CheckBox, Link4_4_CheckBox, Link4_5_CheckBox, Link4_6_CheckBox, Link4_7_CheckBox, Link4_8_CheckBox} _
                                            , {Link5_1_CheckBox, Link5_2_CheckBox, Link5_3_CheckBox, Link5_4_CheckBox, Link5_5_CheckBox, Link5_6_CheckBox, Link5_7_CheckBox, Link5_8_CheckBox} _
                                            , {Link6_1_CheckBox, Link6_2_CheckBox, Link6_3_CheckBox, Link6_4_CheckBox, Link6_5_CheckBox, Link6_6_CheckBox, Link6_7_CheckBox, Link6_8_CheckBox}}
            Dim LinkNameTextBox As TextBox(,) = {{Link1_Name_TextBox, Link2_Name_TextBox, Link3_Name_TextBox, Link4_Name_TextBox, Link5_Name_TextBox, Link6_Name_TextBox, Link7_Name_TextBox, Link8_Name_TextBox} _
                                               , {Link2_1_Name_TextBox, Link2_2_Name_TextBox, Link2_3_Name_TextBox, Link2_4_Name_TextBox, Link2_5_Name_TextBox, Link2_6_Name_TextBox, Link2_7_Name_TextBox, Link2_8_Name_TextBox} _
                                               , {Link3_1_Name_TextBox, Link3_2_Name_TextBox, Link3_3_Name_TextBox, Link3_4_Name_TextBox, Link3_5_Name_TextBox, Link3_6_Name_TextBox, Link3_7_Name_TextBox, Link3_8_Name_TextBox} _
                                               , {Link4_1_Name_TextBox, Link4_2_Name_TextBox, Link4_3_Name_TextBox, Link4_4_Name_TextBox, Link4_5_Name_TextBox, Link4_6_Name_TextBox, Link4_7_Name_TextBox, Link4_8_Name_TextBox} _
                                               , {Link5_1_Name_TextBox, Link5_2_Name_TextBox, Link5_3_Name_TextBox, Link5_4_Name_TextBox, Link5_5_Name_TextBox, Link5_6_Name_TextBox, Link5_7_Name_TextBox, Link5_8_Name_TextBox} _
                                               , {Link6_1_Name_TextBox, Link6_2_Name_TextBox, Link6_3_Name_TextBox, Link6_4_Name_TextBox, Link6_5_Name_TextBox, Link6_6_Name_TextBox, Link6_7_Name_TextBox, Link6_8_Name_TextBox}}
            Dim LinkButton As Button(,) = {{MagicTool.Link1_1_Button, MagicTool.Link1_2_Button, MagicTool.Link1_3_Button, MagicTool.Link1_4_Button, MagicTool.Link1_5_Button, MagicTool.Link1_6_Button, MagicTool.Link1_7_Button, MagicTool.Link1_8_Button} _
                                         , {MagicTool.Link2_1_Button, MagicTool.Link2_2_Button, MagicTool.Link2_3_Button, MagicTool.Link2_4_Button, MagicTool.Link2_5_Button, MagicTool.Link2_6_Button, MagicTool.Link2_7_Button, MagicTool.Link2_8_Button} _
                                         , {MagicTool.Link3_1_Button, MagicTool.Link3_2_Button, MagicTool.Link3_3_Button, MagicTool.Link3_4_Button, MagicTool.Link3_5_Button, MagicTool.Link3_6_Button, MagicTool.Link3_7_Button, MagicTool.Link3_8_Button} _
                                         , {MagicTool.Link4_1_Button, MagicTool.Link4_2_Button, MagicTool.Link4_3_Button, MagicTool.Link4_4_Button, MagicTool.Link4_5_Button, MagicTool.Link4_6_Button, MagicTool.Link4_7_Button, MagicTool.Link4_8_Button} _
                                         , {MagicTool.Link5_1_Button, MagicTool.Link5_2_Button, MagicTool.Link5_3_Button, MagicTool.Link5_4_Button, MagicTool.Link5_5_Button, MagicTool.Link5_6_Button, MagicTool.Link5_7_Button, MagicTool.Link5_8_Button} _
                                         , {MagicTool.Link6_1_Button, MagicTool.Link6_2_Button, MagicTool.Link6_3_Button, MagicTool.Link6_4_Button, MagicTool.Link6_5_Button, MagicTool.Link6_6_Button, MagicTool.Link6_7_Button, MagicTool.Link6_8_Button}}
            Dim LinkDirTextBox As TextBox(,) = {{Link1_Dir_TextBox, Link2_Dir_TextBox, Link3_Dir_TextBox, Link4_Dir_TextBox, Link5_Dir_TextBox, Link6_Dir_TextBox, Link7_Dir_TextBox, Link8_Dir_TextBox} _
                                             , {Link2_1_Dir_TextBox, Link2_2_Dir_TextBox, Link2_3_Dir_TextBox, Link2_4_Dir_TextBox, Link2_5_Dir_TextBox, Link2_6_Dir_TextBox, Link2_7_Dir_TextBox, Link2_8_Dir_TextBox} _
                                             , {Link3_1_Dir_TextBox, Link3_2_Dir_TextBox, Link3_3_Dir_TextBox, Link3_4_Dir_TextBox, Link3_5_Dir_TextBox, Link3_6_Dir_TextBox, Link3_7_Dir_TextBox, Link3_8_Dir_TextBox} _
                                             , {Link4_1_Dir_TextBox, Link4_2_Dir_TextBox, Link4_3_Dir_TextBox, Link4_4_Dir_TextBox, Link4_5_Dir_TextBox, Link4_6_Dir_TextBox, Link4_7_Dir_TextBox, Link4_8_Dir_TextBox} _
                                             , {Link5_1_Dir_TextBox, Link5_2_Dir_TextBox, Link5_3_Dir_TextBox, Link5_4_Dir_TextBox, Link5_5_Dir_TextBox, Link5_6_Dir_TextBox, Link5_7_Dir_TextBox, Link5_8_Dir_TextBox} _
                                             , {Link6_1_Dir_TextBox, Link6_2_Dir_TextBox, Link6_3_Dir_TextBox, Link6_4_Dir_TextBox, Link6_5_Dir_TextBox, Link6_6_Dir_TextBox, Link6_7_Dir_TextBox, Link6_8_Dir_TextBox}}
            Dim CreateNewFolder_Textbox As TextBox() = {childFolder_TextBox1, childFolder_TextBox2, childFolder_TextBox3, childFolder_TextBox4, childFolder_TextBox5 _
                                                    , childFolder_TextBox6}

            For i = 1 To ini_linkBtnCol
                For j = 1 To ini_linkBtnRow
                    ini_LCB = $"{ChangeLink_BasicString.setCont_LCB_}{i}_{j}"
                    GetPrivateProfileString(ChangeLink_BasicString.setTitle_LinkCheckbox,
                                            ini_LCB, "", sKeyValueCB, nSize, sinifilename)
                    LinkCheckBox(i - 1, j - 1).Checked = sKeyValueCB.ToString

                    ini_Name = $"{ChangeLink_BasicString.setCont_Name_}{i}_{j}"
                    GetPrivateProfileString(ChangeLink_BasicString.setTitle_Name,
                                            ini_Name, "", sKeyValueNa, nSize, sinifilename)
                    LinkNameTextBox(i - 1, j - 1).Text = sKeyValueNa.ToString

                    ini_LinkPath = $"{ChangeLink_BasicString.setCont_LinkPath_}{i}_{j}"
                    GetPrivateProfileString(ChangeLink_BasicString.setTitle_LinkPath,
                                            ini_LinkPath, "", sKeyValuePath, nSize, sinifilename)
                    LinkDirTextBox(i - 1, j - 1).Text = sKeyValuePath.ToString
                Next
                'New folder
                ini_NewFolder = $"{ChangeLink_BasicString.setCont_ChildFolder_}{i}"
                GetPrivateProfileString(ChangeLink_BasicString.setTitle_NewFolder,
                                        ini_NewFolder, "", sKeyNewFolder, nSize, sinifilename)
                CreateNewFolder_Textbox(i - 1).Text = sKeyNewFolder.ToString
            Next
        Catch e As Exception
            MsgBox($"ChangLick.initialization_ini1，訊息{e.ToString}",, "錯誤訊息")
        End Try



        '預設路徑TabPage ------------------------------------------------------------------------------------------------+
        Try
            '預設路徑 > JobMaker > Load > 仕樣書
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_DefaultPath,
                                    ChangeLink_BasicString.setCont_JobMaker_Spec_Path, "", sKey_getNameManager, nSize, sinifilename)
            ChgLink_DefaultPath_Spec_TextBox.Text = sKey_getNameManager.ToString

            '預設路徑 > JobMaker > Load > CheckList
            GetPrivateProfileString(ChangeLink_BasicString.setTitle_DefaultPath,
                                    ChangeLink_BasicString.setCont_JobMaker_Spec_Path, "", sKey_getNameManager, nSize, sinifilename)
            ChgLink_DefaultPath_CheckList_TextBox.Text = sKey_getNameManager.ToString

            '預設路徑 > JobMaker > Load > 載入SQLite
            ChgLink_DefaultPath_SQLite_TextBox.Text = ProgramAllPath.SQLite_connectionPath_Tool
        Catch e_defaultPath As Exception
            MsgBox($"ChangeLink > 預設路徑錯誤 : {e_defaultPath}",, "錯誤訊息")
        End Try
        '------------------------------------------------------------------------------------------------ 預設路徑TabPage
        Try
            getSetColor() '顏色設定
        Catch e_setColor As Exception
            MsgBox($"ChangeLink > getSetColor顏色設定錯誤 : {e_setColor}",, "錯誤訊息")
        End Try

        'DWG
        Try
            'folder path and name 、 common job file
            count_allFolderCBPath = AllFolderPath_ComboBox.Items.Count '先計算combox有幾個，一開始一定是0
            count_allFolderCBName = CL_AllFolderName_ComboBox.Items.Count
            count_comJobCBPath = CommonJobFilePath_Combobox.Items.Count '常用工番count
            count_comJobCBName = CommonJobFileName_ComboBox.Items.Count

            count_allFolderCBPath =
                combobox_state.combobox_when_add(count_allFolderCBPath, ini_FolderPath,
                                                 ChangeLink_BasicString.setCont_FolderPath_, ChangeLink_BasicString.setTitle_AllFolderPath,
                                                 sKeyFolderPath, AllFolderPath_ComboBox, MagicTool.FolderPath_ComboBox, nSize, sinifilename)
            count_allFolderCBName =
                combobox_state.combobox_when_add(count_allFolderCBName, ini_FolderName,
                                                 ChangeLink_BasicString.setCont_FolderName_, ChangeLink_BasicString.setTitle_AllFolderName,
                                                 sKeyFolderName, CL_AllFolderName_ComboBox, MagicTool.FolderName_ComboBox, nSize, sinifilename)
            count_comJobCBPath =
                combobox_state.combobox_when_add(count_comJobCBPath, ini_comJobPath,
                                                 ChangeLink_BasicString.setCont_CommonFilePath_, ChangeLink_BasicString.setTitle_CommonJobFile,
                                                 sKeyComJobPath, CommonJobFilePath_Combobox, MagicTool.FileChoUse_ComboBox, nSize, sinifilename)

            '當FileChoUse_Combobox的text為某資料夾時，顯示該資料夾底下的檔案
            MagicTool.FileChoUse_ComboBox.Text = MagicTool.FileChoUse_ComboBox.Items(0).ToString


            Dim MA_ChildFolder_Group As TextBox() = {MagicTool.MAchildFolder_TextBox1, MagicTool.MAchildFolder_TextBox2,
                                                     MagicTool.MAchildFolder_TextBox3, MagicTool.MAchildFolder_TextBox4,
                                                     MagicTool.MAchildFolder_TextBox5, MagicTool.MAchildFolder_TextBox6}
            Dim CL_ChildFolder_Group As TextBox() = {childFolder_TextBox1, childFolder_TextBox2, childFolder_TextBox3,
                                                     childFolder_TextBox4, childFolder_TextBox5, childFolder_TextBox6}
            For a = 0 To MagicTool.childForlder_sum - 1  '0 ~ 5
                MA_ChildFolder_Group(a).Text = CL_ChildFolder_Group(a).Text
            Next a
        Catch e As Exception
            MsgBox($"ChangLick.initialization_ini2，訊息{e.ToString}",, "錯誤訊息")
        End Try

        '自動檢查INI更新狀態
        Try
            updateINI_program()
        Catch e As Exception
            MsgBox($"ChangLick.updateINI_program，訊息{e.ToString}",, "錯誤訊息")
        End Try
        '自動檢查備分更新
        Try
            fileUpdateNotice_check()
        Catch ex As Exception
            MsgBox($"ChangLick.update_file_check，訊息{ex.ToString}",, "錯誤訊息")
        End Try

    End Sub

    Private Sub getSetColor() '顏色設定
        Try
            'LinkBtn
            SetLinkBtn_Result_Button.FlatStyle = Windows.Forms.FlatStyle.Flat

            GetPrivateProfileString(ChangeLink_BasicString.setTitle_SettingColor,
                                    ChangeLink_BasicString.setCont_SetLinkBtn_MouseOverColor, "", sKey_setColor, nSize, sinifilename)
            SetLinkBtn_MouseOverColor_Button.Text = sKey_setColor.ToString 'LinkBtn滑鼠滑過反白顏色
            If SetLinkBtn_MouseOverColor_Button.Text = "" Then
                SetLinkBtn_Result_Button.FlatAppearance.MouseOverBackColor = DefaultBackColor
            Else
                SetLinkBtn_Result_Button.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(SetLinkBtn_MouseOverColor_Button.Text)
            End If

            GetPrivateProfileString(ChangeLink_BasicString.setTitle_SettingColor,
                                    ChangeLink_BasicString.setCont_SetLinkBtn_FontColor, "", sKey_setColor, nSize, sinifilename)
            SetLinkBtn_FontColor_Button.Text = sKey_setColor.ToString 'LinkBtn字體顏色
            If SetLinkBtn_FontColor_Button.Text = "" Then
                SetLinkBtn_FontColor_Button.ForeColor = DefaultForeColor
            Else
                SetLinkBtn_Result_Button.ForeColor = ColorTranslator.FromHtml(SetLinkBtn_FontColor_Button.Text)
            End If

            GetPrivateProfileString(ChangeLink_BasicString.setTitle_SettingColor,
                                    ChangeLink_BasicString.setCont_SetLinkBtn_TransparentColor, "", sKey_setColor, nSize, sinifilename)
            SetLinkBtn_Transparent_Button.Text = sKey_setColor.ToString 'LinkBtn是否透明?
            If SetLinkBtn_Transparent_Button.Text = "YES" Then
                linkBtn_isTransarent = True
                SetLinkBtn_Result_Button.BackColor = Color.Transparent
                SetLinkBtn_MouseOverColor_Button.BackColor = Color.Transparent
            ElseIf SetLinkBtn_Transparent_Button.Text = "NO" Then
                linkBtn_isTransarent = False
                SetLinkBtn_Result_Button.BackColor = DefaultBackColor
                SetLinkBtn_MouseOverColor_Button.BackColor = DefaultBackColor
            Else
                linkBtn_isTransarent = False
                SetLinkBtn_Result_Button.BackColor = DefaultBackColor
                SetLinkBtn_MouseOverColor_Button.BackColor = DefaultBackColor
            End If

            GetPrivateProfileString(ChangeLink_BasicString.setTitle_SettingColor,
                                    ChangeLink_BasicString.setCont_SetLinkBtn_BorderColor, "", sKey_setColor, nSize, sinifilename)
            SetLinkBtn_BorderColor_Button.Text = sKey_setColor.ToString 'LinkBtn邊界顏色
            If SetLinkBtn_BorderColor_Button.Text = "" Then
                SetLinkBtn_Result_Button.FlatAppearance.BorderColor = DefaultBackColor
            Else
                SetLinkBtn_Result_Button.FlatAppearance.BorderColor = ColorTranslator.FromHtml(SetLinkBtn_BorderColor_Button.Text)
            End If

            GetPrivateProfileString(ChangeLink_BasicString.setTitle_SettingColor,
                                    ChangeLink_BasicString.setCont_SetLinkBtn_BgPicture, "", sKey_setColor, nSize, sinifilename)
            SetLinkBtn_BgPicture_TextBox.Text = sKey_setColor.ToString

            GetPrivateProfileString(ChangeLink_BasicString.setTitle_SettingColor,
                                    ChangeLink_BasicString.setCont_SetNote_BackColor, "", sKey_setColor, nSize, sinifilename)
            SetNote_BackColor_Button.Text = sKey_setColor.ToString
            If SetNote_BackColor_Button.Text = "" Then
                SetNote_Result_TextBox.BackColor = DefaultBackColor
            Else
                SetNote_Result_TextBox.BackColor = ColorTranslator.FromHtml(SetNote_BackColor_Button.Text)
            End If

            GetPrivateProfileString(ChangeLink_BasicString.setTitle_SettingColor,
                                    ChangeLink_BasicString.setCont_SetNote_FontColor, "", sKey_setColor, nSize, sinifilename)
            SetNote_FontColor_Button.Text = sKey_setColor.ToString
            If SetNote_FontColor_Button.Text = "" Then
                SetNote_FontColor_Button.ForeColor = DefaultForeColor
            Else
                SetNote_FontColor_Button.ForeColor = ColorTranslator.FromHtml(SetNote_FontColor_Button.Text)
            End If
            'LinkBtn
        Catch e As Exception
            MsgBox($"ChangLick.getSetColor錯誤，訊息{e.Message}",, "錯誤訊息")
        End Try
    End Sub


    Public Sub fileUpdateNotice_check()
        If Backup_Notice_CheckBox.Checked Then
            Dim updateTXT_path As String =
                $"{ProgramAllPath.path_toolProgram}\{ProgramAllPath.folderName_update}\{ProgramAllName.fileName_updateNoticeFile}.txt"
            Try
                Dim myUpdate_txt As IO.StreamReader = New IO.StreamReader(updateTXT_path)
                Dim myLine As String = ""

                Do Until myLine = "-"
                    myLine = myUpdate_txt.ReadLine()
                    If myLine <> "" And myLine <> "-" Then
                        MsgBox("資料文件有更新!!", vbInformation, "提示")
                        Process.Start(updateTXT_path)
                        Exit Do
                    End If
                Loop
            Catch e As Exception
                MsgBox($"錯誤:{ProgramAllName.fileName_updateNoticeFile}找不到{vbCrLf}{e.Message}",, "錯誤訊息")
            End Try
        End If
    End Sub

    ''' <summary>
    ''' [設定Form在螢幕上的位置]
    ''' </summary>
    ''' <param name="mFormName">Form Name</param>
    ''' <param name="mScreen">主螢幕或副螢幕</param>
    ''' <param name="mPosition">左上/右上/左下/右下</param>
    Public Sub formPositionOnScreen_Setting(mFormName As Form, mScreen As String, mPosition As String)

        If mScreen = ChangeLink_BasicString.screen_Main Then '"主螢幕" 

            If mPosition = ChangeLink_BasicString.screenPosition_LeftTop Then '"左上" 
                fromLocation_Setting(mFormName, 0, 0)
            ElseIf mPosition = ChangeLink_BasicString.screenPosition_RightTop Then '"右上" 
                fromLocation_Setting(mFormName, WorkingArea.Width - mFormName.Width, 0)
            ElseIf mPosition = ChangeLink_BasicString.screenPosition_LeftBtm Then '"左下" 
                fromLocation_Setting(mFormName, 0, WorkingArea.Height - mFormName.Height)
            ElseIf mPosition = ChangeLink_BasicString.screenPosition_RightBtm Then '"右下" 
                fromLocation_Setting(mFormName, WorkingArea.Width - mFormName.Width, WorkingArea.Height - mFormName.Height)
            End If
        ElseIf mScreen = ChangeLink_BasicString.screen_Sub Then '"副螢幕" 
            If mPosition = ChangeLink_BasicString.screenPosition_LeftTop Then '"左上" 
                fromLocation_Setting(mFormName, -1920, 0)
            ElseIf mPosition = ChangeLink_BasicString.screenPosition_RightTop Then '"右上" 
                fromLocation_Setting(mFormName, -mFormName.Width, 0)
            ElseIf mPosition = ChangeLink_BasicString.screenPosition_LeftBtm Then '"左下" 
                fromLocation_Setting(mFormName, -1920, WorkingArea.Height - mFormName.Height)
            ElseIf mPosition = ChangeLink_BasicString.screenPosition_RightBtm Then '"右下" 
                fromLocation_Setting(mFormName, -mFormName.Width, WorkingArea.Height - mFormName.Height)
            End If
        ElseIf mScreen = ChangeLink_BasicString.screen_Custom And mPosition = ChangeLink_BasicString.screenPosition_Custom Then '"自訂" And mPosition = "自訂" 
            fromLocation_Setting(mFormName, Me.DesktopLocation.X, Me.DesktopLocation.Y)
        End If

    End Sub

    ''' <summary>
    ''' 設定Form的X,Y位址
    ''' </summary>
    ''' <param name="mFormName"></param>
    ''' <param name="mForm_PositionX"></param>
    ''' <param name="mForm_PositionY"></param>
    Private Sub fromLocation_Setting(mFormName As Form, mForm_PositionX As Integer, mForm_PositionY As Integer)
        mFormName.Location = New System.Drawing.Point(mForm_PositionX, mForm_PositionY)
    End Sub

    Public Sub LinkCB_setting() 'All Link 觸發之狀態
        Dim LinkCheckBox As CheckBox(,) = {{Link1_1_CheckBox, Link1_2_CheckBox, Link1_3_CheckBox, Link1_4_CheckBox, Link1_5_CheckBox, Link1_6_CheckBox, Link1_7_CheckBox, Link1_8_CheckBox} _
                                        , {Link2_1_CheckBox, Link2_2_CheckBox, Link2_3_CheckBox, Link2_4_CheckBox, Link2_5_CheckBox, Link2_6_CheckBox, Link2_7_CheckBox, Link2_8_CheckBox} _
                                        , {Link3_1_CheckBox, Link3_2_CheckBox, Link3_3_CheckBox, Link3_4_CheckBox, Link3_5_CheckBox, Link3_6_CheckBox, Link3_7_CheckBox, Link3_8_CheckBox} _
                                        , {Link4_1_CheckBox, Link4_2_CheckBox, Link4_3_CheckBox, Link4_4_CheckBox, Link4_5_CheckBox, Link4_6_CheckBox, Link4_7_CheckBox, Link4_8_CheckBox} _
                                        , {Link5_1_CheckBox, Link5_2_CheckBox, Link5_3_CheckBox, Link5_4_CheckBox, Link5_5_CheckBox, Link5_6_CheckBox, Link5_7_CheckBox, Link5_8_CheckBox} _
                                        , {Link6_1_CheckBox, Link6_2_CheckBox, Link6_3_CheckBox, Link6_4_CheckBox, Link6_5_CheckBox, Link6_6_CheckBox, Link6_7_CheckBox, Link6_8_CheckBox}}
        Dim LinkNameTextBox As TextBox(,) = {{Link1_Name_TextBox, Link2_Name_TextBox, Link3_Name_TextBox, Link4_Name_TextBox, Link5_Name_TextBox, Link6_Name_TextBox, Link7_Name_TextBox, Link8_Name_TextBox} _
                                           , {Link2_1_Name_TextBox, Link2_2_Name_TextBox, Link2_3_Name_TextBox, Link2_4_Name_TextBox, Link2_5_Name_TextBox, Link2_6_Name_TextBox, Link2_7_Name_TextBox, Link2_8_Name_TextBox} _
                                           , {Link3_1_Name_TextBox, Link3_2_Name_TextBox, Link3_3_Name_TextBox, Link3_4_Name_TextBox, Link3_5_Name_TextBox, Link3_6_Name_TextBox, Link3_7_Name_TextBox, Link3_8_Name_TextBox} _
                                           , {Link4_1_Name_TextBox, Link4_2_Name_TextBox, Link4_3_Name_TextBox, Link4_4_Name_TextBox, Link4_5_Name_TextBox, Link4_6_Name_TextBox, Link4_7_Name_TextBox, Link4_8_Name_TextBox} _
                                           , {Link5_1_Name_TextBox, Link5_2_Name_TextBox, Link5_3_Name_TextBox, Link5_4_Name_TextBox, Link5_5_Name_TextBox, Link5_6_Name_TextBox, Link5_7_Name_TextBox, Link5_8_Name_TextBox} _
                                           , {Link6_1_Name_TextBox, Link6_2_Name_TextBox, Link6_3_Name_TextBox, Link6_4_Name_TextBox, Link6_5_Name_TextBox, Link6_6_Name_TextBox, Link6_7_Name_TextBox, Link6_8_Name_TextBox}}
        Dim LinkButton As Button(,) = {{MagicTool.Link1_1_Button, MagicTool.Link1_2_Button, MagicTool.Link1_3_Button, MagicTool.Link1_4_Button, MagicTool.Link1_5_Button, MagicTool.Link1_6_Button, MagicTool.Link1_7_Button, MagicTool.Link1_8_Button} _
                                     , {MagicTool.Link2_1_Button, MagicTool.Link2_2_Button, MagicTool.Link2_3_Button, MagicTool.Link2_4_Button, MagicTool.Link2_5_Button, MagicTool.Link2_6_Button, MagicTool.Link2_7_Button, MagicTool.Link2_8_Button} _
                                     , {MagicTool.Link3_1_Button, MagicTool.Link3_2_Button, MagicTool.Link3_3_Button, MagicTool.Link3_4_Button, MagicTool.Link3_5_Button, MagicTool.Link3_6_Button, MagicTool.Link3_7_Button, MagicTool.Link3_8_Button} _
                                     , {MagicTool.Link4_1_Button, MagicTool.Link4_2_Button, MagicTool.Link4_3_Button, MagicTool.Link4_4_Button, MagicTool.Link4_5_Button, MagicTool.Link4_6_Button, MagicTool.Link4_7_Button, MagicTool.Link4_8_Button} _
                                     , {MagicTool.Link5_1_Button, MagicTool.Link5_2_Button, MagicTool.Link5_3_Button, MagicTool.Link5_4_Button, MagicTool.Link5_5_Button, MagicTool.Link5_6_Button, MagicTool.Link5_7_Button, MagicTool.Link5_8_Button} _
                                     , {MagicTool.Link6_1_Button, MagicTool.Link6_2_Button, MagicTool.Link6_3_Button, MagicTool.Link6_4_Button, MagicTool.Link6_5_Button, MagicTool.Link6_6_Button, MagicTool.Link6_7_Button, MagicTool.Link6_8_Button}}
        Dim LinkDirTextBox As TextBox(,) = {{Link1_Dir_TextBox, Link2_Dir_TextBox, Link3_Dir_TextBox, Link4_Dir_TextBox, Link5_Dir_TextBox, Link6_Dir_TextBox, Link7_Dir_TextBox, Link8_Dir_TextBox} _
                                         , {Link2_1_Dir_TextBox, Link2_2_Dir_TextBox, Link2_3_Dir_TextBox, Link2_4_Dir_TextBox, Link2_5_Dir_TextBox, Link2_6_Dir_TextBox, Link2_7_Dir_TextBox, Link2_8_Dir_TextBox} _
                                         , {Link3_1_Dir_TextBox, Link3_2_Dir_TextBox, Link3_3_Dir_TextBox, Link3_4_Dir_TextBox, Link3_5_Dir_TextBox, Link3_6_Dir_TextBox, Link3_7_Dir_TextBox, Link3_8_Dir_TextBox} _
                                         , {Link4_1_Dir_TextBox, Link4_2_Dir_TextBox, Link4_3_Dir_TextBox, Link4_4_Dir_TextBox, Link4_5_Dir_TextBox, Link4_6_Dir_TextBox, Link4_7_Dir_TextBox, Link4_8_Dir_TextBox} _
                                         , {Link5_1_Dir_TextBox, Link5_2_Dir_TextBox, Link5_3_Dir_TextBox, Link5_4_Dir_TextBox, Link5_5_Dir_TextBox, Link5_6_Dir_TextBox, Link5_7_Dir_TextBox, Link5_8_Dir_TextBox} _
                                         , {Link6_1_Dir_TextBox, Link6_2_Dir_TextBox, Link6_3_Dir_TextBox, Link6_4_Dir_TextBox, Link6_5_Dir_TextBox, Link6_6_Dir_TextBox, Link6_7_Dir_TextBox, Link6_8_Dir_TextBox}}
        Dim LinkOpenFileButton As Button(,) =
            {{Link1_1_OpenFile_Button, Link1_2_OpenFile_Button, Link1_3_OpenFile_Button, Link1_4_OpenFile_Button, Link1_5_OpenFile_Button, Link1_6_OpenFile_Button, Link1_7_OpenFile_Button, Link1_8_OpenFile_Button},
             {Link2_1_OpenFile_Button, Link2_2_OpenFile_Button, Link2_3_OpenFile_Button, Link2_4_OpenFile_Button, Link2_5_OpenFile_Button, Link2_6_OpenFile_Button, Link2_7_OpenFile_Button, Link2_8_OpenFile_Button},
             {Link3_1_OpenFile_Button, Link3_2_OpenFile_Button, Link3_3_OpenFile_Button, Link3_4_OpenFile_Button, Link3_5_OpenFile_Button, Link3_6_OpenFile_Button, Link3_7_OpenFile_Button, Link3_8_OpenFile_Button},
             {Link4_1_OpenFile_Button, Link4_2_OpenFile_Button, Link4_3_OpenFile_Button, Link4_4_OpenFile_Button, Link4_5_OpenFile_Button, Link4_6_OpenFile_Button, Link4_7_OpenFile_Button, Link4_8_OpenFile_Button},
             {Link5_1_OpenFile_Button, Link5_2_OpenFile_Button, Link5_3_OpenFile_Button, Link5_4_OpenFile_Button, Link5_5_OpenFile_Button, Link5_6_OpenFile_Button, Link5_7_OpenFile_Button, Link5_8_OpenFile_Button},
             {Link6_1_OpenFile_Button, Link6_2_OpenFile_Button, Link6_3_OpenFile_Button, Link6_4_OpenFile_Button, Link6_5_OpenFile_Button, Link6_6_OpenFile_Button, Link6_7_OpenFile_Button, Link6_8_OpenFile_Button}}
        For i = 1 To ini_linkBtnCol
            For j = 1 To ini_linkBtnRow
                ini_LCB = $"{ChangeLink_BasicString.setCont_LCB_}{i}_{j}"
                LinkCB_setting_ifelse(LinkCheckBox(i - 1, j - 1),
                                      sKeyValueCB,
                                      ChangeLink_BasicString.setTitle_LinkCheckbox,
                                      ini_LCB,
                                      LinkButton(i - 1, j - 1),
                                      LinkNameTextBox(i - 1, j - 1),
                                      LinkDirTextBox(i - 1, j - 1),
                                      LinkOpenFileButton(i - 1, j - 1))
            Next
        Next


    End Sub

    ''' <summary>
    ''' 寫入ini
    ''' </summary>
    ''' <param name="sKeyVa"></param>
    ''' <param name="setValue"></param>
    ''' <param name="TitleN"></param>
    ''' <param name="SecN"></param>
    Public Sub WriteInini_Fun(sKeyVa As StringBuilder, setValue As String, TitleN As String, SecN As String) '寫入ini
        If sKeyVa.ToString IsNot "" Then
            sKeyVa.Clear()
        End If
        sKeyVa = sKeyVa.Append(setValue)
        WritePrivateProfileString(TitleN, SecN, sKeyVa, sinifilename)
    End Sub

    ''' <summary>
    ''' Link勾勾按鈕的判斷是
    ''' </summary>
    ''' <param name="Link_ChkBx"></param>
    ''' <param name="sKeyVa"></param>
    ''' <param name="LCB"></param>
    ''' <param name="LCB_In"></param>
    ''' <param name="Ma_BTN"></param>
    ''' <param name="LinkNa_TB"></param>
    ''' <param name="LinkDir_TB"></param>
    ''' <param name="LinkOpenFile_Btn"></param>
    Private Sub LinkCB_setting_ifelse(Link_ChkBx As CheckBox, sKeyVa As StringBuilder, LCB As String, LCB_In As String,
                                      Ma_BTN As Button, LinkNa_TB As TextBox, LinkDir_TB As TextBox, LinkOpenFile_Btn As Button)
        If Link_ChkBx.Checked = True Then
            WriteInini_Fun(sKeyVa, True, LCB, LCB_In)
            IfCB_Click(Link_ChkBx, Ma_BTN, LinkNa_TB, LinkDir_TB, LinkOpenFile_Btn, True)
        Else
            WriteInini_Fun(sKeyVa, False, LCB, LCB_In)
            IfCB_Click(Link_ChkBx, Ma_BTN, LinkNa_TB, LinkDir_TB, LinkOpenFile_Btn, True)
        End If
    End Sub

    ''' <summary>
    ''' 介面是否設定在最上層
    ''' </summary>
    ''' <param name="top_name"></param>
    ''' <param name="writable"></param>
    Public Sub Topmost_setting(top_name As Form, writable As Boolean)

        If Topmost_CheckBox.Checked = True Then

            If writable Then
                WriteInini_Fun(sKeyValueCB_State, CStr(True),
                               ChangeLink_BasicString.setTitle_CheckBox_State, ChangeLink_BasicString.setCont_TopmostSet)
            End If

            top_name.TopMost = True

        Else
            If writable Then
                WriteInini_Fun(sKeyValueCB_State, CStr(False),
                               ChangeLink_BasicString.setTitle_CheckBox_State, ChangeLink_BasicString.setCont_TopmostSet)
            End If

            top_name.TopMost = False

        End If
    End Sub

    ''' <summary>
    ''' All Link 是否被勾選，顯示當前能使用或不能 LinkCB_setting_ifelse 使用
    ''' </summary>
    ''' <param name="Link_CB"></param>
    ''' <param name="Link_B"></param>
    ''' <param name="LinkNa_TB"></param>
    ''' <param name="LinkDir_TB"></param>
    ''' <param name="LinkOpenFile_B"></param>
    ''' <param name="WriteAble"></param>
    Private Sub IfCB_Click(Link_CB As CheckBox, Link_B As Button,
                           LinkNa_TB As TextBox, LinkDir_TB As TextBox,
                           LinkOpenFile_B As Button, WriteAble As Boolean)
        If Link_CB.Checked = True Then
            LinkNa_TB.Enabled = True
            LinkDir_TB.Enabled = True
            LinkOpenFile_B.Enabled = True
            If WriteAble = True Then
                Link_B.Enabled = True
                Link_B.Text = LinkNa_TB.Text
            End If
        Else
            LinkNa_TB.Enabled = False
            LinkDir_TB.Enabled = False
            LinkOpenFile_B.Enabled = False
            If WriteAble = True Then
                Link_B.Enabled = False
                Link_B.Text = Link_CB.Text
            End If
        End If

    End Sub

    ''' <summary>
    ''' 保存所有狀態的按鍵
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Confirm_Button_Click(sender As Object, e As EventArgs) Handles Confirm_Button.Click
        'link check box 
        LinkCB_setting()
        'Topmost
        Topmost_setting(Me, True)
        'link
        Dim LinkCheckBox As CheckBox(,) = {{Link1_1_CheckBox, Link1_2_CheckBox, Link1_3_CheckBox, Link1_4_CheckBox, Link1_5_CheckBox, Link1_6_CheckBox, Link1_7_CheckBox, Link1_8_CheckBox} _
                                        , {Link2_1_CheckBox, Link2_2_CheckBox, Link2_3_CheckBox, Link2_4_CheckBox, Link2_5_CheckBox, Link2_6_CheckBox, Link2_7_CheckBox, Link2_8_CheckBox} _
                                        , {Link3_1_CheckBox, Link3_2_CheckBox, Link3_3_CheckBox, Link3_4_CheckBox, Link3_5_CheckBox, Link3_6_CheckBox, Link3_7_CheckBox, Link3_8_CheckBox} _
                                        , {Link4_1_CheckBox, Link4_2_CheckBox, Link4_3_CheckBox, Link4_4_CheckBox, Link4_5_CheckBox, Link4_6_CheckBox, Link4_7_CheckBox, Link4_8_CheckBox} _
                                        , {Link5_1_CheckBox, Link5_2_CheckBox, Link5_3_CheckBox, Link5_4_CheckBox, Link5_5_CheckBox, Link5_6_CheckBox, Link5_7_CheckBox, Link5_8_CheckBox} _
                                        , {Link6_1_CheckBox, Link6_2_CheckBox, Link6_3_CheckBox, Link6_4_CheckBox, Link6_5_CheckBox, Link6_6_CheckBox, Link6_7_CheckBox, Link6_8_CheckBox}}
        Dim LinkNameTextBox As TextBox(,) = {{Link1_Name_TextBox, Link2_Name_TextBox, Link3_Name_TextBox, Link4_Name_TextBox, Link5_Name_TextBox, Link6_Name_TextBox, Link7_Name_TextBox, Link8_Name_TextBox} _
                                           , {Link2_1_Name_TextBox, Link2_2_Name_TextBox, Link2_3_Name_TextBox, Link2_4_Name_TextBox, Link2_5_Name_TextBox, Link2_6_Name_TextBox, Link2_7_Name_TextBox, Link2_8_Name_TextBox} _
                                           , {Link3_1_Name_TextBox, Link3_2_Name_TextBox, Link3_3_Name_TextBox, Link3_4_Name_TextBox, Link3_5_Name_TextBox, Link3_6_Name_TextBox, Link3_7_Name_TextBox, Link3_8_Name_TextBox} _
                                           , {Link4_1_Name_TextBox, Link4_2_Name_TextBox, Link4_3_Name_TextBox, Link4_4_Name_TextBox, Link4_5_Name_TextBox, Link4_6_Name_TextBox, Link4_7_Name_TextBox, Link4_8_Name_TextBox} _
                                           , {Link5_1_Name_TextBox, Link5_2_Name_TextBox, Link5_3_Name_TextBox, Link5_4_Name_TextBox, Link5_5_Name_TextBox, Link5_6_Name_TextBox, Link5_7_Name_TextBox, Link5_8_Name_TextBox} _
                                           , {Link6_1_Name_TextBox, Link6_2_Name_TextBox, Link6_3_Name_TextBox, Link6_4_Name_TextBox, Link6_5_Name_TextBox, Link6_6_Name_TextBox, Link6_7_Name_TextBox, Link6_8_Name_TextBox}}
        Dim LinkButton As Button(,) = {{MagicTool.Link1_1_Button, MagicTool.Link1_2_Button, MagicTool.Link1_3_Button, MagicTool.Link1_4_Button, MagicTool.Link1_5_Button, MagicTool.Link1_6_Button, MagicTool.Link1_7_Button, MagicTool.Link1_8_Button} _
                                     , {MagicTool.Link2_1_Button, MagicTool.Link2_2_Button, MagicTool.Link2_3_Button, MagicTool.Link2_4_Button, MagicTool.Link2_5_Button, MagicTool.Link2_6_Button, MagicTool.Link2_7_Button, MagicTool.Link2_8_Button} _
                                     , {MagicTool.Link3_1_Button, MagicTool.Link3_2_Button, MagicTool.Link3_3_Button, MagicTool.Link3_4_Button, MagicTool.Link3_5_Button, MagicTool.Link3_6_Button, MagicTool.Link3_7_Button, MagicTool.Link3_8_Button} _
                                     , {MagicTool.Link4_1_Button, MagicTool.Link4_2_Button, MagicTool.Link4_3_Button, MagicTool.Link4_4_Button, MagicTool.Link4_5_Button, MagicTool.Link4_6_Button, MagicTool.Link4_7_Button, MagicTool.Link4_8_Button} _
                                     , {MagicTool.Link5_1_Button, MagicTool.Link5_2_Button, MagicTool.Link5_3_Button, MagicTool.Link5_4_Button, MagicTool.Link5_5_Button, MagicTool.Link5_6_Button, MagicTool.Link5_7_Button, MagicTool.Link5_8_Button} _
                                     , {MagicTool.Link6_1_Button, MagicTool.Link6_2_Button, MagicTool.Link6_3_Button, MagicTool.Link6_4_Button, MagicTool.Link6_5_Button, MagicTool.Link6_6_Button, MagicTool.Link6_7_Button, MagicTool.Link6_8_Button}}
        Dim LinkDirTextBox As TextBox(,) = {{Link1_Dir_TextBox, Link2_Dir_TextBox, Link3_Dir_TextBox, Link4_Dir_TextBox, Link5_Dir_TextBox, Link6_Dir_TextBox, Link7_Dir_TextBox, Link8_Dir_TextBox} _
                                         , {Link2_1_Dir_TextBox, Link2_2_Dir_TextBox, Link2_3_Dir_TextBox, Link2_4_Dir_TextBox, Link2_5_Dir_TextBox, Link2_6_Dir_TextBox, Link2_7_Dir_TextBox, Link2_8_Dir_TextBox} _
                                         , {Link3_1_Dir_TextBox, Link3_2_Dir_TextBox, Link3_3_Dir_TextBox, Link3_4_Dir_TextBox, Link3_5_Dir_TextBox, Link3_6_Dir_TextBox, Link3_7_Dir_TextBox, Link3_8_Dir_TextBox} _
                                         , {Link4_1_Dir_TextBox, Link4_2_Dir_TextBox, Link4_3_Dir_TextBox, Link4_4_Dir_TextBox, Link4_5_Dir_TextBox, Link4_6_Dir_TextBox, Link4_7_Dir_TextBox, Link4_8_Dir_TextBox} _
                                         , {Link5_1_Dir_TextBox, Link5_2_Dir_TextBox, Link5_3_Dir_TextBox, Link5_4_Dir_TextBox, Link5_5_Dir_TextBox, Link5_6_Dir_TextBox, Link5_7_Dir_TextBox, Link5_8_Dir_TextBox} _
                                         , {Link6_1_Dir_TextBox, Link6_2_Dir_TextBox, Link6_3_Dir_TextBox, Link6_4_Dir_TextBox, Link6_5_Dir_TextBox, Link6_6_Dir_TextBox, Link6_7_Dir_TextBox, Link6_8_Dir_TextBox}}
        Dim CreateNewFolder_Textbox As TextBox() = {childFolder_TextBox1, childFolder_TextBox2, childFolder_TextBox3, childFolder_TextBox4, childFolder_TextBox5 _
                                                , childFolder_TextBox6}

        For i = 1 To ini_linkBtnCol
            For j = 1 To ini_linkBtnRow
                ini_Name = $"{ChangeLink_BasicString.setCont_Name_}{i}_{j}"
                WriteInini_Fun(sKeyValueNa, LinkNameTextBox(i - 1, j - 1).Text,
                               ChangeLink_BasicString.setTitle_Name, ini_Name)

                ini_LinkPath = $"{ChangeLink_BasicString.setCont_LinkPath_}{i}_{j}"
                WriteInini_Fun(sKeyValuePath, LinkDirTextBox(i - 1, j - 1).Text,
                               ChangeLink_BasicString.setTitle_LinkPath, ini_LinkPath)

                ini_NewFolder = $"{ChangeLink_BasicString.setCont_ChildFolder_}{i}"
                WriteInini_Fun(sKeyNewFolder, CreateNewFolder_Textbox(i - 1).Text,
                               ChangeLink_BasicString.setTitle_NewFolder, ini_NewFolder)
            Next
        Next



        '預設路徑 -------------------------------------------------------------------------------------------------------
        '預設路徑 > JobMaker > Load > 仕樣書
        WriteInini_Fun(sKey_getNameManager, ChgLink_DefaultPath_Spec_TextBox.Text,
                       ChangeLink_BasicString.setTitle_DefaultPath,
                       ChangeLink_BasicString.setCont_JobMaker_Spec_Path)
        ChgLink_DefaultPath_Spec_TextBox.Text = sKey_getNameManager.ToString

        '預設路徑 > JobMaker > Load > CheckList
        WriteInini_Fun(sKey_getNameManager, ChgLink_DefaultPath_CheckList_TextBox.Text,
                       ChangeLink_BasicString.setTitle_DefaultPath,
                       ChangeLink_BasicString.setCont_JobMaker_CheckList_Path)
        ChgLink_DefaultPath_CheckList_TextBox.Text = sKey_getNameManager.ToString
        '------------------------------------------------------------------------------------------------------- 預設路徑 

        '更改顏色按鈕
        WriteInini_Fun(sKey_setColor, SetLinkBtn_MouseOverColor_Button.Text,
                       ChangeLink_BasicString.setTitle_SettingColor,
                       ChangeLink_BasicString.setCont_SetLinkBtn_MouseOverColor)
        SetLinkBtn_MouseOverColor_Button.Text = sKey_setColor.ToString

        WriteInini_Fun(sKey_setColor, SetLinkBtn_FontColor_Button.Text,
                       ChangeLink_BasicString.setTitle_SettingColor,
                       ChangeLink_BasicString.setCont_SetLinkBtn_FontColor)
        SetLinkBtn_FontColor_Button.Text = sKey_setColor.ToString

        WriteInini_Fun(sKey_setColor, SetLinkBtn_Transparent_Button.Text,
                       ChangeLink_BasicString.setTitle_SettingColor,
                       ChangeLink_BasicString.setCont_SetLinkBtn_TransparentColor)
        SetLinkBtn_Transparent_Button.Text = sKey_setColor.ToString

        WriteInini_Fun(sKey_setColor, SetLinkBtn_BorderColor_Button.Text,
                       ChangeLink_BasicString.setTitle_SettingColor,
                       ChangeLink_BasicString.setCont_SetLinkBtn_BorderColor)
        SetLinkBtn_BorderColor_Button.Text = sKey_setColor.ToString

        WriteInini_Fun(sKey_setColor, SetLinkBtn_BgPicture_TextBox.Text,
                       ChangeLink_BasicString.setTitle_SettingColor,
                       ChangeLink_BasicString.setCont_SetLinkBtn_BgPicture)
        SetLinkBtn_BgPicture_TextBox.Text = sKey_setColor.ToString

        WriteInini_Fun(sKey_setColor, SetNote_BackColor_Button.Text,
                       ChangeLink_BasicString.setTitle_SettingColor,
                       ChangeLink_BasicString.setCont_SetNote_BackColor)
        SetNote_BackColor_Button.Text = sKey_setColor.ToString

        WriteInini_Fun(sKey_setColor, SetNote_FontColor_Button.Text,
                       ChangeLink_BasicString.setTitle_SettingColor,
                       ChangeLink_BasicString.setCont_SetNote_FontColor)
        SetNote_FontColor_Button.Text = sKey_setColor.ToString
        '更改顏色按鈕

        'screenChoose
        WriteInini_Fun(sKeyValueScr, ScreenChoose_CB.Text,
                       ChangeLink_BasicString.setTitle_ScreenChoose,
                       ChangeLink_BasicString.setCont_Screen)
        ScreenChoose_Label.Text = sKeyValueScr.ToString
        'screenPos
        WriteInini_Fun(sKeyValuePos, ScreenPosition_CB.Text,
                       ChangeLink_BasicString.setTitle_ScreenPos,
                       ChangeLink_BasicString.setCont_Pos)
        ScreenPos_Label.Text = sKeyValuePos.ToString
        'screenChoose
        WriteInini_Fun(sKeyValueScr, ScreenChoose_CB.Text,
                       ChangeLink_BasicString.setTitle_ScreenChoose,
                       ChangeLink_BasicString.setCont_Screen)
        ScreenChoose_Label.Text = sKeyValueScr.ToString
        'screenPos
        WriteInini_Fun(sKeyValuePos, ScreenPosition_CB.Text,
                       ChangeLink_BasicString.setTitle_ScreenPos,
                       ChangeLink_BasicString.setCont_Pos)
        ScreenPos_Label.Text = sKeyValuePos.ToString


        formPositionOnScreen_Setting(Me, sKeyValueScr.ToString, sKeyValuePos.ToString)

        'new del folder path
        combobox_state.combobox_when_save(count_allFolderCBPath, ini_FolderPath,
                                          ChangeLink_BasicString.setCont_FolderPath_,
                                          ChangeLink_BasicString.setTitle_AllFolderPath,
                                          sKeyFolderPath, AllFolderPath_ComboBox, nSize, sinifilename)
        combobox_state.combobox_when_save(count_allFolderCBName, ini_FolderName,
                                          ChangeLink_BasicString.setCont_FolderName_,
                                          ChangeLink_BasicString.setTitle_AllFolderName,
                                          sKeyFolderName, CL_AllFolderName_ComboBox, nSize, sinifilename)
        combobox_state.combobox_when_save(count_comJobCBName, ini_comJobName,
                                          ChangeLink_BasicString.setCont_CommonFile_,
                                          ChangeLink_BasicString.setTitle_CommonJobFile,
                                          sKeyComJobName, CommonJobFileName_ComboBox, nSize, sinifilename)
        combobox_state.combobox_when_save(count_comJobCBPath, ini_comJobPath,
                                          ChangeLink_BasicString.setCont_CommonFilePath_,
                                          ChangeLink_BasicString.setTitle_CommonJobFile,
                                          sKeyComJobPath, CommonJobFilePath_Combobox, nSize, sinifilename)

        '自動開啟程式方法
        auto_open_program()

        'updateINI_開啟cmd執行
        updateINI_program()
    End Sub


    ''' <summary>
    ''' 執行更新INI
    ''' </summary>
    Public Sub updateINI_program()
        Dim ini_StartPath As String
        ini_StartPath = $"{Application.StartupPath}\{ProgramAllPath.folderName_ini}"

        Dim file_count As Integer
        Dim filter_name(), allfilter_name(0) As String
        file_count = 0
        filter_name = {$"*.{ProgramAllPath.folderName_ini}"}


        If UpdateINI_CheckBox.Checked = True Then
            Try
                For Each myFilter In filter_name 'count sum and set array
                    For Each file In Directory.GetFileSystemEntries(ini_StartPath, myFilter)
                        file_count = file_count + 1

                        ReDim Preserve allfilter_name(file_count - 1)
                        allfilter_name(file_count - 1) = file
                    Next
                Next

                If file_count > 1 Then
                    Process.Start($"{ini_StartPath}\{ProgramAllName.fileName_SetFileIniBat}")

                    For i = 1 To file_count '刪除更新檔ini
                        If allfilter_name(i - 1).ToString <> sinifilename Then
                            IO.File.Delete(allfilter_name(i - 1).ToString())
                        End If
                    Next
                End If

            Catch ex As Exception
                MsgBox("自動更新ini檔案有誤", vbCritical, "ERROR")
            End Try
        End If
    End Sub
    ''' <summary>
    ''' 自動登錄檔，開機時開啟檔案寫入登入檔
    ''' </summary>
    Private Sub auto_open_program()
        Dim temp As Microsoft.Win32.RegistryKey
        If autoProgram_CheckBox.Checked = True Then
            WriteInini_Fun(sKeyValueCB_State, CStr(True),
                           ChangeLink_BasicString.setTitle_CheckBox_State,
                           ChangeLink_BasicString.setCont_autoProgram)
            '開機時開啟檔案寫入登入檔
            My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SoftWare\Microsoft\Windows\CurrentVersion\Run",
                                          "Run_MyExe_ByAuto",
                                          $"{Application.StartupPath}\{Application.ProductName}")
            temp = My.Computer.Registry.LocalMachine.OpenSubKey("software").
                OpenSubKey("microsoft").OpenSubKey("windows").OpenSubKey("currentversion").OpenSubKey("run", True)
        Else
            '刪除登錄檔
            Try
                WriteInini_Fun(sKeyValueCB_State, CStr(False),
                               ChangeLink_BasicString.setTitle_CheckBox_State,
                               ChangeLink_BasicString.setCont_autoProgram)
                temp = My.Computer.Registry.LocalMachine.OpenSubKey("software").
                    OpenSubKey("microsoft").OpenSubKey("windows").OpenSubKey("currentversion").OpenSubKey("run", True)
                temp.DeleteValue("Run_MyExe_ByAuto", True)
            Catch ex As Exception
                'MsgBox("目前沒有目標登錄檔可刪除")
            End Try
        End If

    End Sub



    '基本仕樣
    ''' <summary>
    ''' 資料夾路徑>增加按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AddFolderPath_Button_Click(sender As Object, e As EventArgs) Handles AddFolderPath_Button.Click
        If NewDelFolder_TextBox.Text <> "" Then
            If NewDelFolder_TextBox.Text = AllFolderPath_ComboBox.Text Then
                Exit Sub
            End If
            AllFolderPath_ComboBox.Items.Add(NewDelFolder_TextBox.Text)
        End If
    End Sub
    ''' <summary>
    ''' 資料夾路徑>減少按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SubFolderPath_Button_Click(sender As Object, e As EventArgs) Handles SubFolderPath_Button.Click
        If NewDelFolder_TextBox.Text = AllFolderPath_ComboBox.Text Then
            AllFolderPath_ComboBox.Items.Remove(NewDelFolder_TextBox.Text)
        End If
    End Sub
    ''' <summary>
    ''' 父資料>增加按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CL_AddFolderName_Button_Click(sender As Object, e As EventArgs) Handles CL_AddFolderName_Button.Click
        If CL_NewDelFolderName_TextBox.Text <> "" Then
            If CL_NewDelFolderName_TextBox.Text = CL_AllFolderName_ComboBox.Text Then
                Exit Sub
            End If
            CL_AllFolderName_ComboBox.Items.Add(CL_NewDelFolderName_TextBox.Text)
        End If
    End Sub
    ''' <summary>
    ''' 父資料>減少按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CL_SubFolderName_SUBButton_Click(sender As Object, e As EventArgs) Handles CL_SubFolderName_SUBButton.Click
        If CL_NewDelFolderName_TextBox.Text = CL_AllFolderName_ComboBox.Text Then
            CL_AllFolderName_ComboBox.Items.Remove(CL_NewDelFolderName_TextBox.Text)
        End If
    End Sub
    ''' <summary>
    ''' Changelink 的資料夾路徑all foldername改變時 new / del folder name 跟著改變
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AllFolderPath_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AllFolderPath_ComboBox.SelectedIndexChanged
        NewDelFolder_TextBox.Text = AllFolderPath_ComboBox.Text
    End Sub
    ''' <summary>
    ''' Changelink 的父資料夾all foldername改變時 new / del folder name 跟著改變
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CL_AllFolderName_ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CL_AllFolderName_ComboBox.SelectedIndexChanged
        CL_NewDelFolderName_TextBox.Text = CL_AllFolderName_ComboBox.Text
    End Sub
    ''' <summary>
    ''' 常用工番>增加按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CommonJob_ADDButton_Click(sender As Object, e As EventArgs) Handles CommonJob_ADDButton.Click
        If CommonJobFilePath_Combobox.Text = "" And CommonJobFileName_ComboBox.Text <> "" Then
            CommonJobFileName_ComboBox.Items.Add(CommonJobFileName_ComboBox.Text)
        ElseIf CommonJobFileName_ComboBox.Text = "" And CommonJobFilePath_Combobox.Text <> "" Then
            CommonJobFilePath_Combobox.Items.Add(CommonJobFilePath_Combobox.Text)
        ElseIf CommonJobFileName_ComboBox.Text <> "" And CommonJobFilePath_Combobox.Text <> "" Then
            CommonJobFileName_ComboBox.Items.Add(CommonJobFileName_ComboBox.Text)
            CommonJobFilePath_Combobox.Items.Add(CommonJobFilePath_Combobox.Text)
        End If
    End Sub
    ''' <summary>
    ''' 常用工番>減少按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CommonJob_SUBButton_Click(sender As Object, e As EventArgs) Handles CommonJob_SUBButton.Click
        If CommonJobFilePath_Combobox.Text = "" And CommonJobFileName_ComboBox.Text <> "" Then
            CommonJobFileName_ComboBox.Items.Remove(CommonJobFileName_ComboBox.Text)
            CommonJobFileName_ComboBox.Text = ""
        ElseIf CommonJobFileName_ComboBox.Text = "" And CommonJobFilePath_Combobox.Text <> "" Then
            CommonJobFilePath_Combobox.Items.Remove(CommonJobFilePath_Combobox.Text)
            CommonJobFilePath_Combobox.Text = ""
        ElseIf CommonJobFileName_ComboBox.Text <> "" And CommonJobFilePath_Combobox.Text <> "" Then
            CommonJobFileName_ComboBox.Items.Remove(CommonJobFileName_ComboBox.Text)
            CommonJobFileName_ComboBox.Text = ""

            CommonJobFilePath_Combobox.Items.Remove(CommonJobFilePath_Combobox.Text)
            CommonJobFilePath_Combobox.Text = ""
        End If
    End Sub
    ''' <summary>
    ''' 動態增加childFolder選項
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub childFolder_ADDButton_Click(sender As Object, e As EventArgs) Handles childFolder_ADDButton.Click
        Dim label_ABC As New Label
        Dim textbox_ABC As New TextBox
        Dim number As Integer = 1
        label_ABC.Location = New Point(number * 40, 10)
        label_ABC.Text = "label" + number.ToString + label_ABC.Location.ToString
        childFolder_FlowLayoutPanel.Controls.Add(label_ABC)
        textbox_ABC.Location = New Point(number * 40, 20)
        textbox_ABC.Text = "textbox" + number.ToString
        childFolder_FlowLayoutPanel.Controls.Add(textbox_ABC)
    End Sub
    '打開Dialog視窗 ----------------------------------------------------------------------
    ''' <summary>
    ''' 打開Dialog視窗後寫入Control
    ''' </summary>
    ''' <param name="Dir"></param>
    Public Overloads Sub OpenFilePath_event(Dir As Control)
        Dim myFolderBrowserDialog As New FolderBrowserDialog()
        myFolderBrowserDialog.SelectedPath = Dir.Text

        If myFolderBrowserDialog.ShowDialog() = DialogResult.OK Then
            Dir.Text = myFolderBrowserDialog.SelectedPath
        End If
    End Sub
    ''' <summary>
    ''' 以預設路徑打開Dialog視窗後寫入Control
    ''' </summary>
    ''' <param name="Dir"></param>
    Public Overloads Sub OpenFilePath_event(defaultPath As String, Dir As Control)
        Dim myFolderBrowserDialog As New FolderBrowserDialog()
        myFolderBrowserDialog.SelectedPath = defaultPath

        If myFolderBrowserDialog.ShowDialog() = DialogResult.OK Then
            Dir.Text = myFolderBrowserDialog.SelectedPath
        End If
    End Sub

    Enum OpenFileType
        mExcel
        mOther
    End Enum
    ''' <summary>
    ''' 打開Dialog視窗，選擇Excel格式檔案並寫入路徑
    ''' </summary>
    ''' <param name="Dir">目標路徑含檔案名稱</param>
    Public Overloads Sub OpenFile_event(Dir As TextBox,
                                        fileType As OpenFileType,
                                        defalutPath As String)
        Dim result As DialogResult
        Dim myFileBrowserDialog As New OpenFileDialog()
        myFileBrowserDialog.InitialDirectory = defalutPath '開啟預設路徑
        If fileType = OpenFileType.mExcel Then
            myFileBrowserDialog.Filter = "Excel(*.xlsx,*.xls,*.xlsm)|*.xlsx;*.xls;*.xlsm|All files(*.*)|*.*"
        ElseIf fileType = OpenFileType.mOther Then
            myFileBrowserDialog.Filter = "All files(*.*)|*.*"
        End If

        result = myFileBrowserDialog.ShowDialog()
        If result = DialogResult.OK Then
            Dir.Text = myFileBrowserDialog.FileName
        End If
    End Sub
    ''' <summary>
    ''' 打開Dialog視窗，選擇Excel格式檔案並寫入路徑
    ''' </summary>
    ''' <param name="Dir">目標路徑含檔案名稱</param>
    Public Overloads Sub OpenFile_event(Dir As ComboBox,
                                        fileType As OpenFileType,
                                        defalutPath As String)
        Dim result As DialogResult
        Dim myFileBrowserDialog As New OpenFileDialog()
        myFileBrowserDialog.InitialDirectory = defalutPath '開啟預設路徑
        If fileType = OpenFileType.mExcel Then
            myFileBrowserDialog.Filter = "Excel(*.xlsx,*.xls,*.xlsm)|*.xlsx;*.xls;*.xlsm|All files(*.*)|*.*"
        ElseIf fileType = OpenFileType.mOther Then
            myFileBrowserDialog.Filter = "All files(*.*)|*.*"
        End If

        result = myFileBrowserDialog.ShowDialog()
        If result = DialogResult.OK Then
            Dir.Text = myFileBrowserDialog.FileName
        End If
    End Sub
    '----------------------------------------------------------------------打開Dialog視窗 

    '預設路徑 ---------------------------------------------------------------------------
    ''' <summary>
    ''' 仕樣書
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChgLink_DefaultPath_Spec_Button_Click(sender As Object, e As EventArgs) Handles ChgLink_DefaultPath_Spec_Button.Click
        OpenFilePath_event(ChgLink_DefaultPath_Spec_TextBox)
    End Sub
    ''' <summary>
    ''' CheckList
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ChgLink_DefaultPath_CheckList_Button_Click(sender As Object, e As EventArgs) Handles ChgLink_DefaultPath_CheckList_Button.Click
        OpenFilePath_event(ChgLink_DefaultPath_CheckList_TextBox)
    End Sub
    '---------------------------------------------------------------------------預設路徑 


    Private Sub Link1_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link1_1_OpenFile_Button.Click
        OpenFilePath_event(Link1_Dir_TextBox)
    End Sub
    Private Sub Link2_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link1_2_OpenFile_Button.Click
        OpenFilePath_event(Link2_Dir_TextBox)
    End Sub
    Private Sub Link3_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link1_3_OpenFile_Button.Click
        OpenFilePath_event(Link3_Dir_TextBox)
    End Sub
    Private Sub Link4_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link1_4_OpenFile_Button.Click
        OpenFilePath_event(Link4_Dir_TextBox)
    End Sub
    Private Sub Link5_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link1_5_OpenFile_Button.Click
        OpenFilePath_event(Link5_Dir_TextBox)
    End Sub
    Private Sub Link6_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link1_6_OpenFile_Button.Click
        OpenFilePath_event(Link6_Dir_TextBox)
    End Sub
    Private Sub Link7_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link1_7_OpenFile_Button.Click
        OpenFilePath_event(Link7_Dir_TextBox)
    End Sub
    Private Sub Link8_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link1_8_OpenFile_Button.Click
        OpenFilePath_event(Link8_Dir_TextBox)
    End Sub
    Private Sub Link21_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link2_1_OpenFile_Button.Click
        OpenFilePath_event(Link2_1_Dir_TextBox)
    End Sub
    Private Sub Link22_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link2_2_OpenFile_Button.Click
        OpenFilePath_event(Link2_2_Dir_TextBox)
    End Sub
    Private Sub Link23_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link2_3_OpenFile_Button.Click
        OpenFilePath_event(Link2_3_Dir_TextBox)
    End Sub
    Private Sub Link24_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link2_4_OpenFile_Button.Click
        OpenFilePath_event(Link2_4_Dir_TextBox)
    End Sub
    Private Sub Link25_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link2_5_OpenFile_Button.Click
        OpenFilePath_event(Link2_5_Dir_TextBox)
    End Sub
    Private Sub Link26_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link2_6_OpenFile_Button.Click
        OpenFilePath_event(Link2_6_Dir_TextBox)
    End Sub
    Private Sub Link27_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link2_7_OpenFile_Button.Click
        OpenFilePath_event(Link2_7_Dir_TextBox)
    End Sub
    Private Sub Link28_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link2_8_OpenFile_Button.Click
        OpenFilePath_event(Link2_8_Dir_TextBox)
    End Sub
    Private Sub Link31_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link3_1_OpenFile_Button.Click
        OpenFilePath_event(Link3_1_Dir_TextBox)
    End Sub
    Private Sub Link32_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link3_2_OpenFile_Button.Click
        OpenFilePath_event(Link3_2_Dir_TextBox)
    End Sub
    Private Sub Link33_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link3_3_OpenFile_Button.Click
        OpenFilePath_event(Link3_3_Dir_TextBox)
    End Sub
    Private Sub Link34_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link3_4_OpenFile_Button.Click
        OpenFilePath_event(Link3_4_Dir_TextBox)
    End Sub
    Private Sub Link35_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link3_5_OpenFile_Button.Click
        OpenFilePath_event(Link3_5_Dir_TextBox)
    End Sub
    Private Sub Link36_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link3_6_OpenFile_Button.Click
        OpenFilePath_event(Link3_6_Dir_TextBox)
    End Sub
    Private Sub Link37_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link3_7_OpenFile_Button.Click
        OpenFilePath_event(Link3_7_Dir_TextBox)
    End Sub
    Private Sub Link38_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link3_8_OpenFile_Button.Click
        OpenFilePath_event(Link3_8_Dir_TextBox)
    End Sub
    Private Sub Link41_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link4_1_OpenFile_Button.Click
        OpenFilePath_event(Link4_1_Dir_TextBox)
    End Sub
    Private Sub Link42_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link4_2_OpenFile_Button.Click
        OpenFilePath_event(Link4_2_Dir_TextBox)
    End Sub
    Private Sub Link43_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link4_3_OpenFile_Button.Click
        OpenFilePath_event(Link4_3_Dir_TextBox)
    End Sub
    Private Sub Link44_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link4_4_OpenFile_Button.Click
        OpenFilePath_event(Link4_4_Dir_TextBox)
    End Sub
    Private Sub Link45_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link4_5_OpenFile_Button.Click
        OpenFilePath_event(Link4_5_Dir_TextBox)
    End Sub
    Private Sub Link46_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link4_6_OpenFile_Button.Click
        OpenFilePath_event(Link4_6_Dir_TextBox)
    End Sub
    Private Sub Link47_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link4_7_OpenFile_Button.Click
        OpenFilePath_event(Link4_7_Dir_TextBox)
    End Sub
    Private Sub Link48_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link4_8_OpenFile_Button.Click
        OpenFilePath_event(Link4_8_Dir_TextBox)
    End Sub
    Private Sub Link51_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link5_1_OpenFile_Button.Click
        OpenFilePath_event(Link5_1_Dir_TextBox)
    End Sub
    Private Sub Link52_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link5_2_OpenFile_Button.Click
        OpenFilePath_event(Link5_2_Dir_TextBox)
    End Sub
    Private Sub Link53_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link5_3_OpenFile_Button.Click
        OpenFilePath_event(Link5_3_Dir_TextBox)
    End Sub
    Private Sub Link54_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link5_4_OpenFile_Button.Click
        OpenFilePath_event(Link5_4_Dir_TextBox)
    End Sub
    Private Sub Link55_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link5_5_OpenFile_Button.Click
        OpenFilePath_event(Link5_5_Dir_TextBox)
    End Sub
    Private Sub Link56_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link5_6_OpenFile_Button.Click
        OpenFilePath_event(Link5_6_Dir_TextBox)
    End Sub
    Private Sub Link57_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link5_7_OpenFile_Button.Click
        OpenFilePath_event(Link5_7_Dir_TextBox)
    End Sub
    Private Sub Link58_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link5_8_OpenFile_Button.Click
        OpenFilePath_event(Link5_8_Dir_TextBox)
    End Sub
    Private Sub Link61_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link6_1_OpenFile_Button.Click
        OpenFilePath_event(Link6_1_Dir_TextBox)
    End Sub
    Private Sub Link62_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link6_2_OpenFile_Button.Click
        OpenFilePath_event(Link6_2_Dir_TextBox)
    End Sub
    Private Sub Link63_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link6_3_OpenFile_Button.Click
        OpenFilePath_event(Link6_3_Dir_TextBox)
    End Sub
    Private Sub Link64_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link6_4_OpenFile_Button.Click
        OpenFilePath_event(Link6_4_Dir_TextBox)
    End Sub
    Private Sub Link65_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link6_5_OpenFile_Button.Click
        OpenFilePath_event(Link6_5_Dir_TextBox)
    End Sub
    Private Sub Link66_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link6_6_OpenFile_Button.Click
        OpenFilePath_event(Link6_6_Dir_TextBox)
    End Sub
    Private Sub Link67_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link6_7_OpenFile_Button.Click
        OpenFilePath_event(Link6_7_Dir_TextBox)
    End Sub
    Private Sub Link68_OpenFile_Button_Click(sender As Object, e As EventArgs) Handles Link6_8_OpenFile_Button.Click
        OpenFilePath_event(Link6_8_Dir_TextBox)
    End Sub
    'OpenFile_Dir
    Private Sub Link1_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link1_1_CheckBox.CheckedChanged
        IfCB_Click(Link1_1_CheckBox, MagicTool.Link1_1_Button,
                   Link1_Name_TextBox, Link1_Dir_TextBox,
                   Link1_1_OpenFile_Button, False)
    End Sub
    Private Sub Link2_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link1_2_CheckBox.CheckedChanged
        IfCB_Click(Link1_2_CheckBox, MagicTool.Link1_2_Button,
                   Link2_Name_TextBox, Link2_Dir_TextBox,
                   Link1_2_OpenFile_Button, False)
    End Sub
    Private Sub Link3_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link1_3_CheckBox.CheckedChanged
        IfCB_Click(Link1_3_CheckBox, MagicTool.Link1_3_Button,
                   Link3_Name_TextBox, Link3_Dir_TextBox,
                   Link1_3_OpenFile_Button, False)
    End Sub
    Private Sub Link4_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link1_4_CheckBox.CheckedChanged
        IfCB_Click(Link1_4_CheckBox, MagicTool.Link1_4_Button,
                   Link4_Name_TextBox, Link4_Dir_TextBox,
                   Link1_4_OpenFile_Button, False)
    End Sub
    Private Sub Link5_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link1_5_CheckBox.CheckedChanged
        IfCB_Click(Link1_5_CheckBox, MagicTool.Link1_5_Button,
                   Link5_Name_TextBox, Link5_Dir_TextBox,
                   Link1_5_OpenFile_Button, False)
    End Sub
    Private Sub Link6_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link1_6_CheckBox.CheckedChanged
        IfCB_Click(Link1_6_CheckBox, MagicTool.Link1_6_Button,
                   Link6_Name_TextBox, Link6_Dir_TextBox,
                   Link1_6_OpenFile_Button, False)
    End Sub
    Private Sub Link7_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link1_7_CheckBox.CheckedChanged
        IfCB_Click(Link1_7_CheckBox, MagicTool.Link1_7_Button,
                   Link7_Name_TextBox, Link7_Dir_TextBox,
                   Link1_7_OpenFile_Button, False)
    End Sub
    Private Sub Link8_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link1_8_CheckBox.CheckedChanged
        IfCB_Click(Link1_8_CheckBox, MagicTool.Link1_8_Button,
                   Link8_Name_TextBox, Link8_Dir_TextBox,
                   Link1_8_OpenFile_Button, False)
    End Sub
    Private Sub Link2_1_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link2_1_CheckBox.CheckedChanged
        IfCB_Click(Link2_1_CheckBox, MagicTool.Link2_1_Button,
                   Link2_1_Name_TextBox, Link2_1_Dir_TextBox,
                   Link2_1_OpenFile_Button, False)
    End Sub
    Private Sub Link2_2_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link2_2_CheckBox.CheckedChanged
        IfCB_Click(Link2_2_CheckBox, MagicTool.Link2_2_Button,
                   Link2_2_Name_TextBox, Link2_2_Dir_TextBox,
                   Link2_2_OpenFile_Button, False)
    End Sub
    Private Sub Link2_3_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link2_3_CheckBox.CheckedChanged
        IfCB_Click(Link2_3_CheckBox, MagicTool.Link2_3_Button,
                   Link2_3_Name_TextBox, Link2_3_Dir_TextBox,
                   Link2_3_OpenFile_Button, False)
    End Sub
    Private Sub Link2_4_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link2_4_CheckBox.CheckedChanged
        IfCB_Click(Link2_4_CheckBox, MagicTool.Link2_4_Button,
                   Link2_4_Name_TextBox, Link2_4_Dir_TextBox,
                   Link2_4_OpenFile_Button, False)
    End Sub
    Private Sub Link2_5_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link2_5_CheckBox.CheckedChanged
        IfCB_Click(Link2_5_CheckBox, MagicTool.Link2_5_Button,
                   Link2_5_Name_TextBox, Link2_5_Dir_TextBox,
                   Link2_5_OpenFile_Button, False)
    End Sub
    Private Sub Link2_6_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link2_6_CheckBox.CheckedChanged
        IfCB_Click(Link2_6_CheckBox, MagicTool.Link2_6_Button,
                   Link2_6_Name_TextBox, Link2_6_Dir_TextBox,
                   Link2_6_OpenFile_Button, False)
    End Sub
    Private Sub Link2_7_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link2_7_CheckBox.CheckedChanged
        IfCB_Click(Link2_7_CheckBox, MagicTool.Link2_7_Button,
                   Link2_7_Name_TextBox, Link2_7_Dir_TextBox,
                   Link2_7_OpenFile_Button, False)
    End Sub
    Private Sub Link2_8_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link2_8_CheckBox.CheckedChanged
        IfCB_Click(Link2_8_CheckBox, MagicTool.Link2_8_Button,
                   Link2_8_Name_TextBox, Link2_8_Dir_TextBox,
                   Link2_8_OpenFile_Button, False)
    End Sub
    Private Sub Link3_1_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link3_1_CheckBox.CheckedChanged
        IfCB_Click(Link3_1_CheckBox, MagicTool.Link3_1_Button,
                   Link3_1_Name_TextBox, Link3_1_Dir_TextBox,
                   Link3_1_OpenFile_Button, False)
    End Sub
    Private Sub Link3_2_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link3_2_CheckBox.CheckedChanged
        IfCB_Click(Link3_2_CheckBox, MagicTool.Link3_2_Button,
                   Link3_2_Name_TextBox, Link3_2_Dir_TextBox,
                   Link3_2_OpenFile_Button, False)
    End Sub
    Private Sub Link3_3_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link3_3_CheckBox.CheckedChanged
        IfCB_Click(Link3_3_CheckBox, MagicTool.Link3_3_Button,
                   Link3_3_Name_TextBox, Link3_3_Dir_TextBox,
                   Link3_3_OpenFile_Button, False)
    End Sub
    Private Sub Link3_4_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link3_4_CheckBox.CheckedChanged
        IfCB_Click(Link3_4_CheckBox, MagicTool.Link3_4_Button,
                   Link3_4_Name_TextBox, Link3_4_Dir_TextBox,
                   Link3_4_OpenFile_Button, False)
    End Sub
    Private Sub Link3_5_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link3_5_CheckBox.CheckedChanged
        IfCB_Click(Link3_5_CheckBox, MagicTool.Link3_5_Button,
                   Link3_5_Name_TextBox, Link3_5_Dir_TextBox,
                   Link3_5_OpenFile_Button, False)
    End Sub
    Private Sub Link3_6_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link3_6_CheckBox.CheckedChanged
        IfCB_Click(Link3_6_CheckBox, MagicTool.Link3_6_Button,
                   Link3_6_Name_TextBox, Link3_6_Dir_TextBox,
                   Link3_6_OpenFile_Button, False)
    End Sub
    Private Sub Link3_7_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link3_7_CheckBox.CheckedChanged
        IfCB_Click(Link3_7_CheckBox, MagicTool.Link3_7_Button,
                   Link3_7_Name_TextBox, Link3_7_Dir_TextBox,
                   Link3_7_OpenFile_Button, False)
    End Sub
    Private Sub Link3_8_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link3_8_CheckBox.CheckedChanged
        IfCB_Click(Link3_8_CheckBox, MagicTool.Link3_8_Button,
                   Link3_8_Name_TextBox, Link3_8_Dir_TextBox,
                   Link3_8_OpenFile_Button, False)
    End Sub
    Private Sub Link4_1_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link4_1_CheckBox.CheckedChanged
        IfCB_Click(Link4_1_CheckBox, MagicTool.Link4_1_Button,
                   Link4_1_Name_TextBox, Link4_1_Dir_TextBox,
                   Link4_1_OpenFile_Button, False)
    End Sub
    Private Sub Link4_2_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link4_2_CheckBox.CheckedChanged
        IfCB_Click(Link4_2_CheckBox, MagicTool.Link4_2_Button,
                   Link4_2_Name_TextBox, Link4_2_Dir_TextBox,
                   Link4_2_OpenFile_Button, False)
    End Sub
    Private Sub Link4_3_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link4_3_CheckBox.CheckedChanged
        IfCB_Click(Link4_3_CheckBox, MagicTool.Link4_3_Button,
                   Link4_3_Name_TextBox, Link4_3_Dir_TextBox,
                   Link4_3_OpenFile_Button, False)
    End Sub
    Private Sub Link4_4_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link4_4_CheckBox.CheckedChanged
        IfCB_Click(Link4_4_CheckBox, MagicTool.Link4_4_Button,
                   Link4_4_Name_TextBox, Link4_4_Dir_TextBox,
                   Link4_4_OpenFile_Button, False)
    End Sub
    Private Sub Link4_5_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link4_5_CheckBox.CheckedChanged
        IfCB_Click(Link4_5_CheckBox, MagicTool.Link4_5_Button,
                   Link4_5_Name_TextBox, Link4_5_Dir_TextBox,
                   Link4_5_OpenFile_Button, False)
    End Sub
    Private Sub Link4_6_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link4_6_CheckBox.CheckedChanged
        IfCB_Click(Link4_6_CheckBox, MagicTool.Link4_6_Button,
                   Link4_6_Name_TextBox, Link4_6_Dir_TextBox,
                   Link4_6_OpenFile_Button, False)
    End Sub
    Private Sub Link4_7_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link4_7_CheckBox.CheckedChanged
        IfCB_Click(Link4_7_CheckBox, MagicTool.Link4_7_Button,
                   Link4_7_Name_TextBox, Link4_7_Dir_TextBox,
                   Link4_7_OpenFile_Button, False)
    End Sub
    Private Sub Link4_8_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link4_8_CheckBox.CheckedChanged
        IfCB_Click(Link4_8_CheckBox, MagicTool.Link4_8_Button,
                   Link4_8_Name_TextBox, Link4_8_Dir_TextBox,
                   Link4_8_OpenFile_Button, False)
    End Sub
    Private Sub Link5_1_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link5_1_CheckBox.CheckedChanged
        IfCB_Click(Link5_1_CheckBox, MagicTool.Link5_1_Button,
                   Link5_1_Name_TextBox, Link5_1_Dir_TextBox,
                   Link5_1_OpenFile_Button, False)
    End Sub
    Private Sub Link5_2_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link5_2_CheckBox.CheckedChanged
        IfCB_Click(Link5_2_CheckBox, MagicTool.Link5_2_Button,
                   Link5_2_Name_TextBox, Link5_2_Dir_TextBox,
                   Link5_2_OpenFile_Button, False)
    End Sub
    Private Sub Link5_3_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link5_3_CheckBox.CheckedChanged
        IfCB_Click(Link5_3_CheckBox, MagicTool.Link5_3_Button,
                   Link5_3_Name_TextBox, Link5_3_Dir_TextBox,
                   Link5_3_OpenFile_Button, False)
    End Sub
    Private Sub Link5_4_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link5_4_CheckBox.CheckedChanged
        IfCB_Click(Link5_4_CheckBox, MagicTool.Link5_4_Button,
                   Link5_4_Name_TextBox, Link5_4_Dir_TextBox,
                   Link5_4_OpenFile_Button, False)
    End Sub
    Private Sub Link5_5_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link5_5_CheckBox.CheckedChanged
        IfCB_Click(Link5_5_CheckBox, MagicTool.Link5_5_Button,
                   Link5_5_Name_TextBox, Link5_5_Dir_TextBox,
                   Link5_5_OpenFile_Button, False)
    End Sub
    Private Sub Link5_6_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link5_6_CheckBox.CheckedChanged
        IfCB_Click(Link5_6_CheckBox, MagicTool.Link5_6_Button,
                   Link5_6_Name_TextBox, Link5_6_Dir_TextBox,
                   Link5_6_OpenFile_Button, False)
    End Sub
    Private Sub Link5_7_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link5_7_CheckBox.CheckedChanged
        IfCB_Click(Link5_7_CheckBox, MagicTool.Link5_7_Button,
                   Link5_7_Name_TextBox, Link5_7_Dir_TextBox,
                   Link5_7_OpenFile_Button, False)
    End Sub
    Private Sub Link5_8_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link5_8_CheckBox.CheckedChanged
        IfCB_Click(Link5_8_CheckBox, MagicTool.Link5_8_Button,
                   Link5_8_Name_TextBox, Link5_8_Dir_TextBox,
                   Link5_8_OpenFile_Button, False)
    End Sub

    Private Sub Link6_1_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link6_1_CheckBox.CheckedChanged
        IfCB_Click(Link6_1_CheckBox, MagicTool.Link6_1_Button,
                   Link6_1_Name_TextBox, Link6_1_Dir_TextBox,
                   Link6_1_OpenFile_Button, False)
    End Sub
    Private Sub Link6_2_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link6_2_CheckBox.CheckedChanged
        IfCB_Click(Link6_2_CheckBox, MagicTool.Link6_2_Button,
                   Link6_2_Name_TextBox, Link6_2_Dir_TextBox,
                   Link6_2_OpenFile_Button, False)
    End Sub
    Private Sub Link6_3_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link6_3_CheckBox.CheckedChanged
        IfCB_Click(Link6_3_CheckBox, MagicTool.Link6_3_Button,
                   Link6_3_Name_TextBox, Link6_3_Dir_TextBox,
                   Link6_3_OpenFile_Button, False)
    End Sub
    Private Sub Link6_4_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link6_4_CheckBox.CheckedChanged
        IfCB_Click(Link6_4_CheckBox, MagicTool.Link6_4_Button,
                   Link6_4_Name_TextBox, Link6_4_Dir_TextBox,
                   Link6_4_OpenFile_Button, False)
    End Sub
    Private Sub Link6_5_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link6_5_CheckBox.CheckedChanged
        IfCB_Click(Link6_5_CheckBox, MagicTool.Link6_5_Button,
                   Link6_5_Name_TextBox, Link6_5_Dir_TextBox,
                   Link6_5_OpenFile_Button, False)
    End Sub

    Private Sub Link6_6_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link6_6_CheckBox.CheckedChanged
        IfCB_Click(Link6_6_CheckBox, MagicTool.Link6_6_Button,
                   Link6_6_Name_TextBox, Link6_6_Dir_TextBox,
                   Link6_6_OpenFile_Button, False)
    End Sub

    Private Sub Link6_7_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link6_7_CheckBox.CheckedChanged
        IfCB_Click(Link6_7_CheckBox, MagicTool.Link6_7_Button,
                   Link6_7_Name_TextBox, Link6_7_Dir_TextBox,
                   Link6_7_OpenFile_Button, False)
    End Sub
    Private Sub Link6_8_CheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles Link6_8_CheckBox.CheckedChanged
        IfCB_Click(Link6_8_CheckBox, MagicTool.Link6_8_Button,
                   Link6_8_Name_TextBox, Link6_8_Dir_TextBox,
                   Link6_8_OpenFile_Button, False)
    End Sub



    '設定顏色btn
    Private Sub SetLinkBtn_MouseOverColor_Button_Click(sender As Object, e As EventArgs) Handles SetLinkBtn_MouseOverColor_Button.Click
        '更換LinkBtn反白
        If ColorDialog_ForBtn.ShowDialog <> DialogResult.Cancel Then
            SetLinkBtn_MouseOverColor_Button.Text = ColorTranslator.ToHtml(ColorDialog_ForBtn.Color)
            SetLinkBtn_Result_Button.FlatAppearance.MouseOverBackColor = ColorDialog_ForBtn.Color
        End If
    End Sub
    Private Sub SetLinkBtn_FontColor_Button_Click(sender As Object, e As EventArgs) Handles SetLinkBtn_FontColor_Button.Click
        '更換LinkBtn文字顏色
        If ColorDialog_ForBtn.ShowDialog <> DialogResult.Cancel Then
            SetLinkBtn_FontColor_Button.Text = ColorTranslator.ToHtml(ColorDialog_ForBtn.Color)
            SetLinkBtn_Result_Button.ForeColor = ColorDialog_ForBtn.Color
        End If
    End Sub
    Private Sub SetLinkBtn_Transparent_Button_Click(sender As Object, e As EventArgs) Handles SetLinkBtn_Transparent_Button.Click
        '是否LinkBtn透明度?
        If linkBtn_isTransarent Then
            SetLinkBtn_Transparent_Button.Text = "NO"
            linkBtn_isTransarent = False
            SetLinkBtn_Result_Button.BackColor = DefaultBackColor
        Else
            SetLinkBtn_Transparent_Button.Text = "YES"
            linkBtn_isTransarent = True
            SetLinkBtn_Result_Button.BackColor = Color.Transparent
        End If
    End Sub

    Private Sub SetLinkBtn_BorderColor_Button_Click(sender As Object, e As EventArgs) Handles SetLinkBtn_BorderColor_Button.Click
        '更換LinkBtn邊線顏色
        If ColorDialog_ForBtn.ShowDialog <> DialogResult.Cancel Then
            SetLinkBtn_BorderColor_Button.Text = ColorTranslator.ToHtml(ColorDialog_ForBtn.Color)
            SetLinkBtn_Result_Button.FlatAppearance.BorderColor = ColorDialog_ForBtn.Color
        End If
    End Sub

    Private Sub SetLinkBtn_BgPicture_Button_Click(sender As Object, e As EventArgs) Handles SetLinkBtn_BgPicture_Button.Click
        '更換LinkBtn背景圖片
        Dim result As DialogResult
        Dim fileDialog As New OpenFileDialog
        result = fileDialog.ShowDialog

        If result = DialogResult.OK Then
            SetLinkBtn_BgPicture_TextBox.Text = fileDialog.FileName
        End If
    End Sub

    Private Sub SetNote_BackColor_Button_Click(sender As Object, e As EventArgs) Handles SetNote_BackColor_Button.Click
        '更換LinkBtn背景顏色
        If ColorDialog_ForBtn.ShowDialog <> DialogResult.Cancel Then
            SetNote_BackColor_Button.Text = ColorTranslator.ToHtml(ColorDialog_ForBtn.Color)
            SetNote_Result_TextBox.BackColor = ColorDialog_ForBtn.Color
        End If
    End Sub

    Private Sub SetNote_FontColor_Button_Click(sender As Object, e As EventArgs) Handles SetNote_FontColor_Button.Click
        '更換LinkBtn文字顏色
        If ColorDialog_ForBtn.ShowDialog <> DialogResult.Cancel Then
            SetNote_FontColor_Button.Text = ColorTranslator.ToHtml(ColorDialog_ForBtn.Color)
            SetNote_Result_TextBox.ForeColor = ColorDialog_ForBtn.Color
        End If
    End Sub

    '設定顏色btn
    Private Sub exit_Button_Click(sender As Object, e As EventArgs) Handles exit_Button.Click
        Me.Close()
        MagicTool.loadIni_form_changLink()
    End Sub

    Private Sub TestScr_button_Click(sender As Object, e As EventArgs) Handles TestScr_button.Click
        Dim index As Integer
        Dim upperbound As Integer

        Dim screnns() As Screen = Screen.AllScreens
        upperbound = screnns.GetUpperBound(0)

        TextBox1.Multiline = True
        TextBox1.Dock = DockStyle.Fill
        TextBox1.Clear()
        TextBox1.AppendText("桌面表單大小及位置 :" & vbTab & Me.DesktopBounds.ToString & vbCrLf)
        TextBox1.AppendText("螢幕數量 :" & vbTab & screnns.Count & vbCrLf)

        For index = 0 To upperbound
            TextBox1.AppendText(vbCrLf)
            TextBox1.AppendText("裝置名稱 :" & vbTab & screnns(index).DeviceName & vbCrLf)
            TextBox1.AppendText("顯示界線 :" & vbTab & screnns(index).Bounds.ToString() & vbCrLf)
            TextBox1.AppendText("類型 :" & vbTab & screnns(index).GetType().ToString() & vbCrLf)
            TextBox1.AppendText("工作區域 :" & vbTab & screnns(index).WorkingArea.ToString() & vbCrLf)
            TextBox1.AppendText("是否為主螢幕 :" & vbTab & screnns(index).Primary.ToString() & vbCrLf)
        Next
    End Sub
End Class