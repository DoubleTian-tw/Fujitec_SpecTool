Imports System.Runtime.InteropServices 'MarshalAs
Imports System.Text 'StringBuilder
Imports System.Windows.Forms.SystemInformation 'WorkingArea

Module SettingLink_modINI
    '* lpAppName：指向包含Section 名稱的字符串地址
    '* lpKeyName：指向包含Key 名稱的字符串地址
    '* lpDefault：如果Key 值沒有找到，缺省返回缺省的字符串
    '* lpReturnedString：用於保存返回字符串的緩衝區
    '* nSize： 緩衝區的長度
    '* lpFileName ：ini 文件的文件名
    Public Declare Function GetPrivateProfileString_class Lib "kernel32" Alias "GetPrivateProfileStringA" (
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpApplicationName As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpDefault As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder, ByVal nSize As UInt32,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As UInt32

    Public Declare Function WritePrivateProfileString_class Lib "kernel32" Alias "WritePrivateProfileStringA" (
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpApplicationName As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpKeyName As String,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpReturnedString As StringBuilder,
     <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String) As UInt32

    Public Declare Ansi Function FlushPrivateProfileString_class Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (
                                      ByVal lpApplicationName As Integer,
                                      ByVal lpKeyName As Integer,
                                      ByVal lpString As Integer,
                                      ByVal lpFileName As String) As Integer

End Module


Public Class SettingLinkOnChangeLink
    'ini
    Public sKeyValue, sKeyValue_note, sKeyValueCB_state, sKeyValueCB, sKeyValueNa, sKeyValuePath, sKeyValueNewFolder _
        , sKeyValueScrCho, sKeyValueScrPos, sKeyValueAllBos, sKeyFolderPath, sKeyFolderName, sKeyComJobPath, sKeyAllEmp As New StringBuilder(512)
    Dim nSize As UInt32 = Convert.ToUInt32(1024)
    Dim sinifilename As String = Application.StartupPath & "\ini\SetFile.ini"

    Dim ini_a As Integer = 6
    Dim ini_b As Integer = 8 '1-1~1-8 ...6-1~6-8
    Dim ini_LCB, ini_Name, ini_LinkPath, ini_NewFolder, ini_FolderPath, ini_FolderName, ini_comJobName _
        , ini_comJobPath, ini_allEmp, ini_allBos As String

    'count combobox
    Dim count_allFolderCBPath As Integer = 0 'count all folder path sum
    Dim count_allFolderCBName As Integer = 0 'count all folder name sum
    Dim count_comJobCBPath As Integer = 0 'count common job path
    Dim count_comJobCBName As Integer = 0 'count common job file name
    Dim count_AllEmployee As Integer = 0 '軟體帥哥項目初始化

    Dim combobox_state As controlStateOnChangeLink = New controlStateOnChangeLink()


    '初始化寫入ini的值
    Public Sub Initialization_ini()
        'Note content
        'GetPrivateProfileString_class("Note", "Content", "", sKeyValue_note, nSize, sinifilename)
        'MagicTool.note_TextBox.Text = sKeyValue_note.ToString
        'Note save always
        GetPrivateProfileString_class("CheckBox_State", "NoteSave", "", sKeyValueCB_state, nSize, sinifilename)
        ChangeLink.note_CheckBox.Checked = sKeyValueCB_state.ToString
        'Topmost
        GetPrivateProfileString_class("CheckBox_State", "TopmostSet", "", sKeyValueCB_state, nSize, sinifilename)
        ChangeLink.Topmost_CheckBox.Checked = sKeyValueCB_state.ToString
        'autoProgram
        GetPrivateProfileString_class("CheckBox_State2", "autoProgram", "", sKeyValueCB_state, nSize, sinifilename)
        ChangeLink.autoProgram_CheckBox.Checked = sKeyValueCB_state.ToString
        'updateINI
        GetPrivateProfileString_class("CheckBox_State2", "UpdateINI", "", sKeyValueCB_state, nSize, sinifilename)
        ChangeLink.UpdateINI_CheckBox.Checked = sKeyValueCB_state.ToString
        'backupNotice
        GetPrivateProfileString_class("CheckBox_State2", "backupNotice", "", sKeyValueCB_state, nSize, sinifilename)
        ChangeLink.Backup_Notice_CheckBox.Checked = sKeyValueCB_state.ToString
        'ScrChoose
        GetPrivateProfileString_class("ScreenChoose", "Screen", "", sKeyValueScrCho, nSize, sinifilename)
        ChangeLink.ScreenChoose_Label.Text = sKeyValueScrCho.ToString
        'ScrPos
        GetPrivateProfileString_class("ScreenPos", "Pos", "", sKeyValueScrPos, nSize, sinifilename)
        ChangeLink.ScreenPos_Label.Text = sKeyValueScrPos.ToString

        'link
        Dim LinkCheckBox As CheckBox(,) = {{ChangeLink.Link1_CheckBox, ChangeLink.Link2_CheckBox, ChangeLink.Link3_CheckBox, ChangeLink.Link4_CheckBox, ChangeLink.Link5_CheckBox, ChangeLink.Link6_CheckBox, ChangeLink.Link7_CheckBox, ChangeLink.Link8_CheckBox} _
                                        , {ChangeLink.Link2_1_CheckBox, ChangeLink.Link2_2_CheckBox, ChangeLink.Link2_3_CheckBox, ChangeLink.Link2_4_CheckBox, ChangeLink.Link2_5_CheckBox, ChangeLink.Link2_6_CheckBox, ChangeLink.Link2_7_CheckBox, ChangeLink.Link2_8_CheckBox} _
                                        , {ChangeLink.Link3_1_CheckBox, ChangeLink.Link3_2_CheckBox, ChangeLink.Link3_3_CheckBox, ChangeLink.Link3_4_CheckBox, ChangeLink.Link3_5_CheckBox, ChangeLink.Link3_6_CheckBox, ChangeLink.Link3_7_CheckBox, ChangeLink.Link3_8_CheckBox} _
                                        , {ChangeLink.Link4_1_CheckBox, ChangeLink.Link4_2_CheckBox, ChangeLink.Link4_3_CheckBox, ChangeLink.Link4_4_CheckBox, ChangeLink.Link4_5_CheckBox, ChangeLink.Link4_6_CheckBox, ChangeLink.Link4_7_CheckBox, ChangeLink.Link4_8_CheckBox} _
                                        , {ChangeLink.Link5_1_CheckBox, ChangeLink.Link5_2_CheckBox, ChangeLink.Link5_3_CheckBox, ChangeLink.Link5_4_CheckBox, ChangeLink.Link5_5_CheckBox, ChangeLink.Link5_6_CheckBox, ChangeLink.Link5_7_CheckBox, ChangeLink.Link5_8_CheckBox} _
                                        , {ChangeLink.Link6_1_CheckBox, ChangeLink.Link6_2_CheckBox, ChangeLink.Link6_3_CheckBox, ChangeLink.Link6_4_CheckBox, ChangeLink.Link6_5_CheckBox, ChangeLink.Link6_6_CheckBox, ChangeLink.Link6_7_CheckBox, ChangeLink.Link6_8_CheckBox}}
        Dim LinkNameTextBox As TextBox(,) = {{ChangeLink.Link1_Name_TextBox, ChangeLink.Link2_Name_TextBox, ChangeLink.Link3_Name_TextBox, ChangeLink.Link4_Name_TextBox, ChangeLink.Link5_Name_TextBox, ChangeLink.Link6_Name_TextBox, ChangeLink.Link7_Name_TextBox, ChangeLink.Link8_Name_TextBox} _
                                           , {ChangeLink.Link2_1_Name_TextBox, ChangeLink.Link2_2_Name_TextBox, ChangeLink.Link2_3_Name_TextBox, ChangeLink.Link2_4_Name_TextBox, ChangeLink.Link2_5_Name_TextBox, ChangeLink.Link2_6_Name_TextBox, ChangeLink.Link2_7_Name_TextBox, ChangeLink.Link2_8_Name_TextBox} _
                                           , {ChangeLink.Link3_1_Name_TextBox, ChangeLink.Link3_2_Name_TextBox, ChangeLink.Link3_3_Name_TextBox, ChangeLink.Link3_4_Name_TextBox, ChangeLink.Link3_5_Name_TextBox, ChangeLink.Link3_6_Name_TextBox, ChangeLink.Link3_7_Name_TextBox, ChangeLink.Link3_8_Name_TextBox} _
                                           , {ChangeLink.Link4_1_Name_TextBox, ChangeLink.Link4_2_Name_TextBox, ChangeLink.Link4_3_Name_TextBox, ChangeLink.Link4_4_Name_TextBox, ChangeLink.Link4_5_Name_TextBox, ChangeLink.Link4_6_Name_TextBox, ChangeLink.Link4_7_Name_TextBox, ChangeLink.Link4_8_Name_TextBox} _
                                           , {ChangeLink.Link5_1_Name_TextBox, ChangeLink.Link5_2_Name_TextBox, ChangeLink.Link5_3_Name_TextBox, ChangeLink.Link5_4_Name_TextBox, ChangeLink.Link5_5_Name_TextBox, ChangeLink.Link5_6_Name_TextBox, ChangeLink.Link5_7_Name_TextBox, ChangeLink.Link5_8_Name_TextBox} _
                                           , {ChangeLink.Link6_1_Name_TextBox, ChangeLink.Link6_2_Name_TextBox, ChangeLink.Link6_3_Name_TextBox, ChangeLink.Link6_4_Name_TextBox, ChangeLink.Link6_5_Name_TextBox, ChangeLink.Link6_6_Name_TextBox, ChangeLink.Link6_7_Name_TextBox, ChangeLink.Link6_8_Name_TextBox}}
        Dim LinkButton As Button(,) = {{MagicTool.Link1_1_Button, MagicTool.Link1_2_Button, MagicTool.Link1_3_Button, MagicTool.Link1_4_Button, MagicTool.Link1_5_Button, MagicTool.Link1_6_Button, MagicTool.Link1_7_Button, MagicTool.Link1_8_Button} _
                                     , {MagicTool.Link2_1_Button, MagicTool.Link2_2_Button, MagicTool.Link2_3_Button, MagicTool.Link2_4_Button, MagicTool.Link2_5_Button, MagicTool.Link2_6_Button, MagicTool.Link2_7_Button, MagicTool.Link2_8_Button} _
                                     , {MagicTool.Link3_1_Button, MagicTool.Link3_2_Button, MagicTool.Link3_3_Button, MagicTool.Link3_4_Button, MagicTool.Link3_5_Button, MagicTool.Link3_6_Button, MagicTool.Link3_7_Button, MagicTool.Link3_8_Button} _
                                     , {MagicTool.Link4_1_Button, MagicTool.Link4_2_Button, MagicTool.Link4_3_Button, MagicTool.Link4_4_Button, MagicTool.Link4_5_Button, MagicTool.Link4_6_Button, MagicTool.Link4_7_Button, MagicTool.Link4_8_Button} _
                                     , {MagicTool.Link5_1_Button, MagicTool.Link5_2_Button, MagicTool.Link5_3_Button, MagicTool.Link5_4_Button, MagicTool.Link5_5_Button, MagicTool.Link5_6_Button, MagicTool.Link5_7_Button, MagicTool.Link5_8_Button} _
                                     , {MagicTool.Link6_1_Button, MagicTool.Link6_2_Button, MagicTool.Link6_3_Button, MagicTool.Link6_4_Button, MagicTool.Link6_5_Button, MagicTool.Link6_6_Button, MagicTool.Link6_7_Button, MagicTool.Link6_8_Button}}
        Dim LinkDirTextBox As TextBox(,) = {{ChangeLink.Link1_Dir_TextBox, ChangeLink.Link2_Dir_TextBox, ChangeLink.Link3_Dir_TextBox, ChangeLink.Link4_Dir_TextBox, ChangeLink.Link5_Dir_TextBox, ChangeLink.Link6_Dir_TextBox, ChangeLink.Link7_Dir_TextBox, ChangeLink.Link8_Dir_TextBox} _
                                         , {ChangeLink.Link2_1_Dir_TextBox, ChangeLink.Link2_2_Dir_TextBox, ChangeLink.Link2_3_Dir_TextBox, ChangeLink.Link2_4_Dir_TextBox, ChangeLink.Link2_5_Dir_TextBox, ChangeLink.Link2_6_Dir_TextBox, ChangeLink.Link2_7_Dir_TextBox, ChangeLink.Link2_8_Dir_TextBox} _
                                         , {ChangeLink.Link3_1_Dir_TextBox, ChangeLink.Link3_2_Dir_TextBox, ChangeLink.Link3_3_Dir_TextBox, ChangeLink.Link3_4_Dir_TextBox, ChangeLink.Link3_5_Dir_TextBox, ChangeLink.Link3_6_Dir_TextBox, ChangeLink.Link3_7_Dir_TextBox, ChangeLink.Link3_8_Dir_TextBox} _
                                         , {ChangeLink.Link4_1_Dir_TextBox, ChangeLink.Link4_2_Dir_TextBox, ChangeLink.Link4_3_Dir_TextBox, ChangeLink.Link4_4_Dir_TextBox, ChangeLink.Link4_5_Dir_TextBox, ChangeLink.Link4_6_Dir_TextBox, ChangeLink.Link4_7_Dir_TextBox, ChangeLink.Link4_8_Dir_TextBox} _
                                         , {ChangeLink.Link5_1_Dir_TextBox, ChangeLink.Link5_2_Dir_TextBox, ChangeLink.Link5_3_Dir_TextBox, ChangeLink.Link5_4_Dir_TextBox, ChangeLink.Link5_5_Dir_TextBox, ChangeLink.Link5_6_Dir_TextBox, ChangeLink.Link5_7_Dir_TextBox, ChangeLink.Link5_8_Dir_TextBox} _
                                         , {ChangeLink.Link6_1_Dir_TextBox, ChangeLink.Link6_2_Dir_TextBox, ChangeLink.Link6_3_Dir_TextBox, ChangeLink.Link6_4_Dir_TextBox, ChangeLink.Link6_5_Dir_TextBox, ChangeLink.Link6_6_Dir_TextBox, ChangeLink.Link6_7_Dir_TextBox, ChangeLink.Link6_8_Dir_TextBox}}
        Dim CreateNewFolder_Textbox As TextBox() = {ChangeLink.childFolder_TextBox1, ChangeLink.childFolder_TextBox2, ChangeLink.childFolder_TextBox3 _
                                                   , ChangeLink.childFolder_TextBox4, ChangeLink.childFolder_TextBox5, ChangeLink.childFolder_TextBox6}

        For i = 1 To ini_a
            For j = 1 To ini_b
                ini_LCB = "LCB_" & i & "_" & j
                GetPrivateProfileString_class("LinkCheckbox", ini_LCB, "", sKeyValueCB, nSize, sinifilename)
                LinkCheckBox(i - 1, j - 1).Checked = sKeyValueCB.ToString

                ini_Name = "Name_" & i & "_" & j
                GetPrivateProfileString_class("Name", ini_Name, "", sKeyValueNa, nSize, sinifilename)
                LinkNameTextBox(i - 1, j - 1).Text = sKeyValueNa.ToString

                ini_LinkPath = "LinkPath_" & i & "_" & j
                GetPrivateProfileString_class("LinkPath", ini_LinkPath, "", sKeyValuePath, nSize, sinifilename)
                LinkDirTextBox(i - 1, j - 1).Text = sKeyValuePath.ToString
            Next
            'New folder
            ini_NewFolder = "ChildFolder_" & i
            GetPrivateProfileString_class("NewFolder", ini_NewFolder, "", sKeyValueNewFolder, nSize, sinifilename)
            CreateNewFolder_Textbox(i - 1).Text = sKeyValueNewFolder.ToString
        Next

        'folder path and name 、 common job file
        count_allFolderCBPath = ChangeLink.AllFolderPath_ComboBox.Items.Count '先計算combox有幾個，一開始一定是0
        count_allFolderCBName = ChangeLink.CL_AllFolderName_ComboBox.Items.Count
        count_comJobCBPath = ChangeLink.CommonJobFilePath_Combobox.Items.Count '常用工番count
        count_comJobCBName = ChangeLink.CommonJobFileName_ComboBox.Items.Count
        'count_AllEmployee = ChangeLink.AllEmployee_ComboBox.Items.Count

        count_allFolderCBPath = combobox_state.combobox_when_add(count_allFolderCBPath, ini_FolderPath, "FolderPath_", "AllFolderPath", sKeyFolderPath, ChangeLink.AllFolderPath_ComboBox, MagicTool.FolderPath_ComboBox, nSize, sinifilename)
        count_allFolderCBName = combobox_state.combobox_when_add(count_allFolderCBName, ini_FolderName, "FolderName_", "AllFolderName", sKeyFolderName, ChangeLink.CL_AllFolderName_ComboBox, MagicTool.FolderName_ComboBox, nSize, sinifilename)
        count_comJobCBPath = combobox_state.combobox_when_add(count_comJobCBPath, ini_comJobPath, "CommonFilePath_", "CommonJobFile", sKeyComJobPath, ChangeLink.CommonJobFilePath_Combobox, MagicTool.FileChoUse_ComboBox, nSize, sinifilename)
        'count_AllEmployee = combobox_state.combobox_when_add(count_AllEmployee, ini_allEmp, "Emp_", "AllEmployee", sKeyAllEmp, ChangeLink.AllEmployee_ComboBox, JobMaker_Form.usr_Desinger_ComboBox, nSize, sinifilename)
        'combobox_state.combobox_when_add(count_AllEmployee, ini_allEmp, "Emp_", "AllEmployee", sKeyAllEmp, ChangeLink.AllEmployee_ComboBox, JobMaker_Form.usr_Checker_ComboBox, nSize, sinifilename)

        '當FileChoUse_Combobox的text為某資料夾時，顯示該資料夾底下的檔案
        MagicTool.FileChoUse_ComboBox.Text = MagicTool.FileChoUse_ComboBox.Items(0).ToString

        'magictool執行過一次後就不在執行
        MagicTool.FolderPath_Name_Bool = True

        Dim MA_ChildFolder_Group As TextBox() = {MagicTool.MAchildFolder_TextBox1, MagicTool.MAchildFolder_TextBox2, MagicTool.MAchildFolder_TextBox3, MagicTool.MAchildFolder_TextBox4 _
                                                , MagicTool.MAchildFolder_TextBox5, MagicTool.MAchildFolder_TextBox6}
        Dim CL_ChildFolder_Group As TextBox() = {ChangeLink.childFolder_TextBox1, ChangeLink.childFolder_TextBox2, ChangeLink.childFolder_TextBox3, ChangeLink.childFolder_TextBox4 _
                                                , ChangeLink.childFolder_TextBox5, ChangeLink.childFolder_TextBox6}
        For a = 0 To MagicTool.childForlder_sum - 1  '0 ~ 5
            MA_ChildFolder_Group(a).Text = CL_ChildFolder_Group(a).Text
        Next a


        ''自動檢查INI更新狀態
        'updateINI_program()
        ''自動檢查備分更新
        'update_file_check()
    End Sub








    'All Link 觸發之狀態
    Public Sub LinkCB_setting()
        Dim LinkCheckBox As CheckBox(,) = {{ChangeLink.Link1_CheckBox, ChangeLink.Link2_CheckBox, ChangeLink.Link3_CheckBox, ChangeLink.Link4_CheckBox, ChangeLink.Link5_CheckBox, ChangeLink.Link6_CheckBox, ChangeLink.Link7_CheckBox, ChangeLink.Link8_CheckBox} _
                                        , {ChangeLink.Link2_1_CheckBox, ChangeLink.Link2_2_CheckBox, ChangeLink.Link2_3_CheckBox, ChangeLink.Link2_4_CheckBox, ChangeLink.Link2_5_CheckBox, ChangeLink.Link2_6_CheckBox, ChangeLink.Link2_7_CheckBox, ChangeLink.Link2_8_CheckBox} _
                                        , {ChangeLink.Link3_1_CheckBox, ChangeLink.Link3_2_CheckBox, ChangeLink.Link3_3_CheckBox, ChangeLink.Link3_4_CheckBox, ChangeLink.Link3_5_CheckBox, ChangeLink.Link3_6_CheckBox, ChangeLink.Link3_7_CheckBox, ChangeLink.Link3_8_CheckBox} _
                                        , {ChangeLink.Link4_1_CheckBox, ChangeLink.Link4_2_CheckBox, ChangeLink.Link4_3_CheckBox, ChangeLink.Link4_4_CheckBox, ChangeLink.Link4_5_CheckBox, ChangeLink.Link4_6_CheckBox, ChangeLink.Link4_7_CheckBox, ChangeLink.Link4_8_CheckBox} _
                                        , {ChangeLink.Link5_1_CheckBox, ChangeLink.Link5_2_CheckBox, ChangeLink.Link5_3_CheckBox, ChangeLink.Link5_4_CheckBox, ChangeLink.Link5_5_CheckBox, ChangeLink.Link5_6_CheckBox, ChangeLink.Link5_7_CheckBox, ChangeLink.Link5_8_CheckBox} _
                                        , {ChangeLink.Link6_1_CheckBox, ChangeLink.Link6_2_CheckBox, ChangeLink.Link6_3_CheckBox, ChangeLink.Link6_4_CheckBox, ChangeLink.Link6_5_CheckBox, ChangeLink.Link6_6_CheckBox, ChangeLink.Link6_7_CheckBox, ChangeLink.Link6_8_CheckBox}}
        Dim LinkNameTextBox As TextBox(,) = {{ChangeLink.Link1_Name_TextBox, ChangeLink.Link2_Name_TextBox, ChangeLink.Link3_Name_TextBox, ChangeLink.Link4_Name_TextBox, ChangeLink.Link5_Name_TextBox, ChangeLink.Link6_Name_TextBox, ChangeLink.Link7_Name_TextBox, ChangeLink.Link8_Name_TextBox} _
                                           , {ChangeLink.Link2_1_Name_TextBox, ChangeLink.Link2_2_Name_TextBox, ChangeLink.Link2_3_Name_TextBox, ChangeLink.Link2_4_Name_TextBox, ChangeLink.Link2_5_Name_TextBox, ChangeLink.Link2_6_Name_TextBox, ChangeLink.Link2_7_Name_TextBox, ChangeLink.Link2_8_Name_TextBox} _
                                           , {ChangeLink.Link3_1_Name_TextBox, ChangeLink.Link3_2_Name_TextBox, ChangeLink.Link3_3_Name_TextBox, ChangeLink.Link3_4_Name_TextBox, ChangeLink.Link3_5_Name_TextBox, ChangeLink.Link3_6_Name_TextBox, ChangeLink.Link3_7_Name_TextBox, ChangeLink.Link3_8_Name_TextBox} _
                                           , {ChangeLink.Link4_1_Name_TextBox, ChangeLink.Link4_2_Name_TextBox, ChangeLink.Link4_3_Name_TextBox, ChangeLink.Link4_4_Name_TextBox, ChangeLink.Link4_5_Name_TextBox, ChangeLink.Link4_6_Name_TextBox, ChangeLink.Link4_7_Name_TextBox, ChangeLink.Link4_8_Name_TextBox} _
                                           , {ChangeLink.Link5_1_Name_TextBox, ChangeLink.Link5_2_Name_TextBox, ChangeLink.Link5_3_Name_TextBox, ChangeLink.Link5_4_Name_TextBox, ChangeLink.Link5_5_Name_TextBox, ChangeLink.Link5_6_Name_TextBox, ChangeLink.Link5_7_Name_TextBox, ChangeLink.Link5_8_Name_TextBox} _
                                           , {ChangeLink.Link6_1_Name_TextBox, ChangeLink.Link6_2_Name_TextBox, ChangeLink.Link6_3_Name_TextBox, ChangeLink.Link6_4_Name_TextBox, ChangeLink.Link6_5_Name_TextBox, ChangeLink.Link6_6_Name_TextBox, ChangeLink.Link6_7_Name_TextBox, ChangeLink.Link6_8_Name_TextBox}}
        Dim LinkButton As Button(,) = {{MagicTool.Link1_1_Button, MagicTool.Link1_2_Button, MagicTool.Link1_3_Button, MagicTool.Link1_4_Button, MagicTool.Link1_5_Button, MagicTool.Link1_6_Button, MagicTool.Link1_7_Button, MagicTool.Link1_8_Button} _
                                     , {MagicTool.Link2_1_Button, MagicTool.Link2_2_Button, MagicTool.Link2_3_Button, MagicTool.Link2_4_Button, MagicTool.Link2_5_Button, MagicTool.Link2_6_Button, MagicTool.Link2_7_Button, MagicTool.Link2_8_Button} _
                                     , {MagicTool.Link3_1_Button, MagicTool.Link3_2_Button, MagicTool.Link3_3_Button, MagicTool.Link3_4_Button, MagicTool.Link3_5_Button, MagicTool.Link3_6_Button, MagicTool.Link3_7_Button, MagicTool.Link3_8_Button} _
                                     , {MagicTool.Link4_1_Button, MagicTool.Link4_2_Button, MagicTool.Link4_3_Button, MagicTool.Link4_4_Button, MagicTool.Link4_5_Button, MagicTool.Link4_6_Button, MagicTool.Link4_7_Button, MagicTool.Link4_8_Button} _
                                     , {MagicTool.Link5_1_Button, MagicTool.Link5_2_Button, MagicTool.Link5_3_Button, MagicTool.Link5_4_Button, MagicTool.Link5_5_Button, MagicTool.Link5_6_Button, MagicTool.Link5_7_Button, MagicTool.Link5_8_Button} _
                                     , {MagicTool.Link6_1_Button, MagicTool.Link6_2_Button, MagicTool.Link6_3_Button, MagicTool.Link6_4_Button, MagicTool.Link6_5_Button, MagicTool.Link6_6_Button, MagicTool.Link6_7_Button, MagicTool.Link6_8_Button}}
        Dim LinkDirTextBox As TextBox(,) = {{ChangeLink.Link1_Dir_TextBox, ChangeLink.Link2_Dir_TextBox, ChangeLink.Link3_Dir_TextBox, ChangeLink.Link4_Dir_TextBox, ChangeLink.Link5_Dir_TextBox, ChangeLink.Link6_Dir_TextBox, ChangeLink.Link7_Dir_TextBox, ChangeLink.Link8_Dir_TextBox} _
                                         , {ChangeLink.Link2_1_Dir_TextBox, ChangeLink.Link2_2_Dir_TextBox, ChangeLink.Link2_3_Dir_TextBox, ChangeLink.Link2_4_Dir_TextBox, ChangeLink.Link2_5_Dir_TextBox, ChangeLink.Link2_6_Dir_TextBox, ChangeLink.Link2_7_Dir_TextBox, ChangeLink.Link2_8_Dir_TextBox} _
                                         , {ChangeLink.Link3_1_Dir_TextBox, ChangeLink.Link3_2_Dir_TextBox, ChangeLink.Link3_3_Dir_TextBox, ChangeLink.Link3_4_Dir_TextBox, ChangeLink.Link3_5_Dir_TextBox, ChangeLink.Link3_6_Dir_TextBox, ChangeLink.Link3_7_Dir_TextBox, ChangeLink.Link3_8_Dir_TextBox} _
                                         , {ChangeLink.Link4_1_Dir_TextBox, ChangeLink.Link4_2_Dir_TextBox, ChangeLink.Link4_3_Dir_TextBox, ChangeLink.Link4_4_Dir_TextBox, ChangeLink.Link4_5_Dir_TextBox, ChangeLink.Link4_6_Dir_TextBox, ChangeLink.Link4_7_Dir_TextBox, ChangeLink.Link4_8_Dir_TextBox} _
                                         , {ChangeLink.Link5_1_Dir_TextBox, ChangeLink.Link5_2_Dir_TextBox, ChangeLink.Link5_3_Dir_TextBox, ChangeLink.Link5_4_Dir_TextBox, ChangeLink.Link5_5_Dir_TextBox, ChangeLink.Link5_6_Dir_TextBox, ChangeLink.Link5_7_Dir_TextBox, ChangeLink.Link5_8_Dir_TextBox} _
                                         , {ChangeLink.Link6_1_Dir_TextBox, ChangeLink.Link6_2_Dir_TextBox, ChangeLink.Link6_3_Dir_TextBox, ChangeLink.Link6_4_Dir_TextBox, ChangeLink.Link6_5_Dir_TextBox, ChangeLink.Link6_6_Dir_TextBox, ChangeLink.Link6_7_Dir_TextBox, ChangeLink.Link6_8_Dir_TextBox}}

        For i = 1 To ini_a
            For j = 1 To ini_b
                ini_LCB = "LCB_" & i & "_" & j
                LinkCB_setting_ifelse(LinkCheckBox(i - 1, j - 1), sKeyValueCB, "LinkCheckbox", ini_LCB, LinkButton(i - 1, j - 1),
                                      LinkNameTextBox(i - 1, j - 1), LinkDirTextBox(i - 1, j - 1))
            Next
        Next


    End Sub
    Private Sub LinkCB_setting_ifelse(Link_ChkBx As CheckBox, sKeyVa As StringBuilder, LCB As String, LCB_In As String,
                                      Ma_BTN As Button, LinkNa_TB As TextBox, LinkDir_TB As TextBox) 'Link勾勾按鈕的判斷是
        If Link_ChkBx.Checked = True Then
            WriteInini_Fun(sKeyVa, True, LCB, LCB_In)
            IfCB_Click(Link_ChkBx, Ma_BTN, LinkNa_TB, LinkDir_TB, True)
        Else
            WriteInini_Fun(sKeyVa, False, LCB, LCB_In)
            IfCB_Click(Link_ChkBx, Ma_BTN, LinkNa_TB, LinkDir_TB, True)
        End If
    End Sub

    Public Sub WriteInini_Fun(sKeyVa As StringBuilder, bool As String, TitleN As String, SecN As String) '寫入ini
        If sKeyVa.ToString IsNot "" Then
            sKeyVa.Clear()
        End If
        sKeyVa = sKeyVa.Append(bool)
        WritePrivateProfileString_class(TitleN, SecN, sKeyVa, sinifilename)
    End Sub

    Private Sub IfCB_Click(Link_CB As CheckBox, Link_B As Button, LinkNa_TB As TextBox, LinkDir_TB As TextBox, WriteAble As Boolean)
        'All Link 是否被勾選，顯示當前能使用或不能 LinkCB_setting_ifelse 使用

        If Link_CB.Checked = True Then
            LinkNa_TB.Enabled = True
            LinkDir_TB.Enabled = True
            If WriteAble = True Then
                Link_B.Enabled = True
                Link_B.Text = LinkNa_TB.Text
            End If
        Else
            LinkNa_TB.Enabled = False
            LinkDir_TB.Enabled = False
            If WriteAble = True Then
                Link_B.Enabled = False
                Link_B.Text = Link_CB.Text
                'Link_B.Text = "付費激活"
            End If
        End If

    End Sub









    '設定介面處於哪個位置
    Public Sub ScreenPosChange(Form_name As Form, Scr As String, Pos As String)
        'Dim Scr, Pos As String

        If Scr = "主螢幕" Then

            If Pos = "左上" Then
                ScreenPosChange_setting(Form_name, 0, 0)
            ElseIf Pos = "右上" Then
                ScreenPosChange_setting(Form_name, WorkingArea.Width - Form_name.Width, 0)
            ElseIf Pos = "左下" Then
                ScreenPosChange_setting(Form_name, 0, WorkingArea.Height - Form_name.Height)
            ElseIf Pos = "右下" Then
                ScreenPosChange_setting(Form_name, WorkingArea.Width - Form_name.Width, WorkingArea.Height - Form_name.Height)
            End If
        ElseIf Scr = "副螢幕" Then

            If Pos = "左上" Then
                ScreenPosChange_setting(Form_name, -1920, 0)
            ElseIf Pos = "右上" Then
                ScreenPosChange_setting(Form_name, -Form_name.Width, 0)
            ElseIf Pos = "左下" Then
                ScreenPosChange_setting(Form_name, -1920, WorkingArea.Height - Form_name.Height)
            ElseIf Pos = "右下" Then
                ScreenPosChange_setting(Form_name, -Form_name.Width, WorkingArea.Height - Form_name.Height)
            End If
            'ElseIf Scr = "自訂" And Pos = "自訂" Then
            'MsgBox(Form_name.ToString & " " & Form_name.DesktopLocation.X & " " & Form_name.DesktopLocation.Y)
            'ScreenPosChange_setting(Form_name, Me.DesktopLocation.X, Me.DesktopLocation.Y)
        End If
    End Sub

    Private Sub ScreenPosChange_setting(FormName As Form, Form_Position_X As Integer, Form_Position_Y As Integer) '傳回設定介面位置之值
        FormName.Location = New System.Drawing.Point(Form_Position_X, Form_Position_Y)


        'countA = countA + 1
        'MsgBox($"Form_Position_X : {Form_Position_X},Form_Position_Y :{Form_Position_Y},countA : {countA},form : {FormName}")


    End Sub









    '介面是否設定在最上層
    Public Sub Topmost_setting(form As Form, save As Boolean)

        If ChangeLink.Topmost_CheckBox.Checked = True Then

            If save = True Then
                WriteInini_Fun(sKeyValue, "True", "CheckBox_State", "TopmostSet")
            End If
            form.Top = True
        Else
            If save = True Then
                WriteInini_Fun(sKeyValue, "False", "CheckBox_State", "TopmostSet")
            End If
            form.Top = False
        End If
    End Sub







    '當combobx加一時防止一次加過多情形
    Overloads Function combobox_when_add(count_CB As Integer, ini_name As String, ini_childname As String _
                                        , ini_fatername As String, skey As System.Text.StringBuilder _
                                        , ChaLink_myCB As ComboBox, Form_myCB As ComboBox, nSize As UInt32, sinifilename As String)
        Dim a As Integer = 0
        Form_myCB.Items.Clear()
        While a <= count_CB  'a從0開始與count_allFolder比較 相等時執行
            ini_name = ini_childname & a + 1 'folderPath_1開始
            GetPrivateProfileString(ini_fatername, ini_name, "", skey, nSize, sinifilename) '取得數值
            If skey.ToString <> "" Then
                ChaLink_myCB.Items.Add(skey.ToString) '如果取得值不為空則加入combox

                '加過後就給true 就不會造成一值add便很多的狀況
                'If magicTool_ifadd = False Then
                Form_myCB.Items.Add(skey.ToString)
                'End If
            Else
                Exit While '否則跳出迴圈
            End If
            a += 1
            count_CB = ChaLink_myCB.Items.Count '重新數combox的數量傳回去，反覆進行
        End While
        Return count_CB
    End Function
    Overloads Function combobox_when_add(count_CB As Integer, ini_name As String, ini_childname As String _
                                         , ini_fatername As String, skey As System.Text.StringBuilder _
                                         , myCB As ComboBox, magicTool_myCkList As CheckedListBox, nSize As UInt32, sinifilename As String) '當combobx加一時防止一次加過多情形
        Dim a As Integer = 0
        magicTool_myCkList.Items.Clear()
        While a <= count_CB  'a從0開始與count_allFolder比較 相等時執行
            ini_name = ini_childname & a + 1 'folderPath_1開始
            GetPrivateProfileString(ini_fatername, ini_name, "", skey, nSize, sinifilename) '取得數值
            If skey.ToString <> "" Then
                myCB.Items.Add(skey.ToString) '如果取得值不為空則加入combox

                '加過後就給true 就不會造成一值add便很多的狀況
                'If magicTool_ifadd = False Then
                magicTool_myCkList.Items.Add(skey.ToString)
                'End If
            Else
                Exit While '否則跳出迴圈
            End If
            a += 1
            count_CB = myCB.Items.Count '重新數combox的數量傳回去，反覆進行
        End While
        Return count_CB
    End Function
    Sub combobox_when_save(count_CB As Integer, ini_name As String, ini_childname As String, ini_fathername As String _
                           , skey As System.Text.StringBuilder, cb As ComboBox _
                           , nSize As UInt32, sinifilename As String)
        Dim a As Integer
        For a = 0 To count_CB
            ini_name = ini_childname & a + 1

            If skey.ToString IsNot "" Then
                skey.Clear()
            End If

            Try   '如果all folder path 有刪除動作時，在跑combox.item時會少1，造成error，使用try catch讓刪除的值達成清空的步驟
                skey = skey.Append(cb.Items(a).ToString())
                WritePrivateProfileString(ini_fathername, ini_name, skey, sinifilename)
            Catch ex As Exception
                skey.Clear()
                WritePrivateProfileString(ini_fathername, ini_name, skey, sinifilename)
            End Try
        Next a

    End Sub






    Sub Add_button(tb As TextBox, cb As ComboBox) '新增鈕用在基本設定內
        If tb.Text <> "" Then
            If tb.Text = cb.Text Then
                Exit Sub
            End If
            cb.Items.Add(tb.Text)
        End If
    End Sub

    Sub Sub_button(tb As TextBox, cb As ComboBox) '刪除鈕用在基本設定內
        If tb.Text = cb.Text Then
            cb.Items.Remove(tb.Text)
        End If
    End Sub
End Class
