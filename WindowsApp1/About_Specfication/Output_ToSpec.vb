Imports Microsoft.Office.Interop

Public Class Output_ToSpec
    ''' <summary>
    ''' 錯誤回報訊息，回傳錯誤的名稱管理員
    ''' </summary>
    Public returnError_specName As String

    ''' <summary>
    ''' 錯誤回報訊息，回傳分頁有沒有打勾
    ''' </summary>
    Public returnError_isPageRestart As Boolean

    Const ctrlTypeName_Panel As String = "Panel"
    Const ctrlTypeName_ComboBox As String = "ComboBox"
    Const ctrlTypeName_Label As String = "Label"
    Const ctrlTypeName_TextBox As String = "TextBox"
    Const ctrlTypeName_CheckBox As String = "CheckBox"

    Dim get_NameManager As Spec_NameManager = New Spec_NameManager()
    'Dim getMathOnExcel As getMath_onExcel = New getMath_onExcel()



    Public Sub Spec_FinalCheck(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application)
        Dim spec_item As Spec_Item = New Spec_Item()
        spec_item.ini_specTW_AllControler()

        Dim finalCheck_item_col, finalCheck_state_col, finalCheck_Spec_col As Integer
        Dim finalCheck_item_row, finalCheck_state_row, finalCheck_Spec_row As Integer


        Dim mCtrlNameForError, mPanelNameForError As String
        mCtrlNameForError = ""
        mPanelNameForError = ""
        Try
            '全部仕樣確認表欄與列
            '項目
            finalCheck_item_col =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_Item)
            finalCheck_item_row =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_Item)
            '有無
            finalCheck_state_col =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_State)
            finalCheck_state_row =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_State)
            '仕樣
            finalCheck_Spec_col =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_Spec)
            finalCheck_Spec_row =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_Spec)


            Dim item_countRow As Integer 'item輸出時的列數
            Dim item_number As Integer '目前為第幾個item
            For Each mPanel As Control In spec_item.specTW_panel
                For Each mCtrlTitle As Control In mPanel.Controls
                    mCtrlNameForError = mCtrlTitle.Name
                    mPanelNameForError = mPanel.Name
                    If mPanel.Enabled And TypeOf (mCtrlTitle) Is ComboBox Then
                        '判斷是否Panel名稱與ComboBox相同，且為打圈
                        If spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mPanel, ctrlTypeName_Panel, "") =
                           spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mCtrlTitle, ctrlTypeName_ComboBox, "") And
                           mCtrlTitle.Text = get_NameManager.TB_O Then
                            item_countRow += 1
                            item_number += 1
                            getMathOnExcel.
                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                     get_NameManager.read_DbmsData(get_NameManager.FinalCheck_Item,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     item_countRow,
                                                                     item_number)
                            getMathOnExcel.
                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                     get_NameManager.read_DbmsData(get_NameManager.FinalCheck_State,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     item_countRow,
                                                                     "O")

                            Dim afterReplaceTitle_Label As String 'Label取代Panel後的名稱，例如:A_Panel > A_Label
                            afterReplaceTitle_Label =
                                spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mPanel, ctrlTypeName_Panel, ctrlTypeName_Label)
                            getMathOnExcel.
                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                    get_NameManager.read_DbmsData(get_NameManager.FinalCheck_Spec,
                                                                                                get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                get_NameManager.SQLite_connectionPath_Tool,
                                                                                                get_NameManager.SQLite_ToolDBMS_Name),
                                                                    item_countRow,
                                                                    spec_item.getRelace_NameText_onPanel(afterReplaceTitle_Label, mPanel))


                            '依照目前的mPanel取得除了標題(例:正背門)之外的控制項Control
                            For Each mCtrlContent As Control In mPanel.Controls
                                '取得除了標題(例:正背門)之外的控制項
                                If spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mPanel, ctrlTypeName_Panel, "") <>
                                   spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mCtrlContent, ctrlTypeName_ComboBox, "") Then
                                    '控制項為ComboBox並<有文字>或為<O>或<With>才寫入
                                    If (TypeOf (mCtrlContent) Is ComboBox And mCtrlContent.Text <> "" And mCtrlContent.Enabled = True) And
                                       (mCtrlContent.Text <> get_NameManager.TB_X And mCtrlContent.Text <> get_NameManager.TB_WITHOUT) Then
                                        item_countRow += 1

                                        Dim afterReplaceContent_Label As String
                                        afterReplaceContent_Label = spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mCtrlContent, ctrlTypeName_ComboBox, ctrlTypeName_Label)
                                        Dim getLabelText As String
                                        getLabelText = spec_item.getRelace_NameText_onPanel(afterReplaceContent_Label, mPanel)

                                        If (mCtrlContent.Text = get_NameManager.TB_WITH Or mCtrlContent.Text = get_NameManager.TB_O) Then
                                            '
                                        Else
                                            '控制項Control內容非With或O時將Label與TextBox內容一同輸出
                                            getLabelText += $"{mCtrlContent.Text} "
                                        End If


                                        getMathOnExcel.
                                            setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                 get_NameManager.read_DbmsData(get_NameManager.FinalCheck_State,
                                                                                                               get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                               get_NameManager.SQLite_connectionPath_Tool,
                                                                                                               get_NameManager.SQLite_ToolDBMS_Name),
                                                                                 item_countRow,
                                                                                 "O")

                                        getMathOnExcel.
                                            setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                 get_NameManager.read_DbmsData(get_NameManager.FinalCheck_Spec,
                                                                                                               get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                               get_NameManager.SQLite_connectionPath_Tool,
                                                                                                               get_NameManager.SQLite_ToolDBMS_Name),
                                                                                 item_countRow,
                                                                                 getLabelText)
                                    ElseIf TypeOf (mCtrlContent) Is TextBox And mCtrlContent.Text <> "" And mCtrlContent.Enabled = True Then
                                        '其他TextBox
                                        item_countRow += 1

                                        Dim nameAfterReplace_ChkBox, nameAfterReplace_Label As String
                                        nameAfterReplace_ChkBox =
                                            spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mCtrlContent, ctrlTypeName_TextBox, ctrlTypeName_CheckBox)
                                        nameAfterReplace_Label =
                                            spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mCtrlContent, ctrlTypeName_TextBox, ctrlTypeName_Label)

                                        Dim is_ChkBox_checked As Boolean '如果控制項為CheckBox時的狀態，僅打勾的才輸出
                                        is_ChkBox_checked =
                                            spec_item.getRelace_ChkBoxState_onPanel(nameAfterReplace_ChkBox, mPanel)

                                        Dim getLabelText, getChkBoxText As String
                                        getLabelText =
                                            spec_item.getRelace_NameText_onPanel(nameAfterReplace_Label, mPanel)


                                        If is_ChkBox_checked = True Then
                                            getMathOnExcel.
                                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                     get_NameManager.read_DbmsData(get_NameManager.FinalCheck_State,
                                                                                                                   get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                                     item_countRow,
                                                                                     "O")
                                            getChkBoxText =
                                                spec_item.getRelace_NameText_onPanel(nameAfterReplace_ChkBox, mPanel)

                                            getMathOnExcel.
                                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                     get_NameManager.read_DbmsData(get_NameManager.FinalCheck_Spec,
                                                                                                                   get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                                     item_countRow,
                                                                                     $"{getChkBoxText} : {mCtrlContent.Text}")
                                        End If

                                        If getLabelText <> "" And mCtrlContent.Text <> "" Then
                                            getMathOnExcel.
                                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                     get_NameManager.read_DbmsData(get_NameManager.FinalCheck_State,
                                                                                                                   get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                                     item_countRow,
                                                                                     "O")

                                            getMathOnExcel.
                                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                     get_NameManager.read_DbmsData(get_NameManager.FinalCheck_Spec,
                                                                                                                   get_NameManager.SQLite_tableName_NameManager_FinalCheck,
                                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                                     item_countRow,
                                                                                     $"{getLabelText} : {mCtrlContent.Text}")
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox($"最後檢查輸出表<{mCtrlNameForError}>錯誤{vbCrLf} {ex.ToString}")
        End Try

    End Sub

    Public Sub Spec_Spec_Std(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application)

        If JobMaker_Form.Use_Basic_CheckBox.Checked Then
            Dim usr_JobNo_New, usr_JobName, usr_Designer, usr_Checker, usr_Approver As String
            Dim usr_drawDate As String

            usr_JobNo_New = JobMaker_Form.Basic_JobNoNew_TextBox.Name
            usr_JobName = JobMaker_Form.Basic_JobName_TextBox.Name

            usr_Designer = ""
            If JobMaker_Form.Basic_DesingerChinese_ComboBox.Text <> "" Then
                usr_Designer =
                        JobMaker_Form.Basic_DesingerChinese_ComboBox.Name '設計者中文
            ElseIf JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text <> "" Then
                usr_Designer =
                        JobMaker_Form.Basic_DesingerEnglish_ComboBox.Name '設計者英文
            End If

            usr_Checker = ""
            If JobMaker_Form.Basic_CheckerChinese_ComboBox.Text <> "" Then
                usr_Checker =
                        JobMaker_Form.Basic_CheckerChinese_ComboBox.Name '檢查者中文
            ElseIf JobMaker_Form.Basic_DesingerChinese_ComboBox.Text <> "" Then
                usr_Checker =
                        JobMaker_Form.Basic_CheckerEnglish_ComboBox.Name '檢查者英文
            End If

            usr_Approver = ""
            If JobMaker_Form.Basic_ApproverChinese_ComboBox.Text <> "" Then
                usr_Approver =
                        JobMaker_Form.Basic_ApproverChinese_ComboBox.Name '承認者中文
            ElseIf JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text <> "" Then
                usr_Approver =
                        JobMaker_Form.Basic_ApproverEnglish_ComboBox.Name '承認者英文
            End If

            usr_drawDate = JobMaker_Form.Basic_DrawDate_DateTimePicker.Name

            Dim usrInput_arr As String()
            usrInput_arr = {usr_JobNo_New, usr_JobName, usr_Designer,
                            usr_Checker, usr_Approver, usr_drawDate}

            For Each i_str In usrInput_arr
                Try
                    Select Case i_str
                        '工番號
                        Case usr_JobNo_New
                            excelWriteIn(JobMaker_Form.Basic_JobNoNew_TextBox.Text,
                                         get_NameManager.read_DbmsData(get_NameManager.JOBNO,
                                                                       get_NameManager.SQLite_tableName_NameManager_TW,
                                                                       get_NameManager.SQLite_connectionPath_Tool,
                                                                       get_NameManager.SQLite_ToolDBMS_Name),
                                         msExcel_workbook)
                        '工番名
                        Case usr_JobName
                            excelWriteIn(JobMaker_Form.Basic_JobName_TextBox.Text,
                                         get_NameManager.read_DbmsData(get_NameManager.JOBNAME,
                                                                       get_NameManager.SQLite_tableName_NameManager_TW,
                                                                       get_NameManager.SQLite_connectionPath_Tool,
                                                                       get_NameManager.SQLite_ToolDBMS_Name),
                                         msExcel_workbook)

                        '設計者
                        Case usr_Designer
                            If JobMaker_Form.Basic_DesingerChinese_ComboBox.Text <> "" Then
                                excelWriteIn(JobMaker_Form.Basic_DesingerChinese_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.DESIGENED,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            ElseIf JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text <> "" Then
                                excelWriteIn(JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.DESIGENED,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            End If
                        '審查者
                        Case usr_Checker
                            If JobMaker_Form.Basic_CheckerChinese_ComboBox.Text <> "" Then
                                excelWriteIn(JobMaker_Form.Basic_CheckerChinese_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.CHECKED,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            ElseIf JobMaker_Form.Basic_CheckerEnglish_ComboBox.Text <> "" Then
                                excelWriteIn(JobMaker_Form.Basic_CheckerEnglish_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.CHECKED,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            End If
                        '承認者
                        Case usr_Approver
                            If JobMaker_Form.Basic_ApproverChinese_ComboBox.Text <> "" Then
                                excelWriteIn(JobMaker_Form.Basic_ApproverChinese_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.APPROVED,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            ElseIf JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text <> "" Then
                                excelWriteIn(JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.APPROVED,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            End If
                        '作圖日
                        Case usr_drawDate
                            Dim usr_drawDate_val As String
                            usr_drawDate_val =
                                    $"{monthTransfer_sub()}.{JobMaker_Form.Basic_DrawDate_DateTimePicker.Value.Day}.{JobMaker_Form.Basic_DrawDate_DateTimePicker.Value.Year}" 'Date出圖時間
                            excelWriteIn(usr_drawDate_val,
                                         get_NameManager.read_DbmsData(get_NameManager.DRAW_DATE,
                                                                       get_NameManager.SQLite_tableName_NameManager_TW,
                                                                       get_NameManager.SQLite_connectionPath_Tool,
                                                                       get_NameManager.SQLite_ToolDBMS_Name),
                                         msExcel_workbook)
                    End Select
                Catch ex As Exception
                    JobMaker_Form.ResultOutput_TextBox.Text +=
                        ($"<{JobMaker_Form.JMFileCho_Spec_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_str}>{vbCrLf}")
                    JobMaker_Form.ResultOutput_TextBox.Text +=
                        $"----------------------------------"
                End Try
            Next
        Else
            JobMaker_Form.ResultFailOutput_TextBox.Text += ($"<提醒> 基本 分頁未輸出，原因:分頁未打勾{vbCrLf}")
            'JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.Basic_TabPage
            Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.Basic_TabPage.Text}」未使用是否重來?"), vbYesNo)
            If basic_result = DialogResult.Yes And msExcel_workbook IsNot Nothing Then
                returnError_isPageRestart = True
                'msExcel_workbook.Close()
                'msExcel_app.Quit()
            End If
        End If
    End Sub


    ''' <summary>
    ''' Job Maker >> 基本 (快速摺疊Code:CRTL+M+M)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="msExcel_app"></param>
    Public Sub Spec_Std(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application) '基本分頁內容
        '使用者輸入的值
        Dim usr_JobNo_New, usr_JobNo_Old, usr_JobNo_MOD, usr_JobName, usr_Designer, usr_Checker, usr_Approver _
            , usr_Local, usr_DrawDate As String
        Dim usrInput_arr As String()

        If JobMaker_Form.Use_Basic_CheckBox.Checked Then
            usr_Local =
                JobMaker_Form.Basic_Local_ComboBox.Name 'Local地區
            usr_JobNo_New =
                JobMaker_Form.Basic_JobNoNew_TextBox.Name 'JobNo新工番番號
            usr_JobNo_Old =
                JobMaker_Form.Basic_JobNoOld_TextBox.Name 'JobNo舊工番番號
            usr_JobNo_MOD =
                JobMaker_Form.Basic_JobNoMOD_TextBox.Text 'JobNo Mod工番番號
            usr_JobName =
                JobMaker_Form.Basic_JobName_TextBox.Text 'JobName工番名字

            If JobMaker_Form.Basic_DesingerChinese_ComboBox.Text <> "" Then
                usr_Designer =
                    JobMaker_Form.Basic_DesingerChinese_ComboBox.Text '設計者中文
            ElseIf JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text <> "" Then
                usr_Designer =
                    JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text '設計者英文
            End If
            If JobMaker_Form.Basic_CheckerChinese_ComboBox.Text <> "" Then
                usr_Checker =
                    JobMaker_Form.Basic_CheckerChinese_ComboBox.Text '檢查者中文
            ElseIf JobMaker_Form.Basic_DesingerChinese_ComboBox.Text <> "" Then
                usr_Checker =
                    JobMaker_Form.Basic_CheckerEnglish_ComboBox.Text '檢查者英文
            End If
            If JobMaker_Form.Basic_ApproverChinese_ComboBox.Text <> "" Then
                usr_Approver =
                    JobMaker_Form.Basic_ApproverChinese_ComboBox.Text '承認者中文
            ElseIf JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text <> "" Then
                usr_Approver =
                    JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text '承認者英文
            End If

            usr_DrawDate =
                 monthTransfer_sub() & "." & JobMaker_Form.Basic_DrawDate_DateTimePicker.Value.Day & "." _
                 & JobMaker_Form.Basic_DrawDate_DateTimePicker.Value.Year 'Date出圖時間
            'usr_Escape_FL = JobMaker_Form.Spec_EscapeFL_TextBox.Text '避難階
            'usr_Escape_FL_only = JobMaker_Form.Spec_EscapeFL_Only_TextBox.Text '避難階只有n樓
            'usr_Main_FL = JobMaker_Form.Main_FL_TextBox.Text '基準階
            'usr_Parking_FL = JobMaker_Form.Spec_Parking_FL_TextBox.Text '停車階

            '儲存每一個使用者輸入的值
            usrInput_arr = {usr_JobNo_New, usr_JobNo_Old, usr_JobNo_MOD, usr_JobName,
                            usr_Designer, usr_Checker, usr_Approver, usr_Local,
                            usr_DrawDate}

            '輸入相對應的基本值
            'If JobMaker_Form.Use_Basic_CheckBox.CheckState Then
            For Each i_str In usrInput_arr
                If i_str <> "" Then
                    Try
                        Select Case i_str
                            Case usr_JobNo_New
                                excelWriteIn(usr_JobNo_New,
                                             get_NameManager.read_DbmsData(get_NameManager.JOBNO,
                                                                           get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            Case usr_JobNo_Old
                                excelWriteIn(usr_JobNo_Old,
                                             get_NameManager.read_DbmsData(get_NameManager.JOBNO_OLD,
                                                                           get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            Case usr_JobNo_MOD
                                excelWriteIn(usr_JobNo_MOD,
                                             get_NameManager.read_DbmsData(get_NameManager.JOBNO_MOD,
                                                                           get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            Case usr_JobName
                                excelWriteIn(usr_JobName,
                                             get_NameManager.read_DbmsData(get_NameManager.JOBNAME,
                                                                           get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            Case usr_Designer
                                excelWriteIn(usr_Designer,
                                             get_NameManager.read_DbmsData(get_NameManager.DESIGENED,
                                                                           get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            Case usr_Approver
                                excelWriteIn(usr_Approver,
                                             get_NameManager.read_DbmsData(get_NameManager.APPROVED,
                                                                           get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            Case usr_Checker
                                excelWriteIn(usr_Checker,
                                             get_NameManager.read_DbmsData(get_NameManager.CHECKED,
                                                                           get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            Case usr_DrawDate
                                excelWriteIn(usr_DrawDate,
                                             get_NameManager.read_DbmsData(get_NameManager.DRAW_DATE,
                                                                           get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                        End Select
                    Catch ex As Exception
                        JobMaker_Form.ResultOutput_TextBox.Text +=
                            ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_str}>{vbCrLf}")
                    End Try
                End If
            Next
        Else
            JobMaker_Form.ResultFailOutput_TextBox.Text += ("<提醒> 基本 分頁未輸出，原因:分頁未打勾")
            JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.Basic_TabPage
            Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.Basic_TabPage.Text}」未使用是否重來?"), vbYesNo)
            If basic_result = DialogResult.Yes And msExcel_workbook IsNot Nothing Then
                returnError_isPageRestart = True
                'msExcel_workbook.Close()
                'msExcel_app.Quit()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Job Maker >> Check List / 程式變更 (快速摺疊Code:CRTL+M+M)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="msExcel_app"></param>
    Public Sub Spec_CheckList(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application)
        If JobMaker_Form.Use_ChkList_CheckBox.Checked Or JobMaker_Form.Use_prk_CheckBox.Checked Then
            JobMaker_Form.ResultOutput_TextBox.Text += "ˇˇˇˇˇˇˇˇˇˇˇ Check List ˇˇˇˇˇˇˇˇˇˇˇ" & vbCrLf
            ' Check List > CheckList_P1 > 名稱管理員 ---------------------------------------------------------------------------
            Dim usr_nameManager_chkList_Q1, usr_chkList_Q1_YCont, usr_chkList_Q1_YResult,
                usr_nameManager_chkList_Q2, usr_chkList_Q2_YCont, usr_chkList_Q2_YResult,
                usr_nameManager_chkList_Q3, usr_chkList_Q3_YMan, usr_chkList_Q3_YCont, usr_chkList_Q3_YResult,
                usr_chkList_Q4_MMIC, usr_chkList_Q4_MMICBase, usr_chkList_Q4_SV, usr_chkList_Q4_SVBase,
                usr_nameManager_chkList_Q5_chkBoxState, usr_chkList_Q5_StdCont, usr_chkList_Q5_nStdCont,
                usr_nameManager_chkList_Q6, usr_nameManager_chkList_Q6_yes, usr_chkList_Q6_YCont,
                usr_nameManager_chkList_Q7, usr_chkList_Q7_YCont,
                usr_nameManager_chkList_Q8, usr_nameManager_chkList_Q9,
                usr_nameManager_chkList_PA, usr_nameManager_chkList_OS, usr_nameManager_chkList_CFM, usr_nameManager_chkList_ELE,
                usr_chkList_PA_year, usr_chkList_PA_month, usr_chkList_PA_date,
                usr_chkList_OS_year, usr_chkList_OS_month, usr_chkList_OS_date,
                usr_chkList_CFM_year, usr_chkList_CFM_month, usr_chkList_CFM_date,
                usr_chkList_ELE_year, usr_chkList_ELE_month, usr_chkList_ELE_date,
                usr_prgm_Reason,
                usr_prgm_2_testCont, usr_prgm_2_CopCont, usr_prgm_2_TowerCont, usr_prgm_2_OtherCont,
                usr_prgm_3_otherCont,
                usr_prgm_4_testCont,
                usr_nameManager_prgm_2_Test, usr_nameManager_prgm_2_COP, usr_nameManager_prgm_2_Tower, usr_nameManager_prgm_2_Other,
                usr_nameManager_prgm_3_Debug, usr_nameManager_prgm_3_Test, usr_nameManager_prgm_3_CFM, usr_nameManager_prgm_3_EXE, usr_nameManager_prgm_3_Other,
                usr_nameManager_prgm_Auto, usr_nameManager_prgm_Output, usr_nameManager_prgm_INI,
                usr_nameManager_prgm_Case, usr_nameManager_prgm_IF, usr_nameManager_prgm_Loop, usr_nameManager_prgm_Range,
                usr_nameManager_prgm_Casting, usr_nameManager_prgm_0, usr_nameManager_prgm_Count,
                usr_nameManager_prgm_Address, usr_nameManager_prgm_Custom As String
            '--------------------------------------------------------------------------- Check List > CheckList_P1 > 名稱管理員 


            ' Check List > CheckList_P1/P2 分頁名稱 ---------------------------------------------------------------------------
            Dim chkListP1_ShtName, chkListP2_ShtName As String   '分頁頁名
            chkListP1_ShtName = get_NameManager.read_DbmsData(get_NameManager.ChkList_P1_PageName,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name)
            chkListP2_ShtName = get_NameManager.read_DbmsData(get_NameManager.ChkList_P2_PageName,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name)
            '--------------------------------------------------------------------------- Check List > CheckList_P1/P2 分頁名稱 

            ' Check List > CheckList_P1 > 名稱管理員 ---------------------------------------------------------------------------
            Dim chkList_PA_ChkBox, chkList_OS_ChkBox,
                chkList_CFM_ChkBox, chkList_ELE_ChkBox As String
            chkList_PA_ChkBox = get_NameManager.read_DbmsData(get_NameManager.ChkList_PA_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name)
            chkList_OS_ChkBox = get_NameManager.read_DbmsData(get_NameManager.ChkList_OS_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name)
            chkList_CFM_ChkBox = get_NameManager.read_DbmsData(get_NameManager.ChkList_CFM_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name)
            chkList_ELE_ChkBox = get_NameManager.read_DbmsData(get_NameManager.ChkList_ELE_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name)
            '--------------------------------------------------------------------------- Check List > CheckList_P1 > 名稱管理員

            ' CheckList中日期的CheckBox ---------------------------------------------------------------------------------------
            usr_nameManager_chkList_PA = chkBoxStateRead(JobMaker_Form.ChkList_PaSheet_CheckBox, chkList_PA_ChkBox)
            usr_nameManager_chkList_OS = chkBoxStateRead(JobMaker_Form.ChkList_OS_CheckBox, chkList_OS_ChkBox)
            usr_nameManager_chkList_CFM = chkBoxStateRead(JobMaker_Form.ChkList_Confirm_CheckBox, chkList_CFM_ChkBox)
            usr_nameManager_chkList_ELE = chkBoxStateRead(JobMaker_Form.ChkList_Elec_CheckBox, chkList_ELE_ChkBox)
            '--------------------------------------------------------------------------------------- CheckList中日期的CheckBox 


            usr_nameManager_chkList_Q1 =
                chkBoxStateRead(JobMaker_Form.ChkList_1_no_RadioButton, JobMaker_Form.ChkList_1_yes_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q1No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q1Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))
            usr_nameManager_chkList_Q2 =
                chkBoxStateRead(JobMaker_Form.ChkList_2_no_RadioButton, JobMaker_Form.ChkList_2_yes_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q2No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q2Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))
            usr_nameManager_chkList_Q3 =
                chkBoxStateRead(JobMaker_Form.ChkList_3_no_RadioButton, JobMaker_Form.ChkList_3_yes_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q3No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q3Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            ' 5.VONIC -------------------------------------------------------
            If JobMaker_Form.ChkList_5_no_RadioButton.Checked Then
                usr_nameManager_chkList_Q5_chkBoxState =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Q5No_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
            ElseIf JobMaker_Form.ChkList_5_nstd_RadioButton.Checked Then
                usr_nameManager_chkList_Q5_chkBoxState =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Q5NoStd_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
            ElseIf JobMaker_Form.ChkList_5_std_RadioButton.Checked Then
                usr_nameManager_chkList_Q5_chkBoxState =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Q5Std_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
            End If
            '------------------------------------------------------- 5.VONIC 

            usr_nameManager_chkList_Q6 =
                chkBoxStateRead(JobMaker_Form.ChkList_6_no_RadioButton, JobMaker_Form.ChkList_6_yes_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q6No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q6Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_chkList_Q6_yes =
                chkBoxStateRead(JobMaker_Form.ChkList_6_yesChk_RadioButton, JobMaker_Form.ChkList_6_yesItem_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q6YesChk_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q6YesItem_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_chkList_Q7 =
                chkBoxStateRead(JobMaker_Form.ChkList_7_no_RadioButton, JobMaker_Form.ChkList_7_yes_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q7No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q7Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_chkList_Q8 =
                chkBoxStateRead(JobMaker_Form.ChkList_8_no_RadioButton, JobMaker_Form.ChkList_8_yes_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q8No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q8Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_chkList_Q9 =
                chkBoxStateRead(JobMaker_Form.ChkList_9_no_RadioButton, JobMaker_Form.ChkList_9_yes_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q9No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Q9Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            ' 程式變更/2.使用裝置　--------------------------------------------
            If JobMaker_Form.PrmList_2_test_CheckBox.Checked Then
                usr_nameManager_prgm_2_Test =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_2_Test_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
            ElseIf JobMaker_Form.PrmList_2_COP_CheckBox.Checked Then
                usr_nameManager_prgm_2_COP =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_2_COP_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
            ElseIf JobMaker_Form.PrmList_2_Tower_CheckBox.Checked Then
                usr_nameManager_prgm_2_Tower =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_2_Tower_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
            ElseIf JobMaker_Form.PrmList_2_Other_CheckBox.Checked Then
                usr_nameManager_prgm_2_Other =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_2_Other_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
            End If
            '-------------------------------------------- 程式變更/2.使用裝置　

            ' 程式變更/3.檢查方法　--------------------------------------------
            If JobMaker_Form.PrmList_3_debug_CheckBox.Checked Then
                usr_nameManager_prgm_3_Debug =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_3_Debug_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)

            ElseIf JobMaker_Form.PrmList_3_test_CheckBox.Checked Then
                usr_nameManager_prgm_3_Test =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_3_Test_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)

            ElseIf JobMaker_Form.PrmList_3_confirm_CheckBox.Checked Then
                usr_nameManager_prgm_3_CFM =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_3_CFM_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)

            ElseIf JobMaker_Form.PrmList_3_excute_CheckBox.Checked Then
                usr_nameManager_prgm_3_EXE =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_3_EXE_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)

            ElseIf JobMaker_Form.PrmList_3_other_Checkbox.Checked Then
                usr_nameManager_prgm_3_Other =
                    get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_3_Other_ChkBox,
                                                  get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)

            End If
            '-------------------------------------------- 程式變更/3.檢查方法　

            ' 程式變更/4.檢查結果　-------------------------------------------------------------------------------------------------------------
            usr_nameManager_prgm_Auto =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no1_RadioButton, JobMaker_Form.PrmList_4_yes1_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_1No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_1Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_Output =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no2_RadioButton, JobMaker_Form.PrmList_4_yes2_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_2No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_2Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_INI =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no3_RadioButton, JobMaker_Form.PrmList_4_yes3_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_3No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_3Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_Case =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no4_RadioButton, JobMaker_Form.PrmList_4_yes4_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_4No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_4Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))
            usr_nameManager_prgm_IF =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no5_RadioButton, JobMaker_Form.PrmList_4_yes5_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_5No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_5Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_Loop =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no6_RadioButton, JobMaker_Form.PrmList_4_yes6_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_6No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_6Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_Range =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no7_RadioButton, JobMaker_Form.PrmList_4_yes7_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_7No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_7Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_Casting =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no8_RadioButton, JobMaker_Form.PrmList_4_yes8_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_8No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_8Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_0 =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no9_RadioButton, JobMaker_Form.PrmList_4_yes9_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_9No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_9Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_Count =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no10_RadioButton, JobMaker_Form.PrmList_4_yes10_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_10No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_10Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_Address =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no11_RadioButton, JobMaker_Form.PrmList_4_yes11_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_11No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_11Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))

            usr_nameManager_prgm_Custom =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no12_RadioButton, JobMaker_Form.PrmList_4_yes12_RadioButton,
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_12No_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_12Yes_ChkBox,
                                                              get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                              get_NameManager.SQLite_ToolDBMS_Name))
            '------------------------------------------------------------------------------------------------------------- 程式變更/4.檢查結果　

            '取得各check box content中的名稱
            usr_chkList_PA_year = JobMaker_Form.ChkList_PaSheet_DateTimePicker.Value.Year.ToString()
            usr_chkList_PA_month = JobMaker_Form.ChkList_PaSheet_DateTimePicker.Value.Month.ToString()
            usr_chkList_PA_date = JobMaker_Form.ChkList_PaSheet_DateTimePicker.Value.Day.ToString()
            usr_chkList_OS_year = JobMaker_Form.ChkList_OS_DateTimePicker.Value.Year.ToString()
            usr_chkList_OS_month = JobMaker_Form.ChkList_OS_DateTimePicker.Value.Month.ToString()
            usr_chkList_OS_date = JobMaker_Form.ChkList_OS_DateTimePicker.Value.Day.ToString()
            usr_chkList_CFM_year = JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Year.ToString()
            usr_chkList_CFM_month = JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Month.ToString()
            usr_chkList_CFM_date = JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Day.ToString()
            usr_chkList_ELE_year = JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Year.ToString()
            usr_chkList_ELE_month = JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Month.ToString()
            usr_chkList_ELE_date = JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Day.ToString()

            usr_chkList_Q1_YCont = JobMaker_Form.ChkList_1_yes_Content_TextBox.Name     '檢查項目1  有   討論內容
            usr_chkList_Q1_YResult = JobMaker_Form.ChkList_1_yes_result_TextBox.Name    '檢查項目1  有   結果
            usr_chkList_Q2_YCont = JobMaker_Form.ChkList_2_yes_Content_TextBox.Name     '檢查項目2  有   討論結果
            usr_chkList_Q2_YResult = JobMaker_Form.ChkList_2_yes_Result_TextBox.Name    '檢查項目2  有   討論結果
            usr_chkList_Q3_YMan = JobMaker_Form.ChkList_3_yes_Man_TextBox.Name          '檢查項目3  有   討論者
            usr_chkList_Q3_YCont = JobMaker_Form.ChkList_3_yes_Content_TextBox.Name     '檢查項目3  有   討論內容
            usr_chkList_Q3_YResult = JobMaker_Form.ChkList_3_yes_Result_TextBox.Name    '檢查項目3  有   討論結果
            usr_chkList_Q4_MMIC = JobMaker_Form.ChkList_4_ObjName_TextBox.Name             '檢查項目4  有   MMIC
            usr_chkList_Q4_MMICBase = JobMaker_Form.ChkList_4_ObjBase_TextBox.Name     '檢查項目4  有   MMIC BASE
            usr_chkList_Q4_SV = JobMaker_Form.ChkList_4_SV_TextBox.Name                 '檢查項目4  有   SV
            usr_chkList_Q4_SVBase = JobMaker_Form.ChkList_4_SVBase_TextBox.Name         '檢查項目4  有   SV BASE
            usr_chkList_Q5_StdCont = JobMaker_Form.ChkList_5_std_Content_TextBox.Name   '檢查項目5  有   標準內容
            usr_chkList_Q5_nStdCont = JobMaker_Form.ChkList_5_nstd_Content_TextBox.Name '檢查項目5  有   工直內容
            usr_chkList_Q6_YCont = JobMaker_Form.ChkList_6_yes_Content_TextBox.Name     '檢查項目6  有   檢驗項目
            usr_chkList_Q7_YCont = JobMaker_Form.ChkList_7_yes1_content_TextBox.Name    '檢查項目7  有   文書No
            usr_prgm_Reason = JobMaker_Form.PrmList_1_reason_TextBox.Name               '程式變更理由    
            usr_prgm_2_testCont = JobMaker_Form.PrmList_2_COP_TextBox.Name              '程式變更        測試裝置
            usr_prgm_2_CopCont = JobMaker_Form.PrmList_2_test_TextBox.Name              '程式變更理由     控制盤 
            usr_prgm_2_TowerCont = JobMaker_Form.PrmList_2_tower_TextBox.Name           '程式變更理由     研修測試塔
            usr_prgm_2_OtherCont = JobMaker_Form.PrmList_2_other_TextBox.Name           '程式變更理由     其他  
            usr_prgm_3_otherCont = JobMaker_Form.PrmList_3_other_TextBox.Name           '程式變更理由     其他
            usr_prgm_4_testCont = JobMaker_Form.PrmList_4_content12_TextBox.Name        '程式變更理由     測試內容 


            Dim usrChkList_arr, usrPrgm_arr As String()
            usrChkList_arr = {usr_nameManager_chkList_Q1, usr_chkList_Q1_YCont, usr_chkList_Q1_YResult,
                              usr_nameManager_chkList_Q2, usr_chkList_Q2_YCont, usr_chkList_Q2_YResult,
                              usr_nameManager_chkList_Q3, usr_chkList_Q3_YMan, usr_chkList_Q3_YCont, usr_chkList_Q3_YResult,
                              usr_chkList_Q4_MMIC, usr_chkList_Q4_MMICBase, usr_chkList_Q4_SV, usr_chkList_Q4_SVBase,
                              usr_nameManager_chkList_Q5_chkBoxState, usr_chkList_Q5_StdCont, usr_chkList_Q5_nStdCont,
                              usr_nameManager_chkList_Q6, usr_nameManager_chkList_Q6_yes, usr_chkList_Q6_YCont,
                              usr_nameManager_chkList_Q7, usr_chkList_Q7_YCont,
                              usr_nameManager_chkList_Q8,
                              usr_nameManager_chkList_Q9,
                              usr_nameManager_chkList_PA, usr_nameManager_chkList_OS, usr_nameManager_chkList_CFM, usr_nameManager_chkList_ELE,
                              usr_chkList_PA_year, usr_chkList_PA_month, usr_chkList_PA_date,
                              usr_chkList_CFM_year, usr_chkList_CFM_month, usr_chkList_CFM_date,
                              usr_chkList_ELE_year, usr_chkList_ELE_month, usr_chkList_ELE_date,
                              usr_chkList_OS_year, usr_chkList_OS_month, usr_chkList_OS_date}
            usrPrgm_arr = {usr_prgm_Reason, usr_nameManager_prgm_2_Test, usr_prgm_2_testCont,
                           usr_nameManager_prgm_2_COP, usr_prgm_2_CopCont, usr_nameManager_prgm_2_Tower,
                           usr_prgm_2_TowerCont, usr_nameManager_prgm_2_Other, usr_prgm_2_OtherCont,
                           usr_nameManager_prgm_3_Debug, usr_nameManager_prgm_3_Test, usr_nameManager_prgm_3_CFM,
                           usr_nameManager_prgm_3_EXE, usr_nameManager_prgm_3_Other, usr_prgm_3_otherCont,
                           usr_prgm_4_testCont,
                           usr_nameManager_prgm_Auto, usr_nameManager_prgm_Output, usr_nameManager_prgm_INI,
                           usr_nameManager_prgm_Case, usr_nameManager_prgm_IF, usr_nameManager_prgm_Loop,
                           usr_nameManager_prgm_Range, usr_nameManager_prgm_Casting, usr_nameManager_prgm_0,
                           usr_nameManager_prgm_Count, usr_nameManager_prgm_Address, usr_nameManager_prgm_Custom}

            Try
                '輸入相對應的check list值
                If JobMaker_Form.Use_ChkList_CheckBox.CheckState Then
                    For Each i_chkListStr In usrChkList_arr
                        If i_chkListStr <> "" Then
                            Try
                                Select Case i_chkListStr
                                    '基本資料
                                    Case usr_nameManager_chkList_Q1
                                        chkboxWriteIn(usr_nameManager_chkList_Q1,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_chkList_Q1, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_Q2
                                        chkboxWriteIn(usr_nameManager_chkList_Q2,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_nameManager_chkList_Q2, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_Q3
                                        chkboxWriteIn(usr_nameManager_chkList_Q3,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_nameManager_chkList_Q3, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_Q5_chkBoxState
                                        chkboxWriteIn(usr_nameManager_chkList_Q5_chkBoxState,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_chkList_Q5, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_Q6
                                        chkboxWriteIn(usr_nameManager_chkList_Q6,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_nameManager_chkList_Q6, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_Q6_yes
                                        chkboxWriteIn(usr_nameManager_chkList_Q6_yes,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_nameManager_chkList_Q6_yes, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_Q7
                                        chkboxWriteIn(usr_nameManager_chkList_Q7,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_nameManager_chkList_Q7, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_Q8
                                        chkboxWriteIn(usr_nameManager_chkList_Q8,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_nameManager_chkList_Q8, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_Q9
                                        chkboxWriteIn(usr_nameManager_chkList_Q9,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_nameManager_chkList_Q9, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_PA
                                        chkboxWriteIn(usr_nameManager_chkList_PA,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_chkList_PA, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_OS
                                        chkboxWriteIn(usr_nameManager_chkList_OS,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_chkList_OS, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_CFM
                                        chkboxWriteIn(usr_nameManager_chkList_CFM,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_chkList_CFM, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    Case usr_nameManager_chkList_ELE
                                        chkboxWriteIn(usr_nameManager_chkList_ELE,
                                                      chkListP1_ShtName,
                                                      msExcel_workbook)
                                        'chkboxWriteIn(usr_chkList_ELE, ChangeLink.get_wkSht_ChkListP1_TextBox.Text, msExcel_workbook)
                                    'PA/OS/確認圖/電器的年月日
                                    Case usr_chkList_PA_year
                                        excelWriteIn_ForReverseState(usr_chkList_PA_year,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_PA_Year,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_PaSheet_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_PA_year, ChangeLink.get_ChkList_PAYear_TextBox.Text, JobMaker_Form.usr_PaSheet_CheckBox, msExcel_workbook)
                                    Case usr_chkList_PA_month
                                        excelWriteIn_ForReverseState(usr_chkList_PA_month,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_PA_Month,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_PaSheet_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_PA_month, ChangeLink.get_ChkList_PAMonth_TextBox.Text, JobMaker_Form.usr_PaSheet_CheckBox, msExcel_workbook)
                                    Case usr_chkList_PA_date
                                        excelWriteIn_ForReverseState(usr_chkList_PA_date,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_PA_Day,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_PaSheet_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_PA_date, ChangeLink.get_ChkList_PADay_TextBox.Text, JobMaker_Form.usr_PaSheet_CheckBox, msExcel_workbook)
                                    Case usr_chkList_OS_year
                                        excelWriteIn_ForReverseState(usr_chkList_OS_year,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_OS_Year,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_OS_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_OS_year, ChangeLink.get_ChkList_OSYear_TextBox.Text, JobMaker_Form.usr_Os_CheckBox, msExcel_workbook)

                                    Case usr_chkList_OS_month
                                        excelWriteIn_ForReverseState(usr_chkList_OS_month,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_OS_Month,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_OS_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_OS_month, ChangeLink.get_ChkList_OSMonth_TextBox.Text, JobMaker_Form.usr_Os_CheckBox, msExcel_workbook)
                                    Case usr_chkList_OS_date
                                        excelWriteIn_ForReverseState(usr_chkList_OS_date,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_OS_Day,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_OS_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_OS_date, ChangeLink.get_ChkList_OSDay_TextBox.Text, JobMaker_Form.usr_Os_CheckBox, msExcel_workbook)
                                    Case usr_chkList_CFM_year
                                        excelWriteIn_ForReverseState(usr_chkList_CFM_year,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_CFM_Year,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_Confirm_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_CFM_year, ChangeLink.get_ChkList_CFMYear_TextBox.Text, JobMaker_Form.usr_Confirm_CheckBox, msExcel_workbook)
                                    Case usr_chkList_CFM_month
                                        excelWriteIn_ForReverseState(usr_chkList_CFM_month,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_CFM_Month,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_Confirm_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_CFM_month, ChangeLink.get_ChkList_CFMMonth_TextBox.Text, JobMaker_Form.usr_Confirm_CheckBox, msExcel_workbook)
                                    Case usr_chkList_CFM_date
                                        excelWriteIn_ForReverseState(usr_chkList_CFM_date,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_CFM_Day,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_Confirm_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_CFM_date, ChangeLink.get_ChkList_CFMDay_TextBox.Text, JobMaker_Form.usr_Confirm_CheckBox, msExcel_workbook)
                                    Case usr_chkList_ELE_year
                                        excelWriteIn_ForReverseState(usr_chkList_ELE_year,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_ELE_Year,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_Elec_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_ELE_year, ChangeLink.get_ChkList_ELEYear_TextBox.Text, JobMaker_Form.usr_Elec_CheckBox, msExcel_workbook)
                                    Case usr_chkList_ELE_month
                                        excelWriteIn_ForReverseState(usr_chkList_ELE_month,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_ELE_Month,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_Elec_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_ELE_month, ChangeLink.get_ChkList_ELEMonth_TextBox.Text, JobMaker_Form.usr_Elec_CheckBox, msExcel_workbook)
                                    Case usr_chkList_ELE_date
                                        excelWriteIn_ForReverseState(usr_chkList_ELE_date,
                                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_ELE_Day,
                                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                                     JobMaker_Form.ChkList_Elec_CheckBox,
                                                                     msExcel_workbook)
                                        'excelWriteIn_ForReverseState(usr_chkList_ELE_date, ChangeLink.get_ChkList_ELEDay_TextBox.Text, JobMaker_Form.usr_Elec_CheckBox, msExcel_workbook)

                                    'Textbox內容寫入
                                    Case usr_chkList_Q1_YCont
                                        'CheckList > 1.主式樣有無不清楚 > 討論內容
                                        excelWriteIn(JobMaker_Form.ChkList_1_yes_Content_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q1Yes_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_1_yes_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q1_YResult
                                        'CheckList > 1.主式樣有無不清楚 > 結果
                                        excelWriteIn(JobMaker_Form.ChkList_1_yes_result_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q1Yes_Result,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_1_yes_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q2_YCont
                                        'CheckList > 2.有沒有發生問題 > 指出內容
                                        excelWriteIn(JobMaker_Form.ChkList_2_yes_Content_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q2Yes_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_2_yes_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q2_YResult
                                        'CheckList > 2.有沒有發生問題 > 結果
                                        excelWriteIn(JobMaker_Form.ChkList_2_yes_Result_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q2Yes_Result,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_2_yes_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q3_YMan
                                        'CheckList > 3.電氣圖有沒有不清楚 > 討論者
                                        excelWriteIn(JobMaker_Form.ChkList_3_yes_Man_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q3Yes_Man,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_3_yes_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q3_YCont
                                        'CheckList > 3.電氣圖有沒有不清楚 > 內容
                                        excelWriteIn(JobMaker_Form.ChkList_3_yes_Content_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q3Yes_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_3_yes_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q3_YResult
                                        'CheckList > 3.電氣圖有沒有不清楚 > 結論
                                        excelWriteIn(JobMaker_Form.ChkList_3_yes_Result_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q3Yes_Result,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_3_yes_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q4_MMIC
                                        'CheckList > 4.MMIC > 內容
                                        excelWriteIn(JobMaker_Form.ChkList_4_ObjName_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q4MMIC,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     msExcel_workbook)
                                    Case usr_chkList_Q4_MMICBase
                                        'CheckList > 4.MMIC Base > 內容
                                        excelWriteIn(JobMaker_Form.ChkList_4_ObjBase_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q4MmicBase,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     msExcel_workbook)
                                    Case usr_chkList_Q4_SV
                                        'CheckList > 4.SV > 內容
                                        excelWriteIn(JobMaker_Form.ChkList_4_SV_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q4SV,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     msExcel_workbook)
                                    Case usr_chkList_Q4_SVBase
                                        'CheckList > 4.SV Base > 內容
                                        excelWriteIn(JobMaker_Form.ChkList_4_SVBase_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q4SVmicBase,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     msExcel_workbook)
                                    Case usr_chkList_Q5_StdCont
                                        'CheckList > 5.VONIC > 標準內容
                                        excelWriteIn(JobMaker_Form.ChkList_5_std_Content_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q5Std_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_5_std_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q5_nStdCont
                                        'CheckList > 5.VONIC > 工直內容
                                        excelWriteIn(JobMaker_Form.ChkList_5_nstd_Content_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q5nStd_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_5_nstd_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q6_YCont
                                        'CheckList > 6.執行動作確認 > 檢驗項目內容
                                        excelWriteIn(JobMaker_Form.ChkList_6_yes_Content_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q6Yes_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_6_yes_RadioButton,
                                                     msExcel_workbook)
                                    Case usr_chkList_Q7_YCont
                                        'CheckList > 7.參考資料 > 文書NO
                                        excelWriteIn(JobMaker_Form.ChkList_7_yes1_content_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Q7Yes_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.ChkList_7_yes_RadioButton,
                                                     msExcel_workbook)

                                End Select
                            Catch ex As Exception
                                JobMaker_Form.ResultOutput_TextBox.Text += ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_chkListStr}>{vbCrLf}")
                            End Try
                        End If
                    Next
                Else
                    JobMaker_Form.ResultFailOutput_TextBox.Text = ($"<提醒> Check List 分頁未輸出，原因:分頁未打勾")
                    JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.CheckList_TabPage
                    Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.CheckList_TabPage.Text}」未使用是否重來?"), vbYesNo)
                    If basic_result = DialogResult.Yes And msExcel_workbook IsNot Nothing Then
                        returnError_isPageRestart = True
                        'msExcel_workbook.Close()
                        'msExcel_app.Quit()
                    End If
                End If

                '輸入相對應的<程式變更>值
                If JobMaker_Form.Use_Program_CheckBox.CheckState Then
                    For Each i_prgmStr In usrPrgm_arr
                        If i_prgmStr <> "" Then
                            Try
                                Select Case i_prgmStr

                                    Case usr_nameManager_prgm_2_Test
                                        'Check List > 程式變更 > 2-1測試裝置
                                        chkboxWriteIn(usr_nameManager_prgm_2_Test,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_2_COP
                                        'Check List > 程式變更 > 2-2控制盤
                                        chkboxWriteIn(usr_nameManager_prgm_2_COP,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_2_Tower
                                        'Check List > 程式變更 > 2-3研修塔測試
                                        chkboxWriteIn(usr_nameManager_prgm_2_Tower,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_2_Other
                                        'Check List > 程式變更 > 2-4其他
                                        chkboxWriteIn(usr_nameManager_prgm_2_Other,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_3_Debug
                                        'Check List > 程式變更 > 3-1Debug
                                        chkboxWriteIn(usr_nameManager_prgm_3_Debug,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_3_Test
                                        'Check List > 程式變更 > 3-2內容測試
                                        chkboxWriteIn(usr_nameManager_prgm_3_Test,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_3_CFM
                                        'Check List > 程式變更 > 3-3一般動作確認
                                        chkboxWriteIn(usr_nameManager_prgm_3_CFM,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_3_EXE
                                        'Check List > 程式變更 > 3-4程式執行確認
                                        chkboxWriteIn(usr_nameManager_prgm_3_EXE,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_3_Other
                                        'Check List > 程式變更 > 3-5其他
                                        chkboxWriteIn(usr_nameManager_prgm_3_Other,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Auto
                                        'Check List > 程式變更 > 4-1全自動運轉
                                        chkboxWriteIn(usr_nameManager_prgm_Auto,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Output
                                        'Check List > 程式變更 > 4-2入出力點
                                        chkboxWriteIn(usr_nameManager_prgm_Output,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_INI
                                        'Check List > 程式變更 > 4-3初始化
                                        chkboxWriteIn(usr_nameManager_prgm_INI,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Case
                                        'Check List > 程式變更 > 4-4Case
                                        chkboxWriteIn(usr_nameManager_prgm_Case,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_IF
                                        'Check List > 程式變更 > 4-5IF
                                        chkboxWriteIn(usr_nameManager_prgm_IF,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Loop
                                        'Check List > 程式變更 > 4-6無限LOOP
                                        chkboxWriteIn(usr_nameManager_prgm_Loop,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Range
                                        'Check List > 程式變更 > 4-7定義範圍
                                        chkboxWriteIn(usr_nameManager_prgm_Range,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Casting
                                        'Check List > 程式變更 > 4-8Casting
                                        chkboxWriteIn(usr_nameManager_prgm_Casting,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_0
                                        'Check List > 程式變更 > 4-9 0除式子
                                        chkboxWriteIn(usr_nameManager_prgm_0,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Count
                                        'Check List > 程式變更 > 4-10 運算子
                                        chkboxWriteIn(usr_nameManager_prgm_Count,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Address
                                        'Check List > 程式變更 > 4-11 分配Address
                                        chkboxWriteIn(usr_nameManager_prgm_Address,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_nameManager_prgm_Custom
                                        'Check List > 程式變更 > 4-12 客戶實現要求
                                        chkboxWriteIn(usr_nameManager_prgm_Custom,
                                                      chkListP2_ShtName,
                                                      msExcel_workbook)
                                    Case usr_prgm_Reason
                                        'Check List > 程式變更 > 1 ROM變更理由
                                        excelWriteIn(JobMaker_Form.PrmList_1_reason_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_1_reason,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     msExcel_workbook)
                                    Case usr_prgm_2_testCont
                                        'Check List > 程式變更 > 2-1測試裝置
                                        excelWriteIn(JobMaker_Form.PrmList_2_COP_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_2_Test_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.PrmList_2_test_CheckBox,
                                                     msExcel_workbook)
                                    Case usr_prgm_2_CopCont
                                        'Check List > 程式變更 > 2-2控制盤
                                        excelWriteIn(JobMaker_Form.PrmList_2_test_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_2_COP_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.PrmList_2_COP_CheckBox,
                                                     msExcel_workbook)
                                    Case usr_prgm_2_TowerCont
                                        'Check List > 程式變更 > 2-3研修測試塔
                                        excelWriteIn(JobMaker_Form.PrmList_2_tower_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_2_Tower_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.PrmList_2_Tower_CheckBox,
                                                     msExcel_workbook)
                                    Case usr_prgm_2_OtherCont
                                        'Check List > 程式變更 > 2-4 其他
                                        excelWriteIn(JobMaker_Form.PrmList_2_other_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_2_Other_Content,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.PrmList_2_Other_CheckBox,
                                                     msExcel_workbook)
                                    Case usr_prgm_3_otherCont
                                        'Check List > 程式變更 > 3-1 其他
                                        excelWriteIn(JobMaker_Form.PrmList_3_other_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_3_OtherContent,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     JobMaker_Form.PrmList_3_other_Checkbox,
                                                     msExcel_workbook)
                                    Case usr_prgm_4_testCont
                                        'Check List > 程式變更 > 4 測試內容
                                        excelWriteIn(JobMaker_Form.PrmList_4_content12_TextBox.Text,
                                                     get_NameManager.read_DbmsData(get_NameManager.ChkList_Prgm_4_TestContent,
                                                                                   get_NameManager.SQLite_tableName_NameManager_CheckList,
                                                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                                                   get_NameManager.SQLite_ToolDBMS_Name),
                                                     msExcel_workbook)
                                End Select
                            Catch ex As Exception
                                JobMaker_Form.ResultOutput_TextBox.Text += ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_prgmStr}>{vbCrLf}")
                            End Try
                        End If
                    Next
                    JobMaker_Form.ResultOutput_TextBox.Text += "^^^^^^^^^^^ Check List ^^^^^^^^^^^" & vbCrLf
                Else
                    JobMaker_Form.ResultFailOutput_TextBox.Text += "$<提醒> 程式變更 分頁未輸出，原因:分頁未打勾"
                    JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.ProgramChange_TabPage
                    Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.ProgramChange_TabPage.Text}」未使用是否重來?"), vbYesNo)
                    If basic_result = DialogResult.Yes And msExcel_workbook IsNot Nothing Then
                        returnError_isPageRestart = True
                        'msExcel_workbook.Close()
                        'msExcel_app.Quit()
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            returnError_specName = "" '釋放回傳錯誤的值
        Else
            MsgBox("Check List/Program分頁左上角的CheckBox沒有勾選", MsgBoxStyle.Exclamation, "Fail Message")
        End If
    End Sub



    ''' <summary>
    ''' [自動生成的控制項寫入Excel中]
    ''' </summary>
    ''' <param name="mNumericUpDown_num">NumericUpDown.value控制項的數量</param>
    ''' <param name="specName">取得nameManger的開始行數的名稱</param>
    ''' <param name="specName_Array">取得需要寫入Excel中的自動生成控制項的Name Manager陣列</param>
    ''' <param name="mPanel">Panel</param>
    ''' <param name="dyCtrl_ArrayCount">取得自動生成控制項的數量</param>
    ''' <param name="dyCtrl_Array">取得自動生成控制項的名稱Name</param>
    ''' <param name="msExcel_workbook"></param>
    Overloads Sub dynamicControl_writeInExcel(mNumericUpDown_num As Integer, specName As String,
                                              specName_Array As Array,
                                              mPanel As Control,
                                              mSpec_Stored As Spec_StoredJobData.LoadStored_PanelType,
                                              dyCtrl_ArrayCount As Integer, dyCtrl_Array As Array,
                                              msExcel_workbook As Excel.Workbook)

        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData

        Dim startWorksheet_name As String
        Dim startCell_Row, startCell_Col As Integer
        Dim startRange_Row As Integer
        Dim startRange_Col As String
        Dim prk_Row, prk_Col, temp_prk_Row As Integer
        Dim merge_num As Integer
        '取得 名稱管理員specName 的Row
        startCell_Row =
            msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(specName,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)).RefersToRange.Row '號機名是第n行
        '取得 名稱管理員specName 的Col
        startCell_Col =
            msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(specName,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)).RefersToRange.Column '號機名是第n行
        '取得 目前使用的worksheet名稱
        startWorksheet_name = msExcel_workbook.Names.Item(specName).RefersToRange.Worksheet.Name

        '取得 名稱管理員specName Range的頭例如A4的4
        startRange_Row = startCell_Row
        startRange_Col =
            getMathOnExcel.convertColumn_fromIntToString(startCell_Col)
        '取得該合併儲存格的數量
        merge_num =
            msExcel_workbook.Worksheets(startWorksheet_name).range(startRange_Col & startRange_Row).MergeArea.Rows.Count

        '從第startCell_Row列往下20行找合併儲存格
        For merge_i As Integer = startCell_Row To startCell_Row + 20
            '找到非合併儲存格就插入 Row列，數量以mNumericUpDown_num為主
            If Not msExcel_workbook.Worksheets(startWorksheet_name).Cells(merge_i, startCell_Col).MergeCells Then
                For in_i = 1 To mNumericUpDown_num - 1
                    '複製後
                    msExcel_workbook.Worksheets(startWorksheet_name).Range($"{startCell_Row + merge_num}:{startCell_Row + 2 * merge_num - 1}").Copy
                    '插入
                    msExcel_workbook.Worksheets(startWorksheet_name).Range($"{merge_i}:{merge_i}").Insert
                Next
                Exit For
            End If
        Next

        '檢查Panel中有幾個控制項就跑幾次
        For Each tempCtrl As Control In mPanel.Controls '填入電梯的相關資訊
            '如果為判斷單一個Panel就跑 > SingleLayer_Panel
            '如果判斷為兩個  Panel就跑 > DoubleLayer_Panel
            If mSpec_Stored = mSpec_Stored.SingleLayer_Panel Then
                For lift_i = 1 To mNumericUpDown_num
                    For lift_j = 1 To dyCtrl_ArrayCount
                        '檢查控制項名稱是否符合需求的(dyCtrl_Array)
                        If tempCtrl.Name = $"{dyCtrl_Array(lift_j - 1)}_{lift_i}" Then
                            '取得欄、行，每執行完一次就會更新"行"的值 --------------------------------------------------
                            prk_Col = msExcel_workbook.Names.Item(specName_Array(lift_j - 1)).RefersToRange.Column '行
                            prk_Row = msExcel_workbook.Names.Item(specName_Array(lift_j - 1)).RefersToRange.Row '列
                            prk_Row += lift_i * merge_num

                            msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_Row, prk_Col).Value = tempCtrl.Text
                            prk_Row = temp_prk_Row
                            '-------------------------------------------------- 取得欄、行，每執行完一次就會更新"行"的值 
                        End If
                    Next
                Next
            ElseIf mSpec_Stored = mSpec_Stored.DoubleLayer_Panel Then
                For Each tempCtrl_Double In tempCtrl.Controls
                    For lift_i = 1 To mNumericUpDown_num
                        For lift_j = 1 To dyCtrl_ArrayCount
                            If tempCtrl_Double.Name = $"{dyCtrl_Array(lift_j - 1)}_{lift_i}" Then
                                prk_Col = msExcel_workbook.Names.Item(specName_Array(lift_j - 1)).RefersToRange.Column '行
                                prk_Row = msExcel_workbook.Names.Item(specName_Array(lift_j - 1)).RefersToRange.Row '列
                                prk_Row += lift_i * merge_num

                                msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_Row, prk_Col).Value = tempCtrl_Double.Text
                                prk_Row = temp_prk_Row

                            End If
                        Next
                    Next
                Next
            End If
        Next

    End Sub

    ''' <summary>
    ''' [自動生成的控制項寫入Excel中 > Spec 基本]
    ''' </summary>
    ''' <param name="mNumericUpDown_num"></param>
    ''' <param name="specName"></param>
    ''' <param name="specName_Array"></param>
    ''' <param name="mPanel"></param>
    ''' <param name="dyCtrl_ArrayCount"></param>
    ''' <param name="dyCtrl_Array"></param>
    ''' <param name="msExcel_workbook"></param>
    Overloads Sub dynamicControl_writeInExcel_SpecBasic(mNumericUpDown_num As Integer, specName As String,
                                                        specName_Array As Array,
                                                        mPanel As Control,
                                                        dyCtrl_ArrayCount As Integer, dyCtrl_Array As Array,
                                                        msExcel_workbook As Excel.Workbook)

        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData

        Dim startWorksheet_name As String
        Dim startCell_Row, startCell_Col As Integer
        Dim startRange_Row As Integer
        Dim startRange_Col As String
        Dim prk_Row, prk_Col, temp_prk_Row As Integer
        Dim merge_num As Integer

        '取得 名稱管理員specName 的Row
        startCell_Row =
            getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, specName)
        '取得 名稱管理員specName 的Col
        startCell_Col =
            getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, specName)
        '取得 目前使用WorkSheet的名稱
        startWorksheet_name =
            getMathOnExcel.getWorksheetName_fromNameManager(msExcel_workbook, specName)

        '取得 名稱管理員specName Range的頭例如A4的4
        startRange_Row = startCell_Row
        '取得 名稱管理員specName Range的尾例如A4的A
        startRange_Col =
            getMathOnExcel.convertColumn_fromIntToString(startCell_Col)

        '取得該合併儲存格的數量
        merge_num =
            msExcel_workbook.Worksheets(startWorksheet_name).range(startRange_Col & startRange_Row).MergeArea.Rows.Count
        For merge_i As Integer = startCell_Row To startCell_Row + 20 '找合併格子
            If Not msExcel_workbook.Worksheets(startWorksheet_name).Cells(merge_i, startCell_Col).MergeCells Then
                'If mNumericUpDown_num > 2 Then '號機數量大於2台就插入col
                For in_i = 1 To mNumericUpDown_num - 1
                    msExcel_workbook.Worksheets(startWorksheet_name).Range($"{startCell_Row + merge_num}:{startCell_Row + 2 * merge_num - 1}").Copy
                    msExcel_workbook.Worksheets(startWorksheet_name).Range($"{merge_i}:{merge_i}").Insert
                Next
                'End If
                Exit For
            End If
        Next

        For Each tempCtrl As Control In mPanel.Controls '填入電梯的相關資訊
            For lift_i = 1 To mNumericUpDown_num
                For lift_j = 1 To dyCtrl_ArrayCount
                    If tempCtrl.Name = $"{dyCtrl_Array(lift_j - 1)}_{lift_i}" Then
                        prk_Col = msExcel_workbook.Names.Item(specName_Array(lift_j - 1)).RefersToRange.Column '行
                        prk_Row = msExcel_workbook.Names.Item(specName_Array(lift_j - 1)).RefersToRange.Row '列
                        prk_Row += lift_i * merge_num



                        Dim tempCtrlText As String
                        tempCtrlText = ""

                        If tempCtrl.Name = $"{JobMaker_Form.Spec_TopFL_TextBox.Name}_{lift_i}" Then
                            For Each realFL As Control In mPanel.Controls
                                If realFL.Name = $"{JobMaker_Form.Spec_TopFL_Real_TextBox.Name}_{lift_i}" Then
                                    tempCtrlText = $"{tempCtrl.Text} {realFL.Text}"
                                End If
                            Next
                        ElseIf tempCtrl.Name = $"{JobMaker_Form.Spec_BtmFL_TextBox.Name}_{lift_i}" Then
                            For Each realFL As Control In mPanel.Controls
                                If realFL.Name = $"{JobMaker_Form.Spec_BtmFL_Real_TextBox.Name}_{lift_i}" Then
                                    tempCtrlText = $"{tempCtrl.Text} {realFL.Text}"
                                End If
                            Next
                        Else
                            tempCtrlText = tempCtrl.Text
                        End If
                        msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_Row, prk_Col).Value = tempCtrlText
                        prk_Row = temp_prk_Row

                    End If
                Next
            Next
        Next

    End Sub

    ''' <summary>
    ''' [自動生成的控制項寫入Excel中 > MMIC]
    ''' </summary>
    ''' <param name="mNumericUpDown1"></param>
    ''' <param name="mNumericUpDown2"></param>
    ''' <param name="specName"></param>
    ''' <param name="specName1_Array"></param>
    ''' <param name="specName2_Array"></param>
    ''' <param name="mPanel1"></param>
    ''' <param name="mPanel2"></param>
    ''' <param name="dyCtrl1_ArrayCount"></param>
    ''' <param name="dyCtrl1_Array"></param>
    ''' <param name="dyCtrl2_ArrayCount"></param>
    ''' <param name="dyCtrl2_Array"></param>
    ''' <param name="msExcel_workbook"></param>
    Overloads Sub dynamicControl_writeInExcel_MMIC(mNumericUpDown1 As NumericUpDown, mNumericUpDown2 As NumericUpDown,
                                                   specName As String,
                                                   specName1_Array As Array, specName2_Array As Array,
                                                   mPanel1 As Control, mPanel2 As Control,
                                                   dyCtrl1_ArrayCount As Integer, dyCtrl1_Array As Array,
                                                   dyCtrl2_ArrayCount As Integer, dyCtrl2_Array As Array,
                                                   msExcel_workbook As Excel.Workbook)

        '針對flashRom 與 EEPROM 比較 取最大值 -----------------------------------
        Dim mNumericUpDown1_num, mNumericUpDown2_num, mNumeric_Max As Integer
        mNumericUpDown1_num = mNumericUpDown1.Value
        mNumericUpDown2_num = mNumericUpDown2.Value
        If mNumericUpDown1_num > mNumericUpDown2_num Then
            mNumeric_Max = mNumericUpDown1_num
        ElseIf mNumericUpDown1_num < mNumericUpDown2_num Then
            mNumeric_Max = mNumericUpDown2_num
        Else
            mNumeric_Max = mNumericUpDown1_num
        End If
        '----------------------------------- 針對flashRom 與 EEPROM 比較 取最大值 



        Dim startWorksheet_name As String
        Dim startCell_Row, startCell_Col As Integer
        Dim startRange_Row As Integer
        Dim startRange_Col As String
        Dim prk_Row, prk_Col, temp_prk_Row As Integer
        Dim merge_num As Integer
        '取得 名稱管理員specName 的Row
        startCell_Row =
            getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, specName)
        '取得 名稱管理員specName 的Col
        startCell_Col =
            getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, specName)
        '取得 目前WorkSheet名稱
        startWorksheet_name =
            getMathOnExcel.getWorksheetName_fromNameManager(msExcel_workbook, specName)

        '取得 名稱管理員specName Range的頭例如A4的4
        startRange_Row = startCell_Row
        '取得 名稱管理員specName Range的尾例如A4的A
        startRange_Col =
           getMathOnExcel.convertColumn_fromIntToString(startCell_Col)

        '取得該合併儲存格的數量
        merge_num =
            getMathOnExcel.getRowCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, startWorksheet_name, startRange_Row, startRange_Col)

        For merge_i As Integer = startCell_Row To startCell_Row + 20 '找合併格子
            If Not msExcel_workbook.Worksheets(startWorksheet_name).Cells(merge_i, startCell_Col).MergeCells Then
                'If mNumericUpDown_num > 2 Then '號機數量大於2台就插入col
                For in_i = 1 To mNumeric_Max - 1
                    msExcel_workbook.Worksheets(startWorksheet_name).Range($"{startCell_Row + merge_num}:{startCell_Row + 2 * merge_num - 1}").Copy
                    msExcel_workbook.Worksheets(startWorksheet_name).Range($"{merge_i}:{merge_i}").Insert
                Next
                Exit For
            End If
        Next

        For Each tempCtrl1 In mPanel1.Controls '填入電梯的相關資訊
            For lift_i As Integer = 1 To mNumericUpDown1_num
                For lift_j As Integer = 1 To dyCtrl1_ArrayCount
                    If tempCtrl1.Name = $"{dyCtrl1_Array(lift_j - 1)}_{lift_i}" Then
                        prk_Col = msExcel_workbook.Names.Item(specName1_Array(lift_j - 1)).RefersToRange.Column '行
                        prk_Row = msExcel_workbook.Names.Item(specName1_Array(lift_j - 1)).RefersToRange.Row '列
                        prk_Row += lift_i * merge_num

                        msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_Row, prk_Col).Value = tempCtrl1.Text
                        prk_Row = temp_prk_Row

                    End If
                Next
            Next
        Next

        prk_Row = 0
        prk_Col = 0

        For Each tempCtrl2 In mPanel2.Controls
            For lift_i As Integer = 1 To mNumericUpDown2_num
                For lift_j As Integer = 1 To dyCtrl2_ArrayCount
                    If tempCtrl2.Name = $"{dyCtrl2_Array(lift_j - 1)}_{lift_i}" Then
                        prk_Col = msExcel_workbook.Names.Item(specName2_Array(lift_j - 1)).RefersToRange.Column '行
                        prk_Row = msExcel_workbook.Names.Item(specName2_Array(lift_j - 1)).RefersToRange.Row '列
                        prk_Row += lift_i * merge_num

                        msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_Row, prk_Col).Value = tempCtrl2.Text
                        prk_Row = temp_prk_Row

                    End If
                Next
            Next
        Next
    End Sub

    ''' <summary>
    ''' [仕樣 > Basic]
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    Public Sub Spec_SPEC_Basic(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application)
        Try
            If JobMaker_Form.Use_SpecBasic_CheckBox.Checked Then

                Dim spec_car_name As String
                spec_car_name =
                    get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_NAME,
                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
                Dim spec_car_no As String
                spec_car_no =
                     get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_NO,
                                                   get_NameManager.SQLite_tableName_NameManager_TW,
                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                   get_NameManager.SQLite_ToolDBMS_Name)
                Dim spec_car_ope As String
                spec_car_ope =
                     get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_OPE,
                                                   get_NameManager.SQLite_tableName_NameManager_TW,
                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                   get_NameManager.SQLite_ToolDBMS_Name)
                Dim spec_car_topfl As String
                spec_car_topfl =
                     get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_TOPFL,
                                                   get_NameManager.SQLite_tableName_NameManager_TW,
                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                   get_NameManager.SQLite_ToolDBMS_Name)
                Dim spec_car_btmfl As String
                spec_car_btmfl =
                     get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_BTMFL,
                                                   get_NameManager.SQLite_tableName_NameManager_TW,
                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                   get_NameManager.SQLite_ToolDBMS_Name)
                Dim spec_car_stop As String
                spec_car_stop =
                     get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_STOP,
                                                   get_NameManager.SQLite_tableName_NameManager_TW,
                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                   get_NameManager.SQLite_ToolDBMS_Name)
                Dim spec_car_speed As String
                spec_car_speed =
                     get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_SPEED,
                                                   get_NameManager.SQLite_tableName_NameManager_TW,
                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                   get_NameManager.SQLite_ToolDBMS_Name)
                Dim spec_car_flname As String
                spec_car_flname =
                     get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_FLNAME,
                                                   get_NameManager.SQLite_tableName_NameManager_TW,
                                                   get_NameManager.SQLite_connectionPath_Tool,
                                                   get_NameManager.SQLite_ToolDBMS_Name)

                Dim JM_Spec_AllBasic As String() = {spec_car_name, spec_car_no,
                                                    spec_car_ope, spec_car_topfl,
                                                    spec_car_btmfl, spec_car_stop,
                                                    spec_car_speed, spec_car_flname}

                'Spec 基本
                Dim dyCtrlName As DynamicControlName = New DynamicControlName
                dyCtrlName.JobMaker_LiftInfo()

                dynamicControl_writeInExcel_SpecBasic(JobMaker_Form.Spec_LiftNum_NumericUpDown.Value,
                                                      get_NameManager.SPEC_CAR_NAME,
                                                      JM_Spec_AllBasic,
                                                      JobMaker_Form.SpecBasic_LiftItem_Dynamic_Panel,
                                                      dyCtrlName.JobMaker_LiftInfoName_Array.Count,
                                                      dyCtrlName.JobMaker_LiftInfoName_Array,
                                                      msExcel_workbook)


                Dim spec_car_machine_type, spec_car_control_way,
                    spec_car_location As String

                spec_car_machine_type =
                    get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_MACHINE_TYPE,
                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
                spec_car_control_way =
                    get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_CONTROL_WAY,
                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
                spec_car_location =
                    get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_LOCATION,
                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                  get_NameManager.SQLite_ToolDBMS_Name)
                Dim JM_Spec_AllBasic_2() As String = {spec_car_machine_type, spec_car_control_way, spec_car_location}
                Dim jm_spec_allBasic2_i As String

                For Each jm_spec_allBasic2_i In JM_Spec_AllBasic_2
                    If jm_spec_allBasic2_i <> "" Then
                        Try
                            Select Case jm_spec_allBasic2_i
                                Case spec_car_machine_type
                                    returnError_specName = spec_car_machine_type
                                    '機種
                                    Dim JM_MACHINE_TYPE As String() = {spec_car_machine_type}
                                    Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
                                    dynamicControl_writeInExcel(JobMaker_Form.Spec_MachineType_NumericUpDown.Value,
                                                                spec_car_machine_type,
                                                                JM_MACHINE_TYPE,
                                                                JobMaker_Form.Spec_MachineType_Panel,
                                                                spec_stored.LoadStored_PanelType.SingleLayer_Panel,
                                                                {dyCtrlName.Spec_MachineType_ComboBox}.Count,
                                                                {dyCtrlName.Spec_MachineType_ComboBox},
                                                                msExcel_workbook)
                                Case spec_car_control_way
                                    returnError_specName = spec_car_control_way
                                    '控制方式 / PURPOSE目的方式
                                    Dim JM_CONTROL_WAY As String() = {spec_car_control_way}

                                    Dim JM_PURPOSE As String() = {get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_PURPOSE,
                                                                                                get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                get_NameManager.SQLite_connectionPath_Tool,
                                                                                                get_NameManager.SQLite_ToolDBMS_Name)}


                                    dynamicControl_writeInExcel_MMIC(JobMaker_Form.Spec_MachineType_NumericUpDown, JobMaker_Form.Spec_Purpose_NumericUpDown,
                                                                     spec_car_control_way,
                                                                     JM_CONTROL_WAY, JM_PURPOSE,
                                                                     JobMaker_Form.Spec_ControlWay_Panel, JobMaker_Form.Spec_Purpose_Panel,
                                                                     {dyCtrlName.Spec_ControlWay_ComboBox}.Count, {dyCtrlName.Spec_ControlWay_ComboBox},
                                                                     {dyCtrlName.Spec_Purpose_ComboBox}.Count, {dyCtrlName.Spec_Purpose_ComboBox},
                                                                     msExcel_workbook)
                                Case spec_car_location
                                    returnError_specName = spec_car_location
                                    '所在地
                                    msExcel_workbook.Names.Item(spec_car_location).RefersToRange.Cells.Value =
                                        JobMaker_Form.Basic_Local_ComboBox.Text
                            End Select
                        Catch ex As Exception
                            JobMaker_Form.ResultFailOutput_TextBox.Text +=
                                ($"<{JobMaker_Form.JMFileCho_Spec_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{jm_spec_allBasic2_i}>{vbCrLf}")
                        End Try
                    End If
                Next
            Else
                JobMaker_Form.ResultFailOutput_TextBox.Text += ("<提醒> 仕樣>基本 分頁未輸出，原因:分頁未打勾")
                JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.Basic_TabPage
                Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.Basic_TabPage.Text}」未使用是否重來?"), vbYesNo)
                If basic_result = DialogResult.Yes And msExcel_workbook IsNot Nothing Then
                    returnError_isPageRestart = True
                    'msExcel_workbook.Close()
                    'msExcel_app.Quit()
                End If
            End If
        Catch ex As Exception
            MsgBox($"Spec_SPEC_Basic funciton error : {ex.ToString}")
        End Try
    End Sub
    ''' <summary>
    ''' Job Maker >> 仕樣 (快速摺疊Code:CRTL+M+M)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="LiftNum"></param>
    ''' <param name="ContainNum"></param>
    Public Sub Spec_SPEC_TW(LiftNum As Integer, ContainNum As Integer,
                            msExcel_workbook As Excel.Workbook, msExcel_App As Excel.Application)

        '台灣仕樣
        If JobMaker_Form.Use_SpecTWIDU_CheckBox.Checked Or JobMaker_Form.Use_SpecTWFP17_CheckBox.Checked Then
            Dim usr_Spec_MachineType,
                usr_Spec_Auto_DR, usr_Spec_Cancell_call, usr_Spec_Lamp_Chk,
                usr_Spec_Cancell_Behind,
                usr_Spec_Auto_Fan, usr_Spec_CC_Cancel, usr_Spec_Auto_Pass,
                usr_Spec_Operation, usr_Spec_Install_Ope, usr_Spec_Indep_Ope, usr_Spec_UCMP, usr_Spec_HIN_CPI,
                usr_Spec_Fire_Ope, usr_Spec_Fireman, usr_Spec_Parking, usr_Spec_Seismic, usr_Spec_CPI,
                usr_Spec_Car_Gong, usr_Spec_Hall_Gong, usr_Spec_HPI, usr_Spec_Dr_Hold, usr_Spec_CRD,
                usr_Spec_Emer_Power, usr_Spec_Landic, usr_Spec_MLF_Return, usr_Spec_Vonic,
                usr_Spec_WHB, usr_Spec_Elvic, usr_Spec_HLL, usr_Spec_ATT, usr_Spec_Flood, usr_Spec_LS1M,
                usr_Spec_PRU, usr_Spec_FrontRear_DR As String


            usr_Spec_MachineType =
                JobMaker_Form.Spec_Base_ComboBox.Name
            usr_Spec_Auto_DR =
                JobMaker_Form.Spec_DRAuto_ComboBox.Name
            usr_Spec_Cancell_call =
                JobMaker_Form.Spec_CancellCall_ComboBox.Name
            usr_Spec_Cancell_Behind =
                JobMaker_Form.Spec_CancellBehind_ComboBox.Name
            usr_Spec_Lamp_Chk =
                JobMaker_Form.Spec_LampChk_ComboBox.Name
            usr_Spec_Auto_Fan =
                JobMaker_Form.Spec_AutoFan_ComboBox.Name
            usr_Spec_CC_Cancel =
                JobMaker_Form.Spec_CCCancell_ComboBox.Name
            usr_Spec_Auto_Pass =
                JobMaker_Form.Spec_AutoPass_ComboBox.Name
            'usr_Spec_Operation =
            '    JobMaker_Form.Spec_Operation_ComboBox.Name
            usr_Spec_Install_Ope =
                JobMaker_Form.Spec_install_ope_ComboBox.Name
            usr_Spec_Indep_Ope =
                JobMaker_Form.Spec_Indep_ComboBox.Name
            usr_Spec_UCMP =
                JobMaker_Form.Spec_UCMP_ComboBox.Name
            usr_Spec_HIN_CPI =
                JobMaker_Form.Spec_HinCpi_ComboBox.Name
            usr_Spec_Fire_Ope =
                JobMaker_Form.Spec_Fire_ComboBox.Name
            usr_Spec_Fireman =
                JobMaker_Form.Spec_Fireman_ComboBox.Name
            usr_Spec_Parking =
                JobMaker_Form.Spec_Parking_ComboBox.Name
            usr_Spec_Seismic =
                JobMaker_Form.Spec_Seismic_ComboBox.Name
            usr_Spec_CPI =
                JobMaker_Form.Spec_CPI_ComboBox.Name
            usr_Spec_Car_Gong =
                JobMaker_Form.Spec_CarGong_ComboBox.Name
            usr_Spec_Hall_Gong =
                JobMaker_Form.Spec_HallGong_ComboBox.Name
            usr_Spec_HPI =
                JobMaker_Form.Spec_HPIMsg_ComboBox.Name
            usr_Spec_Dr_Hold =
                JobMaker_Form.Spec_DrHold_ComboBox.Name
            usr_Spec_CRD =
                JobMaker_Form.Spec_CRD_ComboBox.Name
            usr_Spec_Emer_Power =
                JobMaker_Form.Spec_Emer_ComboBox.Name
            usr_Spec_Landic =
                JobMaker_Form.Spec_Landic_ComboBox.Name
            usr_Spec_MLF_Return =
                JobMaker_Form.Spec_MFLReturn_ComboBox.Name
            usr_Spec_Vonic =
                JobMaker_Form.Spec_Vonic_ComboBox.Name
            usr_Spec_WHB =
                JobMaker_Form.Spec_WCOB_ComboBox.Name
            usr_Spec_Elvic =
                JobMaker_Form.Spec_Elvic_ComboBox.Name
            usr_Spec_HLL =
                JobMaker_Form.Spec_HLL_ComboBox.Name
            usr_Spec_ATT =
                JobMaker_Form.Spec_ATT_ComboBox.Name
            usr_Spec_Flood =
                JobMaker_Form.Spec_Flood_ComboBox.Name
            usr_Spec_LS1M =
                JobMaker_Form.Spec_LS1M_ComboBox.Name
            usr_Spec_PRU =
                JobMaker_Form.Spec_PRU_ComboBox.Name
            usr_Spec_FrontRear_DR =
                JobMaker_Form.Spec_FrontRearDr_ComboBox.Name
            Dim usr_Spec_OpeSw As String
            usr_Spec_OpeSw =
                JobMaker_Form.Spec_OpeSw_ComboBox.Name

            Dim usrInput_TWSpec_arr() As String = {usr_Spec_MachineType,
                                                   usr_Spec_Auto_DR, usr_Spec_Cancell_call, usr_Spec_Lamp_Chk,
                                                   usr_Spec_Cancell_Behind, usr_Spec_Auto_Fan,
                                                    usr_Spec_CC_Cancel, usr_Spec_Auto_Pass,
                                                    usr_Spec_Install_Ope,
                                                   usr_Spec_Indep_Ope, usr_Spec_UCMP, usr_Spec_HIN_CPI, usr_Spec_Fire_Ope,
                                                   usr_Spec_Fireman, usr_Spec_Parking, usr_Spec_Seismic, usr_Spec_CPI,
                                                   usr_Spec_Car_Gong, usr_Spec_Hall_Gong, usr_Spec_HPI, usr_Spec_Dr_Hold,
                                                   usr_Spec_CRD, usr_Spec_Emer_Power, usr_Spec_Landic, usr_Spec_MLF_Return,
                                                   usr_Spec_Vonic, usr_Spec_WHB, usr_Spec_Elvic, usr_Spec_HLL,
                                                   usr_Spec_ATT, usr_Spec_Flood, usr_Spec_LS1M, usr_Spec_PRU,
                                                   usr_Spec_FrontRear_DR, usr_Spec_OpeSw}

            Dim with_val, without_val, no_val, nc_val As String
            with_val =
                msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_RESULT_WITH,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name)
                                            ).RefersToRange.Value '取得 有 內的文字內容
            without_val =
                msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_RESULT_WITHOUT,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name)
                                            ).RefersToRange.Value '取得 無 內的文字內容
            no_val =
                msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_NO,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name)
                                            ).RefersToRange.Value '取得 訊號NO 內的文字內容
            nc_val =
                msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_NC,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name)
                                            ).RefersToRange.Value '取得 訊號NC 內的文字內容
            Dim i_TWSpec_str As String
            For Each i_TWSpec_str In usrInput_TWSpec_arr
                If i_TWSpec_str <> "" Then
                    Try
                        Select Case i_TWSpec_str
                            ' 機種 ------------------------------------------------------------------------------------------------------
                            Case usr_Spec_MachineType
                                excelWriteIn(JobMaker_Form.Spec_Base_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SetTable_MACHINE_TYPE,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            ' 開門時限自動調節 ------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Auto_DR
                                excelWriteIn(JobMaker_Form.Spec_DRAuto_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_AUTO_DR,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                                If JobMaker_Form.Spec_DRAuto_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim spec_auto_dr_photoeye, spec_auto_dr_safety As String
                                    spec_auto_dr_photoeye =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_AUTO_DR_PHOTOEYE,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_auto_dr_safety =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_AUTO_DR_SAFETY,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    Dim pho_val, safety_val As String
                                    pho_val =
                                        msExcel_workbook.Names.Item(spec_auto_dr_photoeye
                                                                    ).RefersToRange.Value '取得 光電裝置 內的文字內容
                                    safety_val =
                                        msExcel_workbook.Names.Item(spec_auto_dr_safety
                                                                    ).RefersToRange.Value '取得 安全履 內的文字內容

                                    msExcel_workbook.Names.Item(spec_auto_dr_photoeye
                                                                ).RefersToRange.Cells.Font.Strikethrough = False

                                    msExcel_workbook.Names.Item(spec_auto_dr_safety
                                                                ).RefersToRange.Cells.Font.Strikethrough = False

                                    If JobMaker_Form.Spec_PhotoEye_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                        msExcel_workbook.Names.Item(spec_auto_dr_photoeye
                                                                    ).RefersToRange.Characters(InStr(pho_val, with_val), Len(with_val)).
                                                                    Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_auto_dr_photoeye
                                                                    ).RefersToRange.Characters(InStr(pho_val, without_val), Len(without_val)).
                                                                    Font.Strikethrough = True
                                    End If

                                    If JobMaker_Form.Spec_MechSafety_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                        msExcel_workbook.Names.Item(spec_auto_dr_safety
                                                                    ).RefersToRange.Characters(InStr(safety_val, with_val), Len(with_val)).
                                                                    Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_auto_dr_safety
                                                                    ).RefersToRange.Characters(InStr(safety_val, without_val), Len(without_val)).
                                                                    Font.Strikethrough = True
                                    End If

                                End If
                            '------------------------------------------------------------------------------------------------------ 開門時限自動調節 

                            ' 取消嬉戲呼叫 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Cancell_call
                                excelWriteIn(JobMaker_Form.Spec_CancellCall_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_CANCELL_CALL,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_CancellCall_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim spec_cancell_call_scob, spec_cancell_call_six As String
                                    spec_cancell_call_scob =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CANCELL_CALL_SCOB,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_cancell_call_six =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CANCELL_CALL_SIX,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    Dim scob_val, six_val As String
                                    scob_val =
                                        msExcel_workbook.Names.Item(spec_cancell_call_scob
                                                                    ).RefersToRange.Value '取得 SCOB 內的文字內容
                                    six_val =
                                        msExcel_workbook.Names.Item(spec_cancell_call_six
                                                                    ).RefersToRange.Value '取得 SCOB 內的文字內容

                                    If JobMaker_Form.Spec_SCOB_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                        msExcel_workbook.Names.Item(spec_cancell_call_scob
                                                                    ).RefersToRange.Characters(InStr(scob_val, with_val), Len(with_val)).
                                                                    Font.Strikethrough = True
                                        msExcel_workbook.Names.Item(spec_cancell_call_six
                                                                   ).RefersToRange.Characters(InStr(six_val, "副COB"), Len("副COB")).
                                                                   Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_cancell_call_scob
                                                                    ).RefersToRange.Characters(InStr(scob_val, without_val), Len(without_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                End If
                            '------------------------------------------------------------------------------------------------------ 取消嬉戲呼叫

                            ' 逆呼無效 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Cancell_Behind
                                excelWriteIn(JobMaker_Form.Spec_CancellBehind_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_CANCELL_BEHIND,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 逆呼無效

                            ' 燈點檢模式 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Lamp_Chk
                                excelWriteIn(JobMaker_Form.Spec_LampChk_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_LAMP_CHK,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 燈點檢模式


                            ' 風扇連動 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Auto_Fan
                                excelWriteIn(JobMaker_Form.Spec_AutoFan_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_AUTO_FAN,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_ION_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                    '離子除菌
                                    Dim spec_auto_fan_ion As String
                                    spec_auto_fan_ion =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_AUTO_FAN_ION,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    msExcel_workbook.Names.Item(spec_auto_fan_ion).RefersToRange.Cells.Font.Strikethrough = True

                                    Dim ion_row, ion_col As Integer
                                    ion_row =
                                        getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, spec_auto_fan_ion)
                                    'msExcel_workbook.Names.Item(spec_auto_fan_ion).RefersToRange.Row
                                    ion_col =
                                        getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, spec_auto_fan_ion)
                                    'msExcel_workbook.Names.Item(spec_auto_fan_ion).RefersToRange.Column

                                    Dim spec_shtName As String
                                    spec_shtName =
                                        getMathOnExcel.getWorksheetName_fromNameManager(msExcel_workbook, spec_auto_fan_ion)
                                    'msExcel_workbook.Names.Item(spec_auto_fan_ion).RefersToRange.Worksheet.Name
                                    For i = 0 To 3
                                        'If i <> 2 Then
                                        msExcel_workbook.Worksheets(spec_shtName).Cells(ion_row + i, ion_col).font.Strikethrough = True
                                        'End If
                                    Next
                                End If

                                '重要設定 ion
                                Dim ion_val As String
                                ion_val = ""
                                If JobMaker_Form.Spec_AutoFan_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                    ion_val = get_NameManager.TB_WITHOUT
                                Else
                                    If JobMaker_Form.Spec_ION_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                        ion_val = get_NameManager.TB_WITH
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_FAN_CONTENT,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    ElseIf JobMaker_Form.Spec_ION_ComboBox.Text = get_NameManager.TB_WITH Then
                                        ion_val = get_NameManager.TB_WITH & "(ION)"
                                    End If
                                End If

                                excelWriteIn(ion_val,
                                             get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_FAN,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 風扇連動


                            ' 車廂呼叫取消機能 ----------------------------------------------------------------------------------------------------
                            Case usr_Spec_CC_Cancel
                                excelWriteIn(JobMaker_Form.Spec_CCCancell_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_CC_CANCEL,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                                                            '---------------------------------------------------------------------------------------------------- 車廂呼叫取消機能

                            ' 自動滿員通過 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Auto_Pass
                                excelWriteIn(JobMaker_Form.Spec_AutoPass_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_AUTO_PASS,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 自動滿員通過


                            ' 操作方式 --------------------------------------------------------------------------------------------------------
                            'Case usr_Spec_Operation
                            '    excelWriteIn(JobMaker_Form.Spec_Operation_ComboBox.Text,
                            '                 get_NameManager.read_DbmsData(get_NameManager.SPEC_OPERATION,
                            '                                               get_NameManager.SQLite_tableName_NameManager_TW,
                            '                                               get_NameManager.SQLite_connectionPath_Tool,
                            '                                               get_NameManager.SQLite_ToolDBMS_Name),
                            '                 msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 操作方式

                            ' 拒付運轉 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Install_Ope
                                excelWriteIn(JobMaker_Form.Spec_install_ope_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_INSTALL_OPE,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 拒付運轉

                            ' 專用運轉 -----------------------------------------------------------------------------------------------------
                            Case usr_Spec_Indep_Ope
                                excelWriteIn(JobMaker_Form.Spec_Indep_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_INDEP_OPE,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 專用運轉

                            ' 戶開行走保護裝置 --------------------------------------------------------------------------------------------------
                            Case usr_Spec_UCMP
                                excelWriteIn(JobMaker_Form.Spec_UCMP_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_UCMP,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '--------------------------------------------------------------------------------------------------- 戶開行走保護裝置

                            ' HIN CPI --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_HIN_CPI
                                excelWriteIn(JobMaker_Form.Spec_HinCpi_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_HIN_CPI,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ HIN CPI

                            ' 火災 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Fire_Ope
                                excelWriteIn(JobMaker_Form.Spec_Fire_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_FIRE_OPE,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                                '避難階
                                msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_ESCAPE_FL,
                                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                                          get_NameManager.SQLite_ToolDBMS_Name)
                                                            ).RefersToRange.Cells.Value = JobMaker_Form.Spec_EscapeFL_TextBox.Text
                                'If JobMaker_Form.Spec_Fire_ComboBox.Text = get_NameManager.TB_O Then
                                '    Dim sig_val As String
                                '    sig_val =
                                '        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_FIRE_OPE_SIGNAL,
                                '                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                '                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                '                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                '                                    ).RefersToRange.Value '取得 火災訊號 內的文字內容
                                '    If JobMaker_Form.Spec_FireSignal_ComboBox.Text = get_NameManager.TB_NO Then
                                '        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_NO,
                                '                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                '                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                '                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                '                                    ).RefersToRange.Characters(InStr(sig_val, nc_val), Len(nc_val)).
                                '                                    Font.Strikethrough = True
                                '    ElseIf JobMaker_Form.Spec_FireSignal_ComboBox.Text = get_NameManager.TB_NC Then
                                '        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_NC,
                                '                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                '                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                '                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                '                                    ).RefersToRange.Characters(InStr(sig_val, no_val), Len(no_val)).
                                '                                    Font.Strikethrough = True
                                '    End If
                                'End If
                                If JobMaker_Form.Spec_Fire_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim spec_fire_ope_signal As String
                                    spec_fire_ope_signal =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_FIRE_OPE_SIGNAL,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    Dim signal_val As String
                                    signal_val =
                                        msExcel_workbook.Names.Item(spec_fire_ope_signal).RefersToRange.Value


                                    If JobMaker_Form.Spec_FireSignal_ComboBox.Text = get_NameManager.TB_NO Then
                                        msExcel_workbook.Names.Item(spec_fire_ope_signal
                                                                    ).RefersToRange.Characters(InStr(signal_val, nc_val), Len(nc_val)).
                                                                    Font.Strikethrough = True
                                    ElseIf JobMaker_Form.Spec_FireSignal_ComboBox.Text = get_NameManager.TB_NC Then
                                        msExcel_workbook.Names.Item(spec_fire_ope_signal
                                                                    ).RefersToRange.Characters(InStr(signal_val, no_val), Len(no_val)).
                                                                    Font.Strikethrough = True
                                    End If


                                    If JobMaker_Form.Spec_Fire_Only_CheckBox.Checked Then
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_FIRE_ONLY,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value =
                                                                    $"(Only {JobMaker_Form.Spec_Fire_Only_TextBox.Text})"
                                    End If
                                End If
                            '------------------------------------------------------------------------------------------------------ 火災

                            ' 消防梯 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Fireman
                                excelWriteIn(JobMaker_Form.Spec_Fireman_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_FIREMAN,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_Fireman_ComboBox.Text = get_NameManager.TB_O And
                                   JobMaker_Form.Spec_Fireman_Only_CheckBox.Checked Then
                                    msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_ESCAPE_FL_ONLY,
                                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                                              get_NameManager.SQLite_ToolDBMS_Name)
                                                                ).RefersToRange.Cells.Value =
                                                                $"(Only {JobMaker_Form.Spec_Fireman_Only_TextBox.Text})"
                                End If

                            '-----------------------------------------------------------------------------------------------------  消防梯

                            ' 停車階運轉 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Parking
                                excelWriteIn(JobMaker_Form.Spec_Parking_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_PARKING,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_Parking_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim spec_pk_cmd1, spec_pk_cmd2 As String
                                    spec_pk_cmd1 =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_PK_CMD1,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_pk_cmd2 =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_PK_CMD2,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)


                                    Dim cmd1, cmd2 As String
                                    Dim elv_val, wtb_val, cob_val, hal_val, dro_val, drc_val As String
                                    cmd1 =
                                        msExcel_workbook.Names.Item(spec_pk_cmd1).RefersToRange.Value '取得 cmd 內的文字內容

                                    cmd2 =
                                        msExcel_workbook.Names.Item(spec_pk_cmd2).RefersToRange.Value '取得 cmd 內的文字內容

                                    elv_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_PK_ELVIC,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 elvic 內的文字內容

                                    wtb_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_PK_WTB,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 WTB 內的文字內容

                                    cob_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_PK_COB,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 COB 內的文字內容

                                    dro_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_PK_DROPEN,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 OPEN 內的文字內容
                                    hal_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_PK_SW,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 SW 內的文字內容
                                    drc_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_PK_DRCLOSE,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 CLOSE 內的文字內容

                                    msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_PARKING_FL,
                                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                ).RefersToRange.Cells.Value = JobMaker_Form.Spec_Parking_FL_TextBox.Text

                                    If JobMaker_Form.Spec_ParkingFL_ELVIC_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(spec_pk_cmd1
                                                                    ).RefersToRange.Characters(InStr(cmd1, elv_val), Len(elv_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    If JobMaker_Form.Spec_ParkingFL_WTB_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(spec_pk_cmd1
                                                                    ).RefersToRange.Characters(InStr(cmd1, wtb_val), Len(wtb_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    If JobMaker_Form.Spec_ParkingFL_COB_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(spec_pk_cmd1
                                                                    ).RefersToRange.Characters(InStr(cmd1, cob_val), Len(cob_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    If JobMaker_Form.Spec_ParkingFL_HALL_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(spec_pk_cmd2
                                                                    ).RefersToRange.Characters(InStr(cmd2, hal_val), Len(hal_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    If JobMaker_Form.Spec_ParkingFL_DR_ComboBox.Text = get_NameManager.TB_DR_OPEN Then
                                        msExcel_workbook.Names.Item(spec_pk_cmd2
                                                                    ).RefersToRange.Characters(InStr(cmd2, drc_val), Len(drc_val)).
                                                                    Font.Strikethrough = True
                                    ElseIf JobMaker_Form.Spec_ParkingFL_DR_ComboBox.Text = get_NameManager.TB_DR_CLOSE Then
                                        msExcel_workbook.Names.Item(spec_pk_cmd2
                                                                    ).RefersToRange.Characters(InStr(cmd2, dro_val), Len(dro_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '停車階運轉 Only
                                    If JobMaker_Form.Spec_Parking_Only_CheckBox.Checked Then
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_PARKING_ONLY,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value =
                                                                    $"(Only {JobMaker_Form.Spec_Parking_Only_TextBox.Text})"
                                    End If
                                End If
                            '------------------------------------------------------------------------------------------------------ 停車階運轉

                            ' 地震管制運轉 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Seismic
                                excelWriteIn(JobMaker_Form.Spec_Seismic_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_SEISMIC,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_Seismic_ComboBox.Text = get_NameManager.TB_O Then
                                    '地震管制Only
                                    If JobMaker_Form.Spec_Seismic_Only_CheckBox.Checked Then
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_Seismic_ONLY,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value =
                                                                    $"(Only {JobMaker_Form.Spec_Seismic_Only_TextBox.Text})"
                                    End If

                                    '地震管制 感知器Only ------------------------------------------
                                    If JobMaker_Form.Spec_SeismicSensor_Only_CheckBox.Checked Then
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_Seismic_ONLY,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value =
                                                                    $"(Only {JobMaker_Form.Spec_SeismicSensor_Only_TextBox.Text})"
                                    End If
                                    msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_Seismic_SENSOR,
                                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                                              get_NameManager.SQLite_ToolDBMS_Name)
                                                                ).RefersToRange.Cells.Value = JobMaker_Form.Spec_SeismicSensor_ComboBox.Text
                                    '------------------------------------------ 地震管制 感知器Only 

                                    '地震管制 自動解除開關Only ------------------------------------
                                    If JobMaker_Form.Spec_SeismicSW_Only_CheckBox.Checked Then

                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_SeismicSW_ONLY,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value =
                                                                    $"(Only {JobMaker_Form.Spec_SeismicSW_Only_TextBox.Text})"
                                    End If

                                    If JobMaker_Form.Spec_SeismicSW_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_SeismicSW_WITH,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_SeismicSW_WITHOUT,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    End If
                                    '------------------------------------ 地震管制 自動解除開關Only 
                                End If
                            '------------------------------------------------------------------------------------------------------ 地震管制運轉

                            ' 車廂管制運轉燈 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_CPI
                                excelWriteIn(JobMaker_Form.Spec_CPI_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_CPI_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim sei_val, fire_val, emerP_val, fm_val, olt_val As String
                                    Dim cpiEmr_val, cpiFm_val, cpiOlt_val As String
                                    cpiEmr_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI_EMER,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 管制 運轉燈內的文字內容
                                    cpiFm_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI_FM,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 緊急 運轉燈內的文字內容
                                    cpiOlt_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI_OLT,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Value '取得 滿載 運轉燈內的文字內容

                                    '車廂管制燈-地震
                                    If JobMaker_Form.Spec_CpiSeismic_ComboBox.Text = get_NameManager.TB_X Then
                                        sei_val =
                                            msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CPI_SEISMIC,
                                                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                                                        ).RefersToRange.Cells.Value '取得地震時的文字內容
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI_EMER,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Characters(InStr(cpiEmr_val, sei_val), Len(sei_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '車廂管制燈-火災
                                    If JobMaker_Form.Spec_CpiFire_ComboBox.Text = get_NameManager.TB_X Then
                                        fire_val =
                                            msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CPI_FIRE,
                                                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                                                        ).RefersToRange.Cells.Value '取得火災時的文字內容
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI_EMER,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Characters(InStr(cpiEmr_val, fire_val), Len(fire_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '車廂管制燈-自家發
                                    If JobMaker_Form.Spec_CpiEmer_ComboBox.Text = get_NameManager.TB_X Then
                                        emerP_val =
                                            msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CPI_EMER,
                                                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                                                        ).RefersToRange.Cells.Value '取得自家發時的文字內容
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI_EMER,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Characters(InStr(cpiEmr_val, emerP_val), Len(emerP_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '車廂管制燈-緊急
                                    If JobMaker_Form.Spec_CpiFM_ComboBox.Text = get_NameManager.TB_X Then
                                        fm_val =
                                            msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CPI_FIREMAN,
                                                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                                                        ).RefersToRange.Cells.Value '取得緊急時的文字內容
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI_FM,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    End If
                                    '車廂管制燈-滿載
                                    If JobMaker_Form.Spec_CpiOLT_ComboBox.Text = get_NameManager.TB_X Then
                                        olt_val =
                                            msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CPI_OLT,
                                                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                                                        ).RefersToRange.Cells.Value '取得超載時的文字內容
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_CPI_OLT,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    End If
                                    '車廂管制燈-滿載Only
                                    If JobMaker_Form.Spec_CpiOLT_Only_CheckBox.Checked Then
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CPI_OLT_ONLY,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value =
                                                                    $"(Only {JobMaker_Form.Spec_CpiOLT_Only_TextBox.Text})"
                                    End If
                                End If
                            '------------------------------------------------------------------------------------------------------ 車廂管制運轉燈

                            ' 車廂上到著鈴 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Car_Gong
                                excelWriteIn(JobMaker_Form.Spec_CarGong_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_GONG,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_CarGong_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim spec_car_gong_pos,
                                        spec_car_gong_cartop, spec_car_gong_cartopbtm,
                                        spec_car_gong_cob, spec_car_gong_vonic As String

                                    spec_car_gong_pos =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_GONG_POS,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_car_gong_cartop =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_GONG_CARTOP,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_car_gong_cartopbtm =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_GONG_CARTOPBTM,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_car_gong_cob =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_GONG_COB,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_car_gong_vonic =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CAR_GONG_VONIC,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    Dim settable_car_top, settable_car_top_btm, settable_car_cob, settable_car_vonic As String
                                    settable_car_top =
                                        get_NameManager.read_DbmsData(get_NameManager.SetTable_CAR_TOP,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    settable_car_top_btm =
                                        get_NameManager.read_DbmsData(get_NameManager.SetTable_CAR_TOP_BTM,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    settable_car_cob =
                                        get_NameManager.read_DbmsData(get_NameManager.SetTable_CAR_COB,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    settable_car_vonic =
                                        get_NameManager.read_DbmsData(get_NameManager.SetTable_CAR_VONIC,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    Dim pos_val As String
                                    Dim carTop_val, carTopBtm_val, cob_val, inVonic_val As String
                                    pos_val =
                                        msExcel_workbook.Names.Item(spec_car_gong_pos).RefersToRange.Cells.Value '取得 位置 的文字內容
                                    carTop_val =
                                        msExcel_workbook.Names.Item(settable_car_top).RefersToRange.Cells.Value '取得 車廂上 的文字內容
                                    carTopBtm_val =
                                        msExcel_workbook.Names.Item(settable_car_top_btm).RefersToRange.Cells.Value '取得 車廂上下 的文字內容
                                    cob_val =
                                        msExcel_workbook.Names.Item(settable_car_cob).RefersToRange.Cells.Value '取得 COB 的文字內容
                                    inVonic_val =
                                        msExcel_workbook.Names.Item(settable_car_vonic).RefersToRange.Cells.Value '取得 VONIC 的文字內容

                                    'Car 車廂上
                                    If JobMaker_Form.Spec_CarGong_Top_CheckBox.Checked = False And
                                       JobMaker_Form.Spec_CarGong_Top_TextBox.Text = get_NameManager.TB_CarTop Then
                                        '無
                                        msExcel_workbook.Names.Item(spec_car_gong_pos
                                                                    ).RefersToRange.Characters(InStr(pos_val, carTop_val), Len(carTop_val)).
                                                                    Font.Strikethrough = True
                                        'msExcel_workbook.Names.Item(spec_car_gong_cartop
                                        '                            ).RefersToRange.Cells.Font.Strikethrough = False
                                    Else
                                        '有
                                        If JobMaker_Form.Spec_CarGong_Top_Only_CheckBox.Checked Then
                                            msExcel_workbook.Names.Item(spec_car_gong_cartop).RefersToRange.Cells.Value =
                                                $"(Only {JobMaker_Form.Spec_CarGong_Top_Only_TextBox.Text})"
                                        End If
                                    End If

                                    'Car 車廂上下
                                    If JobMaker_Form.Spec_CarGong_TopBtm_CheckBox.Checked = False And
                                       JobMaker_Form.Spec_CarGong_TopBtm_TextBox.Text = get_NameManager.TB_CarTopBtm Then
                                        '無
                                        msExcel_workbook.Names.Item(spec_car_gong_pos
                                                                    ).RefersToRange.Characters(InStr(pos_val, carTopBtm_val), Len(carTopBtm_val)).
                                                                    Font.Strikethrough = True
                                    Else
                                        '有
                                        If JobMaker_Form.Spec_CarGong_TopBtm_Only_CheckBox.Checked Then
                                            msExcel_workbook.Names.Item(spec_car_gong_cartopbtm).RefersToRange.Cells.Value =
                                                $"(Only {JobMaker_Form.Spec_CarGong_TopBtm_Only_TextBox.Text})"
                                        End If
                                    End If

                                    'Car 和COB組合
                                    If JobMaker_Form.Spec_CarGong_COB_CheckBox.Checked = False And
                                       JobMaker_Form.Spec_CarGong_COB_TextBox.Text = get_NameManager.TB_WithCOB Then
                                        '無
                                        msExcel_workbook.Names.Item(spec_car_gong_pos
                                                                    ).RefersToRange.Characters(InStr(pos_val, cob_val), Len(cob_val)).
                                                                    Font.Strikethrough = True
                                    Else
                                        '有
                                        If JobMaker_Form.Spec_CarGong_COB_Only_CheckBox.Checked Then
                                            msExcel_workbook.Names.Item(spec_car_gong_cob).RefersToRange.Cells.Value =
                                                $"(Only {JobMaker_Form.Spec_CarGong_COB_Only_TextBox.Text})"
                                        End If
                                    End If

                                    'Car 在Vonic
                                    If JobMaker_Form.Spec_CarGong_VONIC_CheckBox.Checked = False And
                                       JobMaker_Form.Spec_CarGong_VONIC_CheckBox.Text = get_NameManager.TB_InVONIC Then
                                        '無
                                        msExcel_workbook.Names.Item(spec_car_gong_pos
                                                                    ).RefersToRange.Characters(InStr(pos_val, inVonic_val), Len(inVonic_val)).
                                                                    Font.Strikethrough = True
                                        'msExcel_workbook.Names.Item(spec_car_gong_vonic
                                        '                            ).RefersToRange.Cells.Font.Strikethrough = False
                                    Else
                                        '有
                                        If JobMaker_Form.Spec_CarGong_VONIC_Only_CheckBox.Checked Then
                                            msExcel_workbook.Names.Item(spec_car_gong_vonic).RefersToRange.Cells.Value =
                                                $"(Only {JobMaker_Form.Spec_CarGong_VONIC_Only_TextBox.Text})"
                                        End If
                                    End If
                                End If
                            '------------------------------------------------------------------------------------------------------ 開車廂上到著鈴

                            ' 乘場到著鈴 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Hall_Gong
                                excelWriteIn(JobMaker_Form.Spec_HallGong_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_HALL_GONG,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 乘場到著鈴

                            ' 乘場信號文字 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_HPI
                                excelWriteIn(JobMaker_Form.Spec_HPIMsg_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_HPI,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                                If JobMaker_Form.Spec_HPIMsg_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim spec_hpi_msg, spec_hpi_main As String
                                    spec_hpi_msg =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_HPI_MSG,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_hpi_main =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_HPI_MAIN,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    Dim halMsg_val, halMain_val As String
                                    Dim olt_val, main_val, indep_val, fm_val As String
                                    halMsg_val =
                                        msExcel_workbook.Names.Item(spec_hpi_msg).RefersToRange.Cells.Value '取得 乘場燈 的文字內容
                                    halMain_val =
                                        msExcel_workbook.Names.Item(spec_hpi_main).RefersToRange.Cells.Value '取得 保養中 的文字內容
                                    olt_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_HALL_OLT,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 滿載 的文字內容
                                    main_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_HALL_MAIN,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 保養 的文字內容
                                    indep_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_HALL_INDEP,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 專用 的文字內容
                                    fm_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_HALL_FM,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 緊急 的文字內容

                                    '滿載
                                    If JobMaker_Form.Spec_HpiOLT_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(spec_hpi_msg
                                                                    ).RefersToRange.Characters(InStr(halMsg_val, olt_val), Len(olt_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '保養
                                    If JobMaker_Form.Spec_HpiMain_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(spec_hpi_msg
                                                                    ).RefersToRange.Characters(InStr(halMsg_val, main_val), Len(main_val)).
                                                                    Font.Strikethrough = True

                                        msExcel_workbook.Names.Item(spec_hpi_main
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    End If
                                    '專用
                                    If JobMaker_Form.Spec_HpiIndep_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(spec_hpi_msg
                                                                    ).RefersToRange.Characters(InStr(halMsg_val, indep_val), Len(indep_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '緊急
                                    If JobMaker_Form.Spec_HpiFM_ComboBox.Text = get_NameManager.TB_X Then
                                        msExcel_workbook.Names.Item(spec_hpi_msg
                                                                    ).RefersToRange.Characters(InStr(halMsg_val, fm_val), Len(fm_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                End If
                            '------------------------------------------------------------------------------------------------------ 乘場信號文字

                            ' 開門延長 -----------------------------------------------------------------------------------------------------
                            Case usr_Spec_Dr_Hold
                                excelWriteIn(JobMaker_Form.Spec_DrHold_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_DR_HOLD,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 開門延長

                            ' 刷卡機 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_CRD
                                excelWriteIn(JobMaker_Form.Spec_CRD_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_CRD_ComboBox.Text = get_NameManager.TB_O Then


                                    Dim type_all_val, type_notall_val,
                                        crd_Y_val, crd_N_val, rvs_crd_Y, rvs_crd_N_val,
                                        anti_crd_Y_val, anti_crd_N_val, time_crd_val As String

                                    type_all_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_TYPE_ALL,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 全層管制 的文字內容
                                    type_notall_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_TYPE_NOTALL,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 分層管制 的文字內容
                                    crd_Y_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_SPEC_Y,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 式樣有 的文字內容
                                    crd_N_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_SPEC_N,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 式樣無 的文字內容
                                    rvs_crd_Y =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_RVS_CALL_Y,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 逆呼有 的文字內容
                                    rvs_crd_N_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_RVS_CALL_N,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 逆呼無 的文字內容
                                    anti_crd_Y_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_ANTI_Y,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 嬉戲有 的文字內容
                                    anti_crd_N_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_ANTI_N,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 嬉戲無 的文字內容
                                    time_crd_val =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_CRD_TIME_SET,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 時間 的文字內容

                                    Dim spec_crd_type, spec_crd_spec, spec_crd_rvs_call, spec_crd_anti,
                                        spec_crd_rgl4_y, spec_crd_rgl4_n,
                                        spec_crd_rgl5_y, spec_crd_rgl5_n,
                                        spec_crd_auto_n, spec_crd_auto_y As String
                                    spec_crd_type =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_TYPE,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_spec =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_SPEC,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_rvs_call =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_RVS_CALL,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_anti =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_ANTI,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_rgl4_y =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_RGL4_Y,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_rgl4_n =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_RGL4_N,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_rgl5_y =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_RGL5_Y,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_rgl5_n =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_RGL5_N,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_auto_n =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_AUTO_N,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    spec_crd_auto_y =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_CRD_AUTO_Y,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    Dim crd_type, crd, rvs_crd, anti_crd,
                                        auto_crd_y, auto_crd_n As String

                                    crd_type =
                                        msExcel_workbook.Names.Item(spec_crd_type).RefersToRange.Cells.Value '分層或全層
                                    crd =
                                        msExcel_workbook.Names.Item(spec_crd_spec).RefersToRange.Cells.Value '式樣有無
                                    rvs_crd =
                                        msExcel_workbook.Names.Item(spec_crd_rvs_call).RefersToRange.Cells.Value '逆呼有無
                                    anti_crd =
                                        msExcel_workbook.Names.Item(spec_crd_anti).RefersToRange.Cells.Value '防嬉戲有無

                                    auto_crd_y =
                                        msExcel_workbook.Names.Item(spec_crd_auto_y).RefersToRange.Cells.Value '自動登陸 有
                                    auto_crd_n =
                                        msExcel_workbook.Names.Item(spec_crd_auto_n).RefersToRange.Cells.Value '自動登陸 無

                                    '分層/全層管制
                                    If JobMaker_Form.Spec_CRDType_ComboBox.Text = get_NameManager.TB_O Then
                                        msExcel_workbook.Names.Item(spec_crd_type
                                                                    ).RefersToRange.Characters(InStr(crd_type, type_all_val), Len(type_all_val)).
                                                                    Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_crd_type
                                                                    ).RefersToRange.Characters(InStr(crd_type, type_notall_val), Len(type_notall_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '式樣有無
                                    If JobMaker_Form.Spec_CRDSpec_ComboBox.Text = get_NameManager.TB_O Then
                                        msExcel_workbook.Names.Item(spec_crd_spec
                                                                    ).RefersToRange.Characters(InStr(crd, crd_N_val), Len(crd_N_val)).
                                                                    Font.Strikethrough = True
                                        msExcel_workbook.Names.Item(spec_crd_spec
                                                                    ).RefersToRange.Characters(InStr(crd, crd_Y_val), Len(crd_Y_val)).
                                                                    Font.Strikethrough = False
                                    Else
                                        msExcel_workbook.Names.Item(spec_crd_spec
                                                                    ).RefersToRange.Characters(InStr(crd, crd_N_val), Len(crd_N_val)).
                                                                    Font.Strikethrough = False
                                        msExcel_workbook.Names.Item(spec_crd_spec
                                                                    ).RefersToRange.Characters(InStr(crd, crd_Y_val), Len(crd_Y_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '逆向呼叫無效
                                    If JobMaker_Form.Spec_CRDCancell_ComboBox.Text = get_NameManager.TB_O Then
                                        msExcel_workbook.Names.Item(spec_crd_rvs_call
                                                                    ).RefersToRange.Characters(InStr(rvs_crd, rvs_crd_N_val), Len(rvs_crd_N_val)).
                                                                    Font.Strikethrough = True
                                        msExcel_workbook.Names.Item(spec_crd_rvs_call
                                                                    ).RefersToRange.Characters(InStr(rvs_crd, rvs_crd_Y), Len(rvs_crd_Y)).
                                                                    Font.Strikethrough = False
                                    Else
                                        msExcel_workbook.Names.Item(spec_crd_rvs_call
                                                                    ).RefersToRange.Characters(InStr(rvs_crd, rvs_crd_N_val), Len(rvs_crd_N_val)).
                                                                    Font.Strikethrough = False
                                        msExcel_workbook.Names.Item(spec_crd_rvs_call
                                                                    ).RefersToRange.Characters(InStr(rvs_crd, rvs_crd_Y), Len(rvs_crd_Y)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '防止嬉戲呼叫
                                    If JobMaker_Form.Spec_CRDNuisance_ComboBox.Text = get_NameManager.TB_O Then
                                        msExcel_workbook.Names.Item(spec_crd_anti
                                                                    ).RefersToRange.Characters(InStr(anti_crd, anti_crd_N_val), Len(anti_crd_N_val)).
                                                                    Font.Strikethrough = True
                                        msExcel_workbook.Names.Item(spec_crd_anti
                                                                    ).RefersToRange.Characters(InStr(anti_crd, anti_crd_Y_val), Len(anti_crd_Y_val)).
                                                                    Font.Strikethrough = False
                                    Else
                                        msExcel_workbook.Names.Item(spec_crd_anti
                                                                    ).RefersToRange.Characters(InStr(anti_crd, anti_crd_N_val), Len(anti_crd_N_val)).
                                                                    Font.Strikethrough = False
                                        msExcel_workbook.Names.Item(spec_crd_anti
                                                                    ).RefersToRange.Characters(InStr(anti_crd, anti_crd_Y_val), Len(anti_crd_Y_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    '自動登陸
                                    If JobMaker_Form.Spec_CRDReg_ComboBox.Text = get_NameManager.TB_O Then
                                        msExcel_workbook.Names.Item(spec_crd_auto_n
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_crd_auto_y
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    End If
                                    'if79 id=4 / 5
                                    If JobMaker_Form.Spec_CRDID4_ComboBox.Text = get_NameManager.TB_O Then
                                        msExcel_workbook.Names.Item(spec_crd_rgl4_n
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_crd_rgl4_y
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    End If
                                    If JobMaker_Form.Spec_CRDID5_ComboBox.Text = get_NameManager.TB_O Then
                                        msExcel_workbook.Names.Item(spec_crd_rgl5_n
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_crd_rgl5_y
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    End If
                                End If
                            '------------------------------------------------------------------------------------------------------ 刷卡機

                            ' 自家發電 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Emer_Power
                                excelWriteIn(JobMaker_Form.Spec_Emer_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_POWER,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_Emer_ComboBox.Text = get_NameManager.TB_O Then

                                    '自家發Signal --------------------------------------------------------------------------------
                                    Dim spec_emer_signal As String
                                    spec_emer_signal =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_SIGNAL,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)
                                    Dim sig_val As String
                                    sig_val =
                                        msExcel_workbook.Names.Item(spec_emer_signal).RefersToRange.Value '取得 自家發訊號 內的文字內容

                                    If JobMaker_Form.Spec_EmerSignal_ComboBox.Text = get_NameManager.TB_NO Then
                                        msExcel_workbook.Names.Item(spec_emer_signal
                                                                    ).RefersToRange.Characters(InStr(sig_val, nc_val), Len(nc_val)
                                                                    ).Font.Strikethrough = True
                                    ElseIf JobMaker_Form.Spec_EmerSignal_ComboBox.Text = get_NameManager.TB_NC Then
                                        msExcel_workbook.Names.Item(spec_emer_signal
                                                                    ).RefersToRange.Characters(InStr(sig_val, no_val), Len(no_val)
                                                                    ).Font.Strikethrough = True
                                    End If
                                    '-------------------------------------------------------------------------------- 自家發Signal 

                                    '自家發容量 ----------------------------------------------------------------------------------
                                    excelWriteIn(JobMaker_Form.Spec_EmerCapacity_TextBox.Text,
                                                 get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_CAPCITY,
                                                                               get_NameManager.SQLite_tableName_NameManager_TW,
                                                                               get_NameManager.SQLite_connectionPath_Tool,
                                                                               get_NameManager.SQLite_ToolDBMS_Name),
                                                 msExcel_workbook)
                                    '---------------------------------------------------------------------------------- 自家發容量 

                                    '自家發入力點 -------------------------------------------------------------------------------
                                    excelWriteIn(JobMaker_Form.Spec_EmerInput_ComboBox.Text,
                                                 get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_INPUT,
                                                                               get_NameManager.SQLite_tableName_NameManager_TW,
                                                                               get_NameManager.SQLite_connectionPath_Tool,
                                                                               get_NameManager.SQLite_ToolDBMS_Name),
                                                 msExcel_workbook)
                                    '------------------------------------------------------------------------------- 自家發入力點 

                                    '自家發Address -----------------------------------------------------------------------------
                                    excelWriteIn(JobMaker_Form.Spec_EmerAddress_ComboBox.Text,
                                                 get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_ADDRESS,
                                                                               get_NameManager.SQLite_tableName_NameManager_TW,
                                                                               get_NameManager.SQLite_connectionPath_Tool,
                                                                               get_NameManager.SQLite_ToolDBMS_Name),
                                                 msExcel_workbook)
                                    '----------------------------------------------------------------------------- 自家發Address 

                                    '自家發Group -----------------------------------------------------------------------------

                                    Dim JM_Spec_Emer As String() =
                                               {get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_POWER_GROUP,
                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                                get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_POWER_CarName,
                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                                get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_POWER_EscapeFL,
                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                                get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_POWER_RETURN,
                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                                get_NameManager.read_DbmsData(get_NameManager.SPEC_EMER_POWER_CONTINUE,
                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                              get_NameManager.SQLite_ToolDBMS_Name)
                                               }
                                    Dim dyCtrlName As DynamicControlName = New DynamicControlName
                                    dyCtrlName.JobMaker_EmerInfo()

                                    Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData

                                    dynamicControl_writeInExcel(JobMaker_Form.Spec_EmerNum_NumericUpDown.Value,
                                                                get_NameManager.SPEC_EMER_POWER_GROUP,
                                                                JM_Spec_Emer,
                                                                JobMaker_Form.Spec_emerGroup_TabControl,
                                                                spec_stored.LoadStored_PanelType.DoubleLayer_Panel,
                                                                dyCtrlName.JobMaker_EmerTBInfoName_Array.Count,
                                                                dyCtrlName.JobMaker_EmerTBInfoName_Array,
                                                                msExcel_workbook)
                                    '----------------------------------------------------------------------------- 自家發Group 
                                End If
                            '-------------------------------------------------------------------------------------------------------- 自家發電

                            ' Landic -----------------------------------------------------------------------------------------------------
                            Case usr_Spec_Landic
                                excelWriteIn(JobMaker_Form.Spec_Landic_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_LANDIC,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ Landic

                            ' 基準階賦歸 -----------------------------------------------------------------------------------------------------
                            Case usr_Spec_MLF_Return
                                excelWriteIn(JobMaker_Form.Spec_MFLReturn_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_MLF_RETURN,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_MFLReturn_ComboBox.Text = get_NameManager.TB_O Then
                                    msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_MAIN_FL,
                                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                                              get_NameManager.SQLite_ToolDBMS_Name)
                                                                ).RefersToRange.Cells.Value = JobMaker_Form.Spec_MFLReturn_FL_TextBox.Text & "階"
                                End If
                            '------------------------------------------------------------------------------------------------------ 基準階賦歸

                            ' VONIC --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Vonic
                                excelWriteIn(JobMaker_Form.Spec_Vonic_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_Vonic_ComboBox.Text = get_NameManager.TB_O Then
                                    If JobMaker_Form.Spec_Vonic_standard_ComboBox.Text = get_NameManager.TB_O Then
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC_NSTD_C,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC_NSTD_E,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC_STD_C,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = False
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC_STD_E,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = False
                                    Else
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC_NSTD_C,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = False
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC_NSTD_E,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = False
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC_STD_C,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_VONIC_STD_E,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Font.Strikethrough = True
                                    End If
                                End If

                            '------------------------------------------------------------------------------------------------------ VONIC

                            ' 殘障HB --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_WHB
                                excelWriteIn(JobMaker_Form.Spec_WCOB_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_WCOB,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)


                                If JobMaker_Form.Spec_WCOB_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim spec_sub_wcob, spec_whb_bz, spec_whb_ring As String

                                    spec_sub_wcob =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_WCOB_SUB,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    spec_whb_bz =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_WCOB_BZ,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    spec_whb_ring =
                                        get_NameManager.read_DbmsData(get_NameManager.SPEC_WCOB_RING,
                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                      get_NameManager.SQLite_connectionPath_Tool,
                                                                      get_NameManager.SQLite_ToolDBMS_Name)

                                    Dim bz_Y, bz_N, ring_Y, ring_N As String
                                    Dim sub_wcob_val, bz_val, ring_val As String

                                    bz_Y =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_WCOB_BZ_Y,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 BZ有 的文字內容
                                    bz_N =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_WCOB_BZ_N,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 BZ無 的文字內容
                                    ring_Y =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_WCOB_RING_Y,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 RING有 的文字內容
                                    ring_N =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_WCOB_RING_N,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Cells.Value '取得 RING無 的文字內容
                                    sub_wcob_val =
                                        msExcel_workbook.Names.Item(spec_sub_wcob).RefersToRange.Cells.Value
                                    bz_val =
                                        msExcel_workbook.Names.Item(spec_whb_bz).RefersToRange.Cells.Value
                                    ring_val =
                                        msExcel_workbook.Names.Item(spec_whb_ring).RefersToRange.Cells.Value

                                    If JobMaker_Form.Spec_WCOB_only_CheckBox.Checked Then 'COB Only
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_WCOB_ONLY,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                               ).RefersToRange.Cells.Value =
                                                               $"(Only {JobMaker_Form.Spec_WCOB_only_TextBox.Text})"
                                    End If
                                    If JobMaker_Form.Spec_WSCOB_only_CheckBox.Checked Then 'SCOB Only
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_WSCOB_ONLY,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                               ).RefersToRange.Cells.Value =
                                                               $"(Only {JobMaker_Form.Spec_WSCOB_only_TextBox.Text})"
                                    End If

                                    If JobMaker_Form.Spec_WSCOB_ComboBox.Text = get_NameManager.TB_O Then 'SCOB
                                        msExcel_workbook.Names.Item(spec_sub_wcob
                                                                    ).RefersToRange.Characters(InStr(sub_wcob_val, without_val), Len(without_val)).
                                                                    Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_sub_wcob
                                                                    ).RefersToRange.Characters(InStr(sub_wcob_val, with_val), Len(with_val)).
                                                                    Font.Strikethrough = True
                                    End If
                                    If JobMaker_Form.Spec_WCOB_Ring_ComboBox.Text = get_NameManager.TB_O Then '鳴動
                                        msExcel_workbook.Names.Item(spec_whb_bz
                                                                    ).RefersToRange.Characters(InStr(bz_val, bz_N), Len(bz_N)).
                                                                    Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_whb_bz
                                                                    ).RefersToRange.Characters(InStr(bz_val, bz_Y), Len(bz_Y)).
                                                                    Font.Strikethrough = True
                                    End If
                                    If JobMaker_Form.Spec_WCOB_Ring_ComboBox.Text = get_NameManager.TB_O Then 'Ring
                                        msExcel_workbook.Names.Item(spec_whb_ring
                                                                    ).RefersToRange.Characters(InStr(ring_val, ring_N), Len(ring_N)).
                                                                    Font.Strikethrough = True
                                    Else
                                        msExcel_workbook.Names.Item(spec_whb_ring
                                                                    ).RefersToRange.Characters(InStr(ring_val, ring_Y), Len(ring_Y)).
                                                                    Font.Strikethrough = True
                                    End If
                                End If
                            '------------------------------------------------------------------------------------------------------ 殘障HB

                            ' ELVIC --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_Elvic
                                excelWriteIn(JobMaker_Form.Spec_Elvic_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_ELVIC,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_Elvic_ComboBox.Text = get_NameManager.TB_O Then
                                    Dim elv_ele_grp As CheckBox() =
                                                                    {JobMaker_Form.Spec_Elvic_Parking_CheckBox, JobMaker_Form.Spec_Elvic_VIP_CheckBox,
                                                                     JobMaker_Form.Spec_Elvic_Indep_CheckBox, JobMaker_Form.Spec_Elvic_FloorLockOut_CheckBox,
                                                                     JobMaker_Form.Spec_Elvic_Express_CheckBox, JobMaker_Form.Spec_Elvic_ReturnFL_CheckBox
                                                                    }
                                    Dim elv_grp_grp As CheckBox() =
                                                                    {JobMaker_Form.Spec_Elvic_Traffic_Peak_CheckBox, JobMaker_Form.Spec_Elvic_MainFL_CheckBox,
                                                                     JobMaker_Form.Spec_Elvic_FloorLockOut_CheckBox, JobMaker_Form.Spec_Elvic_Zoning_CheckBox,
                                                                     JobMaker_Form.Spec_Elvic_CarCall_CheckBox
                                                                    }
                                    Dim elv_other_grp As CheckBox() =
                                                                      {JobMaker_Form.Spec_Elvic_Fire_CheckBox, JobMaker_Form.Spec_Elvic_Wavic_CheckBox,
                                                                       JobMaker_Form.Spec_Elvic_CRD_CheckBox
                                                                      }
                                    Dim num_grp As String() = {"①", "②", "③", "④", "⑤", "⑥"}
                                    Dim sh_name As String =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_ELVIC,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Worksheet.Name
                                    Dim elv_Row As Integer =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_ELVIC_CMD,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Row '號機名是第n行
                                    Dim elv_Col As Integer =
                                        msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_ELVIC_CMD,
                                                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                    ).RefersToRange.Column '號機名是第n列
                                    Dim first_i As Integer = 0
                                    '第一大象
                                    For i = 1 To elv_ele_grp.Count
                                        If elv_ele_grp(i - 1).Checked Then
                                            msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SPEC_ELVIC_CMD,
                                                                                                      get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                                                        ).RefersToRange.Cells.Value =
                                                                        "1." & num_grp(i - 1) & elv_ele_grp(i - 1).Text
                                            elv_Row = elv_Row + 1
                                            msExcel_workbook.Worksheets(sh_name).Range($"{elv_Row}:{elv_Row}").Insert

                                            first_i = i
                                            Exit For
                                        End If
                                    Next
                                    For ii = 1 To elv_ele_grp.Count
                                        If ii <> first_i Then
                                            If elv_ele_grp(ii - 1).Checked Then
                                                msExcel_workbook.Worksheets(sh_name).Cells(elv_Row, elv_Col).Value =
                                                    "   " & num_grp(ii - 1) & elv_ele_grp(ii - 1).Text
                                                elv_Row = elv_Row + 1
                                                msExcel_workbook.Worksheets(sh_name).Range($"{elv_Row }:{elv_Row}").Insert
                                            End If
                                        End If
                                    Next
                                    '第二大象
                                    For i_2 = 1 To elv_grp_grp.Count
                                        If elv_grp_grp(i_2 - 1).Checked Then
                                            msExcel_workbook.Worksheets(sh_name).Cells(elv_Row, elv_Col).Value =
                                                "2." & num_grp(i_2 - 1) & elv_grp_grp(i_2 - 1).Text
                                            elv_Row = elv_Row + 1
                                            msExcel_workbook.Worksheets(sh_name).Range($"{elv_Row}:{elv_Row}").Insert

                                            first_i = i_2

                                            If elv_grp_grp(i_2 - 1).Name = JobMaker_Form.Spec_Elvic_Traffic_Peak_CheckBox.Name Then
                                                msExcel_workbook.Worksheets(sh_name).Cells(elv_Row, elv_Col).Value =
                                                    "   " & JobMaker_Form.Spec_Elvic_Traffic_Peak_ComboBox.Text
                                                elv_Row = elv_Row + 1
                                                msExcel_workbook.Worksheets(sh_name).Range($"{elv_Row}:{elv_Row}").Insert
                                            End If
                                            Exit For
                                        End If

                                    Next
                                    For ii_2 = 1 To elv_grp_grp.Count
                                        If ii_2 <> first_i Then
                                            If elv_grp_grp(ii_2 - 1).Checked Then
                                                msExcel_workbook.Worksheets(sh_name).Cells(elv_Row, elv_Col).Value =
                                                    "   " & num_grp(ii_2 - 1) & elv_grp_grp(ii_2 - 1).Text
                                                elv_Row = elv_Row + 1
                                                msExcel_workbook.Worksheets(sh_name).Range($"{elv_Row }:{elv_Row}").Insert
                                            End If
                                        End If
                                    Next
                                    '其他大象
                                    For i_3 = 1 To elv_other_grp.Count
                                        If elv_other_grp(i_3 - 1).Checked Then
                                            msExcel_workbook.Worksheets(sh_name).Cells(elv_Row, elv_Col).Value =
                                                "・" & elv_other_grp(i_3 - 1).Text
                                            elv_Row = elv_Row + 1
                                            msExcel_workbook.Worksheets(sh_name).Range($"{elv_Row }:{elv_Row}").Insert
                                        End If
                                    Next
                                End If
                            '------------------------------------------------------------------------------------------------------ ELVIC

                            ' 乘場廳燈 -----------------------------------------------------------------------------------------------------
                            Case usr_Spec_HLL
                                excelWriteIn(JobMaker_Form.Spec_HLL_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_HLL,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 乘場廳燈

                            ' 運轉手盤 --------------------------------------------------------------------------------------------------------
                            Case usr_Spec_ATT
                                excelWriteIn(JobMaker_Form.Spec_ATT_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_ATT,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 運轉手盤 

                            ' 浸水管制運轉 -----------------------------------------------------------------------------------------------------
                            Case usr_Spec_Flood
                                excelWriteIn(JobMaker_Form.Spec_Flood_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_FLOOD,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_FLOOD_FL,
                                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                                          get_NameManager.SQLite_ToolDBMS_Name)
                                                            ).RefersToRange.Cells.Value = JobMaker_Form.Spec_Flood_FL_TextBox.Text
                            '------------------------------------------------------------------------------------------------------ 浸水管制運轉

                            ' LS1M ------------------------------------------------------------------------------------------------------
                            Case usr_Spec_LS1M
                                excelWriteIn(JobMaker_Form.Spec_LS1M_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_LS1M,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ LS1M

                            ' 電力回升 ------------------------------------------------------------------------------------------------------
                            Case usr_Spec_PRU
                                excelWriteIn(JobMaker_Form.Spec_PRU_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_PRU,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 電力回升

                            ' 正背門 -----------------------------------------------------------------------------------------------------
                            Case usr_Spec_FrontRear_DR
                                excelWriteIn(JobMaker_Form.Spec_FrontRearDr_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_FRONT_REAR_DR,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 正背門

                            ' 單群控切換 -------------------------------------------------------------------------------------------------
                            Case usr_Spec_OpeSw
                                excelWriteIn(JobMaker_Form.Spec_OpeSw_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_OPE_SW,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)

                                If JobMaker_Form.Spec_OpeSw_ComboBox.Text = get_NameManager.TB_O Then
                                    excelWriteIn(JobMaker_Form.Spec_OpeSw_DevicePos_TextBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SetTable_OpeSW_Content,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                                    excelWriteIn(JobMaker_Form.Spec_OpeSw_InputPos_ComboBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_OPE_SW_POS,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                                    excelWriteIn(JobMaker_Form.Spec_OpeSw_InputAddress_TextBox.Text,
                                             get_NameManager.read_DbmsData(get_NameManager.SPEC_OPE_SW_ADDRESS,
                                                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                                                           get_NameManager.SQLite_connectionPath_Tool,
                                                                           get_NameManager.SQLite_ToolDBMS_Name),
                                             msExcel_workbook)
                                End If
                                '---------------------------------------------------------------------------------------------- 單群控切換

                        End Select
                    Catch ex As Exception
                        JobMaker_Form.ResultFailOutput_TextBox.Text +=
                            ($"<{JobMaker_Form.JMFileCho_Spec_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_TWSpec_str}>{vbCrLf}")
                    End Try
                End If
            Next



        Else
            JobMaker_Form.ResultFailOutput_TextBox.Text += "<提醒> 仕樣 分頁未輸出，原因:分頁未打勾"
            JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.Spec_TabPage
            Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.Spec_TabPage.Text}仕樣分頁」未勾選是否重來?"), vbYesNo)
            If basic_result = DialogResult.Yes And msExcel_workbook IsNot Nothing Then
                returnError_isPageRestart = True
                'msExcel_workbook.Close()
                'msExcel_App.Quit()
            End If
            'MsgBox("仕樣分頁左上角的CheckBox沒有勾選", MsgBoxStyle.Exclamation, "Fail Message")
        End If

    End Sub

    ''' <summary>
    ''' Job Maker >> 重要設定 (快速摺疊Code:CRTL+M+M)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="msExcel_app"></param>
    Public Sub Spec_Important(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application)
        '重要設定
        If JobMaker_Form.Use_Imp_CheckBox.Checked Then
            Dim usr_IMP_MachineType, usr_IMP_FAN, usr_IMP_OverBalance, usr_IMP_WHB, usr_IMP_DoorType As String
            Dim usr_IMP_HIN As String

            usr_IMP_MachineType =
                JobMaker_Form.Imp_MachineRoom_ComboBox.Name
            'usr_IMP_FAN =
            '    JobMaker_Form.Imp_FAN_ComboBox.Name
            usr_IMP_OverBalance =
                JobMaker_Form.Imp_OverBalance_ComboBox.Name
            usr_IMP_WHB =
                JobMaker_Form.Imp_WHB_ComboBox.Name
            usr_IMP_DoorType =
                JobMaker_Form.Imp_DoorType_TextBox.Name
            usr_IMP_HIN =
                JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls.Count


            Dim usrInput_IMP_arr() As String = {usr_IMP_MachineType, usr_IMP_OverBalance,
                                                usr_IMP_WHB, usr_IMP_DoorType, usr_IMP_HIN}
            Dim i_ImpStr As String

            '輸入相對應的check list值
            For Each i_ImpStr In usrInput_IMP_arr
                If i_ImpStr <> "" Then
                    Try
                        Select Case i_ImpStr
                            'Case usr_IMP_MachineType

                            'Case usr_IMP_FAN
                                'excelWriteIn(JobMaker_Form.Imp_FAN_ComboBox.Text,
                                '             get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_FAN,
                                '                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                '                                           get_NameManager.SQLite_connectionPath_Tool,
                                '                                           get_NameManager.SQLite_ToolDBMS_Name),
                                '            msExcel_workbook)
                                'If JobMaker_Form.Imp_FAN_ComboBox.Text <> "WITH(ION)" Then
                                '    msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_FAN_CONTENT,
                                '                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                '                                           get_NameManager.SQLite_connectionPath_Tool,
                                '                                           get_NameManager.SQLite_ToolDBMS_Name)
                                '                                ).RefersToRange.Cells.Font.Strikethrough = True
                                'End If
                            Case usr_IMP_OverBalance
                                excelWriteIn(JobMaker_Form.Imp_OverBalance_ComboBox.Text,
                                            get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_BALANCE,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name),
                                            msExcel_workbook)
                            Case usr_IMP_WHB
                                excelWriteIn(JobMaker_Form.Imp_WHB_ComboBox.Text,
                                            get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_WCOB,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name),
                                            msExcel_workbook)
                            Case usr_IMP_DoorType
                                excelWriteIn(JobMaker_Form.Imp_DoorType_TextBox.Text,
                                            get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_DOOR,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name),
                                            msExcel_workbook)
                            Case usr_IMP_HIN
                                Imp_HIN_Write(msExcel_workbook, msExcel_app)
                        End Select
                    Catch ex As Exception
                        JobMaker_Form.ResultFailOutput_TextBox.Text +=
                            ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_ImpStr}>{vbCrLf}")
                    End Try
                Else
                    JobMaker_Form.ResultFailOutput_TextBox.Text +=
                            ($"<提醒> {i_ImpStr} 為空值寫入失敗")
                End If
            Next

        Else
            JobMaker_Form.ResultFailOutput_TextBox.Text += "<提醒> 重要設定 分頁未輸出，原因:分頁未打勾"
            JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.Important_TabPage
            Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.Important_TabPage.Text}仕樣分頁」未勾選是否重來?"), vbYesNo)
            If basic_result = DialogResult.Yes And msExcel_workbook IsNot Nothing Then
                returnError_isPageRestart = True
                'msExcel_workbook.Close()
                'msExcel_app.Quit()
            End If
        End If
    End Sub

    ''' <summary>
    ''' [重要設定 > HIN > 將Hall Indicator內的值寫入Excel中]
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    Private Sub Imp_HIN_Write(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application)
        'Dim HinLiftDiff_bool, HinFLDiff_bool As Boolean
        'Dim lift_i, stop_i As Integer
        If JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls.Count <> 0 Then
            Dim Imp_HIN_FL_Col, Imp_HIN_FL_Row,
            Imp_HIN_Col, Imp_HIN_Row As Integer
            Imp_HIN_FL_Col =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook,
                                                            get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN_FL,
                                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                                          get_NameManager.SQLite_ToolDBMS_Name))
            'msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN_FL,
            '                                                              get_NameManager.SQLite_tableName_NameManager_TW,
            '                                                              get_NameManager.SQLite_connectionPath_Tool,
            '                                                              get_NameManager.SQLite_ToolDBMS_Name)).RefersToRange.Column '行
            Imp_HIN_FL_Row =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook,
                                                            get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN_FL,
                                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                                          get_NameManager.SQLite_ToolDBMS_Name))
            'msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN_FL,
            '                                                              get_NameManager.SQLite_tableName_NameManager_TW,
            '                                                              get_NameManager.SQLite_connectionPath_Tool,
            '                                                              get_NameManager.SQLite_ToolDBMS_Name)).RefersToRange.Row '列
            Imp_HIN_Col =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook,
                                                            get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN,
                                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                                          get_NameManager.SQLite_ToolDBMS_Name))
            'msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN,
            '                                                              get_NameManager.SQLite_tableName_NameManager_TW,
            '                                                              get_NameManager.SQLite_connectionPath_Tool,
            '                                                              get_NameManager.SQLite_ToolDBMS_Name)).RefersToRange.Column '行
            Imp_HIN_Row =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook,
                                                            get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN,
                                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                                          get_NameManager.SQLite_ToolDBMS_Name))
            'msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN,
            '                                                              get_NameManager.SQLite_tableName_NameManager_TW,
            '                                                              get_NameManager.SQLite_connectionPath_Tool,
            '                                                              get_NameManager.SQLite_ToolDBMS_Name)).RefersToRange.Row '列
            Dim Imp_SheetName As String

            Imp_SheetName =
                getMathOnExcel.getWorksheetName_fromNameManager(msExcel_workbook,
                                                           get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN_FL,
                                                                                         get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                         get_NameManager.SQLite_connectionPath_Tool,
                                                                                         get_NameManager.SQLite_ToolDBMS_Name))
            'msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.IMPORTANT_HIN_FL,
            '                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
            '                                                                          get_NameManager.SQLite_connectionPath_Tool,
            '                                                                          get_NameManager.SQLite_ToolDBMS_Name)
            '                                                                          ).RefersToRange.Worksheet.Name

            Dim HinLiftDiff_bool, HinFLDiff_bool As Boolean
            Dim lift_i, stop_i As Integer
            Dim HinRowNum_InExcel As Integer '目前在Excel中特定欄位後第N行

            '求最高樓層 ----------------------------------------------
            Dim stopFL_MAX, stopFL_MIN As Integer 'HIN中最高樓層
            For lift_i = 1 To JobMaker_Form.LiftNum
                For stop_i = 1 To JobMaker_Form.arr_liftStopFL(lift_i - 1)
                    If stop_i > stopFL_MAX Then
                        stopFL_MAX = stop_i
                    Else
                        stopFL_MIN = stop_i
                    End If
                Next
            Next

            'Console.WriteLine($"所有電梯中最大樓層數為:{stopFL_MAX} / 最小樓層數為:{stopFL_MIN}")
            Dim arr_liftStopFL_userContent(JobMaker_Form.LiftNum - 1, stopFL_MAX - 1) As String
            'ResultOutput_TextBox.Text += $"最高樓層數:{stopFL_MAX} 目前陣列數 {arr_liftStopFL_userContent.Length} {vbCrLf}"
            '---------------------------------------------- 求最高樓層 

            Dim dyCtrlName As DynamicControlName = New DynamicControlName


            '儲存使用者值得內容 ----------------------------------------------------------------
            For Each flp In JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls.OfType(Of FlowLayoutPanel)
                For Each cb In flp.Controls.OfType(Of CheckBox)
                    For lift_i = 1 To JobMaker_Form.LiftNum
                        For stop_i = 1 To JobMaker_Form.arr_liftStopFL(lift_i - 1)
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
                For i = 1 To JobMaker_Form.arr_liftStopFl_StdContent.Count
                    For lift_i = 1 To JobMaker_Form.LiftNum
                        If JobMaker_Form.arr_liftStopFl_EachContent(i - 1, lift_i) <> Nothing Then '共三列，第一列為標準值
                            JobMaker_Form.arr_liftStopFl_EachContent(i - 1, lift_i) = Nothing '將值都清空做後續比對
                        End If
                    Next
                Next
                '---- 每次換樓層時清空arr_liftStopFl_EachContent內資料 


                '每次換樓層時判斷 #1~#N 號機該樓層HIN是否都相同? ---------------------------------
                For lift_i = 1 To JobMaker_Form.LiftNum
                    If lift_i < JobMaker_Form.LiftNum Then
                        'Console.WriteLine($"{stop_i}FL:{arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)},{arr_liftStopFL_userContent(lift_i, stop_i - 1)}")
                        If arr_liftStopFL_userContent(lift_i - 1, stop_i - 1) =
                            arr_liftStopFL_userContent(lift_i, stop_i - 1) Then
                            '號機之間值相同 -------------------
                            HinLiftDiff_bool = False
                            '------------------- 號機之間值相同

                            '上下樓層之間不同 ------------
                            For lift_ii = 1 To JobMaker_Form.LiftNum
                                If stop_i + 1 < stopFL_MAX Then
                                    If arr_liftStopFL_userContent(lift_ii - 1, stop_i) <>
                                        arr_liftStopFL_userContent(lift_ii - 1, stop_i - 1) Then
                                        HinFLDiff_bool = True
                                        HinPoint_bool = False
                                    End If
                                End If
                            Next
                            '------------ 上下樓層之間不同 
                        Else
                            '號機之間值不相同 -----------------
                            HinLiftDiff_bool = True
                            HinLiftDiff_cnt = HinLiftSame_cnt + 1
                            '----------------- 號機之間值不相同
                            Exit For
                        End If
                    End If
                Next
                lift_i = 0

                If HinLiftDiff_bool Then '表示同樓層的號機之間值都不相同


                    For lift_i = 1 To JobMaker_Form.LiftNum
                        '當使用者輸入的HIN為空時 ----------------------------------------------
                        If arr_liftStopFL_userContent(lift_i - 1, stop_i - 1) = "" Then
                            'ResultOutput_TextBox.Text += $"號機#{lift_i} 第{stop_i}樓不相同 : #{lift_i}:None {vbCrLf}"
                        End If
                        '---------------------------------------------- 當使用者輸入的HIN為空時 

                        '如果使用者輸入與標準值相同時就先儲存在EachContent陣列中 ----------------------------------------------
                        For i = 1 To JobMaker_Form.arr_liftStopFl_StdContent.Count
                            If arr_liftStopFL_userContent(lift_i - 1, stop_i - 1) = JobMaker_Form.arr_liftStopFl_StdContent(i - 1) Then
                                JobMaker_Form.arr_liftStopFl_EachContent(i - 1, lift_i) = arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)
                                'ResultOutput_TextBox.Text += $"號機#{lift_i} 第{stop_i}樓不相同 : #{lift_i}:{arr_liftStopFl_EachContent(i - 1, lift_i)} {vbCrLf}"
                            End If
                        Next
                        '---------------------------------------------- 如果使用者輸入與標準值相同時就先儲存在EachContent陣列中 
                    Next
                    lift_i = 0

                    '輸出以下值 e.g #1,2:without/#3:with 字樣 -------------------------------------------------
                    Dim temp_OnlyString As String
                    temp_OnlyString = ""

                    '當同樓層不同時剛好為最後一號機時
                    If HinLiftDiff_cnt = stopFL_MAX Then
                        msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_FL_Row + HinRowNum_InExcel, Imp_HIN_FL_Col).value =
                            $"Hall Indicator {stop_i - 1} FL : {arr_liftStopFL_userContent(lift_i, stop_i - 1)}{vbCrLf}"
                        HinRowNum_InExcel += 2
                        'JobMaker_Form.ResultOutput_TextBox.Text += $"Hall Indicator {stop_i - 1} FL : {arr_liftStopFL_userContent(lift_i, stop_i - 1)}{vbCrLf}"
                    End If

                    JobMaker_Form.ResultOutput_TextBox.Text += $"Hall Indicator {stop_i} FL : Only 號機  "
                    temp_OnlyString += $"Only 號機"
                    Dim EachContent_Bool As Boolean
                    For i = 1 To JobMaker_Form.arr_liftStopFl_StdContent.Count
                        EachContent_Bool = False
                        For lift_i = 1 To JobMaker_Form.LiftNum
                            If JobMaker_Form.arr_liftStopFl_EachContent(i - 1, lift_i) <> "" Then
                                JobMaker_Form.ResultOutput_TextBox.Text += $"#{lift_i},"
                                temp_OnlyString += $"#{lift_i},"
                                EachContent_Bool = True
                            End If
                        Next
                        If EachContent_Bool And JobMaker_Form.arr_liftStopFl_EachContent(i - 1, 0) <> "" Then
                            JobMaker_Form.ResultOutput_TextBox.Text += $":{JobMaker_Form.arr_liftStopFl_EachContent(i - 1, 0)}/"
                            temp_OnlyString += $":{JobMaker_Form.arr_liftStopFl_EachContent(i - 1, 0)}/"
                        End If
                    Next

                    msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_FL_Row + HinRowNum_InExcel, Imp_HIN_FL_Col).Value =
                        $"Hall Indicator {stop_i} FL"
                    msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_Row + HinRowNum_InExcel, Imp_HIN_Col).Value =
                        temp_OnlyString

                    HinRowNum_InExcel += 2

                    If stop_i = stopFL_MAX Then
                        topFL_End_bool = True
                    Else
                        topFL_End_bool = False
                    End If
                    JobMaker_Form.ResultOutput_TextBox.Text += $"{vbCrLf}"
                    '------------------------------------------------- 輸出以下值 e.g #1,2:without/#3:with 字樣 

                ElseIf HinLiftDiff_bool = False Then '表示同樓層號機之間值都相同

                    lift_i = 1
                    HinLiftSame_cnt += 1
                    If HinLiftSame_cnt = 1 Then
                        If stop_i = 1 Then '最底樓層
                            JobMaker_Form.ResultOutput_TextBox.Text +=
                                $"Hall Indicator BOTTOM FL : {arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}"
                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_FL_Row + HinRowNum_InExcel, Imp_HIN_FL_Col).Value =
                                    $"Hall Indicator BOTTOM FL"
                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_Row + HinRowNum_InExcel, Imp_HIN_Col).Value =
                                    arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)

                            HinRowNum_InExcel += 2
                        Else '當其他樓層從HinLiftSame_cnt = 1開始
                            JobMaker_Form.ResultOutput_TextBox.Text +=
                                $"Hall Indicator {stop_i} FL : {arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}"

                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_FL_Row + HinRowNum_InExcel, Imp_HIN_FL_Col).Value =
                                    $"Hall Indicator {stop_i} FL"
                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_Row + HinRowNum_InExcel, Imp_HIN_Col).Value =
                                    arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)

                            HinRowNum_InExcel += 2
                        End If
                    ElseIf HinLiftSame_cnt = 2 Then
                        If HinFLDiff_bool Then
                            'HinLiftSame_cnt = 0
                            JobMaker_Form.ResultOutput_TextBox.Text +=
                                $"Hall Indicator {stop_i} FL : {arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}"

                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_FL_Row + HinRowNum_InExcel, Imp_HIN_FL_Col).Value =
                                    $"Hall Indicator {stop_i} FL"
                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_Row + HinRowNum_InExcel, Imp_HIN_Col).Value =
                                    arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)

                            HinRowNum_InExcel += 2
                        End If
                    ElseIf HinLiftSame_cnt > 2 Then
                        If HinPoint_bool = False Then
                            JobMaker_Form.ResultOutput_TextBox.Text += $".........{vbCrLf}"
                            HinPoint_bool = True

                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_FL_Row + HinRowNum_InExcel, Imp_HIN_FL_Col).Value =
                                    $":"
                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_Row + HinRowNum_InExcel, Imp_HIN_Col).Value =
                                    $":"

                            HinRowNum_InExcel += 2

                        End If
                        If HinFLDiff_bool Then
                            'HinLiftSame_cnt = 0
                            JobMaker_Form.ResultOutput_TextBox.Text +=
                                $"Hall Indicator {stop_i} FL : {arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)}{vbCrLf}"

                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_FL_Row + HinRowNum_InExcel, Imp_HIN_FL_Col).Value =
                                    $"Hall Indicator {stop_i} FL"
                            msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_Row + HinRowNum_InExcel, Imp_HIN_Col).Value =
                                    arr_liftStopFL_userContent(lift_i - 1, stop_i - 1)

                            HinRowNum_InExcel += 2
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

            If topFL_End_bool = False Then
                JobMaker_Form.ResultOutput_TextBox.Text +=
                    $"Hall Indicator TOP FL : {arr_liftStopFL_userContent(lift_i - test, stop_i - 2)}{vbCrLf}"

                msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_FL_Row + HinRowNum_InExcel, Imp_HIN_FL_Col).Value =
                    $"Hall Indicator TOP FL"
                msExcel_workbook.Worksheets(Imp_SheetName).Cells(Imp_HIN_Row + HinRowNum_InExcel, Imp_HIN_Col).Value =
                    arr_liftStopFL_userContent(lift_i - test, stop_i - 2)
            End If
        Else
            'JobMaker_Form.ResultFailOutput_TextBox.Text += ("<提醒> 重要 分頁未輸出，原因:分頁未打勾")
            'JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.Basic_TabPage
            'Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.Basic_TabPage.Text}」未使用是否重來?"), vbYesNo)
            'If basic_result.Yes And msExcel_workbook IsNot Nothing Then
            '    msExcel_workbook.Close()
            '    msExcel_app.Quit()
            'End If
        End If
    End Sub

    Public Sub Spec_MMIC(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application)
        'MMIC
        If JobMaker_Form.Use_mmic_CheckBox.Checked Then
            Dim with_val, without_val As String
            with_val =
                getMathOnExcel.getValue_formNameManager(msExcel_workbook,
                                                   get_NameManager.read_DbmsData(get_NameManager.SetTable_RESULT_WITH,
                                                                                 get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                 get_NameManager.SQLite_connectionPath_Tool,
                                                                                 get_NameManager.SQLite_ToolDBMS_Name))
            'msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_RESULT_WITH,
            '                                                              get_NameManager.SQLite_tableName_NameManager_TW,
            '                                                              get_NameManager.SQLite_connectionPath_Tool,
            '                                                              get_NameManager.SQLite_ToolDBMS_Name)
            '                                ).RefersToRange.Value '取得 有 內的文字內容
            without_val =
                getMathOnExcel.getValue_formNameManager(msExcel_workbook,
                                                   get_NameManager.read_DbmsData(get_NameManager.SetTable_RESULT_WITHOUT,
                                                                                 get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                 get_NameManager.SQLite_connectionPath_Tool,
                                                                                 get_NameManager.SQLite_ToolDBMS_Name))
            'msExcel_workbook.Names.Item(get_NameManager.read_DbmsData(get_NameManager.SetTable_RESULT_WITH,
            '                                                              get_NameManager.SQLite_tableName_NameManager_TW,
            '                                                              get_NameManager.SQLite_connectionPath_Tool,
            '                                                              get_NameManager.SQLite_ToolDBMS_Name)
            '                                ).RefersToRange.Value '取得 有 內的文字內容

            Dim usr_MMIC_MachineType, usr_MMIC_FLEX,
                usr_MMIC_CP43x, usr_MMIC_CarObj,
                usr_MMIC_E_BASE, usr_MMIC_E_CarObj,
                usr_SV_CarObj, usr_SV_E_CarObj,
                usr_SV_E_BASE,
                usr_VD10_CarObj,
                usr_VD10_ROM_Device, usr_VD10_Quantity As String


            'usr_MMIC_MachineType =
            '    JobMaker_Form.MMIC_MachineType_ComboBox.Name
            'usr_MMIC_FLEX =
            '    JobMaker_Form.MMIC_FLEX_N_ComboBox.Name

            usr_MMIC_CP43x =
                JobMaker_Form.MMIC_MR_CP43x_ComboBox.Name
            usr_MMIC_CarObj =
                JobMaker_Form.MMIC_MR_NumericUpDown.Name
            usr_MMIC_E_BASE =
                JobMaker_Form.MMIC_MR_EBase_ComboBox.Name
            usr_MMIC_E_CarObj =
                JobMaker_Form.MMIC_MR_E_NumericUpDown.Name

            usr_SV_CarObj =
                JobMaker_Form.MMIC_SV_NumericUpDown.Name
            usr_SV_E_BASE =
                JobMaker_Form.MMIC_SV_EBase_ComboBox.Name
            usr_SV_E_CarObj =
                JobMaker_Form.MMIC_SV_E_NumericUpDown.Name

            usr_VD10_ROM_Device =
                JobMaker_Form.MMIC_VD10_ROM_ComboBox.Name
            usr_VD10_Quantity =
                JobMaker_Form.MMIC_VD10_Quantity_ComboBox.Name
            usr_VD10_CarObj =
                JobMaker_Form.MMIC_VD10_NumericUpDown.Name

            Dim usrInput_MMIC_arr() As String = {usr_MMIC_CP43x, usr_MMIC_CarObj,
                                                 usr_MMIC_E_BASE, usr_MMIC_E_CarObj,
                                                 usr_SV_CarObj, usr_SV_E_CarObj,
                                                 usr_SV_E_BASE,
                                                 usr_VD10_CarObj,
                                                 usr_VD10_ROM_Device, usr_VD10_Quantity}
            'Dim i_mmicStr As String
            Dim dyCtrlName As DynamicControlName = New DynamicControlName
            dyCtrlName.JobMaker_MMICInfo()

            '輸入相對應的MMIC值
            For Each i_mmicStr As String In usrInput_MMIC_arr
                If i_mmicStr <> "" Then
                    Try
                        Select Case i_mmicStr
                            'Case usr_MMIC_MachineType
                                '[機種]
                                'excelWriteIn(JobMaker_Form.MMIC_MachineType_ComboBox.Text,
                                '             get_NameManager.read_DbmsData(get_NameManager.MMIC_CAR_TYPE,
                                '                                           get_NameManager.SQLite_tableName_NameManager_TW,
                                '                                           get_NameManager.SQLite_connectionPath_Tool,
                                '                                           get_NameManager.SQLite_ToolDBMS_Name),
                                '            msExcel_workbook)
                            'Case usr_MMIC_FLEX
                                '[FLEX-N幾百]
                                'excelWriteIn(JobMaker_Form.MMIC_FLEX_N_ComboBox.Text,
                                '            get_NameManager.read_DbmsData(get_NameManager.MMIC_FLEX,
                                '                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                '                                          get_NameManager.SQLite_connectionPath_Tool,
                                '                                          get_NameManager.SQLite_ToolDBMS_Name),
                                '            msExcel_workbook)
                            Case usr_MMIC_CP43x
                                '[MR-MMIC > 有無CP43]
                                Dim mmic_cp43x As String
                                mmic_cp43x =
                                    get_NameManager.read_DbmsData(get_NameManager.MMIC_CP43x,
                                                                  get_NameManager.SQLite_tableName_NameManager_TW,
                                                                  get_NameManager.SQLite_connectionPath_Tool,
                                                                  get_NameManager.SQLite_ToolDBMS_Name)
                                Dim cp43x_val As String

                                cp43x_val =
                                    msExcel_workbook.Names.Item(mmic_cp43x).RefersToRange.Value '取得 有 內的文字內容

                                If JobMaker_Form.MMIC_MR_CP43x_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                    msExcel_workbook.Names.Item(mmic_cp43x
                                                                ).RefersToRange.Characters(InStr(cp43x_val, with_val), Len(with_val)).
                                                                Font.Strikethrough = True
                                Else
                                    msExcel_workbook.Names.Item(mmic_cp43x
                                                                ).RefersToRange.Characters(InStr(cp43x_val, without_val), Len(without_val)).
                                                                Font.Strikethrough = True
                                End If


                            Case usr_MMIC_E_BASE
                                '[MR-MMIC > EEPROM DATA > BASE]
                                excelWriteIn(JobMaker_Form.MMIC_MR_EBase_ComboBox.Text,
                                            get_NameManager.read_DbmsData(get_NameManager.MMIC_EBase,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name),
                                            msExcel_workbook)
                            Case usr_MMIC_E_CarObj
                                '[MR-MMIC > EEPROM DATA > 自動生成控制項]
                                If usr_MMIC_E_CarObj <> 0 Then

                                End If
                            Case usr_SV_E_BASE
                                '[SV > EEPROM DATA > BASE]
                                excelWriteIn(JobMaker_Form.MMIC_SV_EBase_ComboBox.Text,
                                            get_NameManager.read_DbmsData(get_NameManager.SV_EBase,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name),
                                            msExcel_workbook)
                            Case usr_SV_E_CarObj
                                '[SV > 自動生成控制項]
                                If usr_SV_E_CarObj <> 0 Then

                                End If
                            Case usr_VD10_ROM_Device
                                '[VD10 > ROM DEVICE]
                                excelWriteIn(JobMaker_Form.MMIC_VD10_ROM_ComboBox.Text,
                                            get_NameManager.read_DbmsData(get_NameManager.VONIC_ROM_Device,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name),
                                            msExcel_workbook)
                            Case usr_VD10_Quantity
                                '[VD10 > QUANTITY 幾片]
                                excelWriteIn(JobMaker_Form.MMIC_VD10_Quantity_ComboBox.Text,
                                            get_NameManager.read_DbmsData(get_NameManager.VONIC_Quantity,
                                                                          get_NameManager.SQLite_tableName_NameManager_TW,
                                                                          get_NameManager.SQLite_connectionPath_Tool,
                                                                          get_NameManager.SQLite_ToolDBMS_Name),
                                            msExcel_workbook)
                            Case usr_MMIC_CarObj
                                '[MR-MMIC > 自動生成控制項]
                                Dim JM_MMIC_MR As String() = {get_NameManager.read_DbmsData(get_NameManager.MMIC_CarNo,
                                                                                            get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                            get_NameManager.SQLite_connectionPath_Tool,
                                                                                            get_NameManager.SQLite_ToolDBMS_Name),
                                                              get_NameManager.read_DbmsData(get_NameManager.MMIC_CarObj,
                                                                                            get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                            get_NameManager.SQLite_connectionPath_Tool,
                                                                                            get_NameManager.SQLite_ToolDBMS_Name)}

                                Dim JM_MMIC_MR_E As String() = {get_NameManager.read_DbmsData(get_NameManager.MMIC_ECarNo,
                                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                                              get_NameManager.SQLite_ToolDBMS_Name),
                                                                get_NameManager.read_DbmsData(get_NameManager.MMIC_ECarObj,
                                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                                              get_NameManager.SQLite_ToolDBMS_Name)}

                                dynamicControl_writeInExcel_MMIC(JobMaker_Form.MMIC_MR_NumericUpDown, JobMaker_Form.MMIC_MR_E_NumericUpDown,
                                                                 get_NameManager.MMIC_CarNo,
                                                                 JM_MMIC_MR, JM_MMIC_MR_E,
                                                                 JobMaker_Form.MMIC_MR_Panel, JobMaker_Form.MMIC_MR_E_Panel,
                                                                 dyCtrlName.JobMaker_MMIC_MrBase_InfoName_Array.Count, dyCtrlName.JobMaker_MMIC_MrBase_InfoName_Array,
                                                                 dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array.Count, dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array,
                                                                 msExcel_workbook)
                            Case usr_SV_CarObj
                                '[SV > 自動生成控制項]

                                Dim JM_MMIC_SV As String() = {get_NameManager.read_DbmsData(get_NameManager.SV_CarNo,
                                                                                            get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                            get_NameManager.SQLite_connectionPath_Tool,
                                                                                            get_NameManager.SQLite_ToolDBMS_Name),
                                                              get_NameManager.read_DbmsData(get_NameManager.SV_CarObj,
                                                                                            get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                            get_NameManager.SQLite_connectionPath_Tool,
                                                                                            get_NameManager.SQLite_ToolDBMS_Name)}

                                Dim JM_MMIC_SV_E As String() = {get_NameManager.read_DbmsData(get_NameManager.SV_ECarNo,
                                                                                                get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                get_NameManager.SQLite_connectionPath_Tool,
                                                                                                get_NameManager.SQLite_ToolDBMS_Name),
                                                                get_NameManager.read_DbmsData(get_NameManager.SV_ECarObj,
                                                                                                get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                                get_NameManager.SQLite_connectionPath_Tool,
                                                                                                get_NameManager.SQLite_ToolDBMS_Name)}

                                dynamicControl_writeInExcel_MMIC(JobMaker_Form.MMIC_SV_NumericUpDown, JobMaker_Form.MMIC_SV_E_NumericUpDown,
                                                                 get_NameManager.SV_CarNo,
                                                                 JM_MMIC_SV, JM_MMIC_SV_E,
                                                                 JobMaker_Form.MMIC_SV_Panel, JobMaker_Form.MMIC_SV_E_Panel,
                                                                 dyCtrlName.JobMaker_MMIC_SvBase_InfoName_Array.Count, dyCtrlName.JobMaker_MMIC_SvBase_InfoName_Array,
                                                                 dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array.Count, dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array,
                                                                 msExcel_workbook)

                            Case usr_VD10_CarObj
                                '[SV > 自動生成控制項]
                                Dim JM_MMIC_VONIC As String() = {get_NameManager.read_DbmsData(get_NameManager.VONIC_CarNo,
                                                                                               get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                               get_NameManager.SQLite_connectionPath_Tool,
                                                                                               get_NameManager.SQLite_ToolDBMS_Name),
                                                                get_NameManager.read_DbmsData(get_NameManager.VONIC_CarObj,
                                                                                              get_NameManager.SQLite_tableName_NameManager_TW,
                                                                                              get_NameManager.SQLite_connectionPath_Tool,
                                                                                              get_NameManager.SQLite_ToolDBMS_Name)}
                                Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
                                dynamicControl_writeInExcel(JobMaker_Form.MMIC_VD10_NumericUpDown.Value,
                                                            get_NameManager.VONIC_CarNo,
                                                            JM_MMIC_VONIC,
                                                            JobMaker_Form.MMIC_VD10_Panel,
                                                            spec_stored.LoadStored_PanelType.SingleLayer_Panel,
                                                            dyCtrlName.JobMaker_MMIC_VD10Base_InfoName_Array.Count,
                                                            dyCtrlName.JobMaker_MMIC_VD10Base_InfoName_Array,
                                                            msExcel_workbook)
                        End Select
                    Catch ex As Exception
                        JobMaker_Form.ResultFailOutput_TextBox.Text +=
                            ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_mmicStr}>{vbCrLf}")
                    End Try
                Else
                    JobMaker_Form.ResultFailOutput_TextBox.Text +=
                            ($"<提醒> {i_mmicStr} 為空值寫入失敗")
                End If
            Next
        Else
            JobMaker_Form.ResultFailOutput_TextBox.Text += "<提醒> 重要設定 分頁未輸出，原因:分頁未打勾"
            JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.Important_TabPage
            Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.Important_TabPage.Text}仕樣分頁」未勾選是否重來?"), vbYesNo)
            If basic_result = DialogResult.Yes And msExcel_workbook IsNot Nothing Then
                returnError_isPageRestart = True
                'msExcel_workbook.Close()
                'msExcel_app.Quit()
            End If
        End If
    End Sub

    ''' <summary>
    ''' [暫不使用][MMIC > CarNo and Object Name的寫入方法]
    ''' </summary>
    ''' <param name="mPanel">取得產生TextBox的表格</param>
    ''' <param name="sheetName">分頁名稱</param>
    ''' <param name="CarNo_Col"></param>
    ''' <param name="CarNo_Row"></param>
    ''' <param name="ObjName_Col"></param>
    ''' <param name="ObjName_Row"></param>
    ''' <param name="wkBook"></param>
    Private Sub MMIC_CarNo_ObjectName_WriteIn(nNumericUpDown As NumericUpDown, mPanel As Panel,
                                              tb_name_CarNo As String, tb_name_CarObj As String,
                                              sheetName As String,
                                              CarNo_Col As Integer, CarNo_Row As Integer,
                                              ObjName_Col As Integer, ObjName_Row As Integer,
                                              wkBook As Excel.Workbook)
        Dim merge_num As Integer

        With wkBook.Worksheets(sheetName)
            For Each tb As TextBox In mPanel.Controls
                'For i = 1 To mPanel.Controls.Count - 2 / 2
                For i = 1 To nNumericUpDown.Value
                    If tb.Name = $"{tb_name_CarNo}_{i}" Then
                        .Cells(CarNo_Row + merge_num * i, CarNo_Col).Value = tb.Text
                        If .Cells(CarNo_Row + merge_num * i, CarNo_Col).MergeCells Then
                            merge_num = .Range(.Cells(CarNo_Row + merge_num * i, CarNo_Col),
                                               .Cells(CarNo_Row + merge_num * i, CarNo_Col)).MergeArea.Rows.Count
                        Else
                            .Range(.Cells(CarNo_Row + merge_num * i, CarNo_Col),
                                   .Cells(CarNo_Row + merge_num * i, CarNo_Col)).Insert
                        End If

                    ElseIf tb.Name = $"{tb_name_CarObj}_{i}" Then
                        .Cells(ObjName_Row + merge_num * i, ObjName_Col).Value = tb.Text
                        If .Cells(ObjName_Row + merge_num * i, ObjName_Col).MergeCells Then
                            merge_num = .Range(.Cells(ObjName_Row + merge_num * i, ObjName_Col),
                                               .Cells(ObjName_Row + merge_num * i, ObjName_Col)).MergeArea.Rows.Count
                        Else
                            .Range(.Cells(ObjName_Row + merge_num * i, ObjName_Col),
                                   .Cells(ObjName_Row + merge_num * i, ObjName_Col)).Insert
                        End If

                    End If

                Next
            Next
        End With
    End Sub



    '寫入excel內的方法
    ''' <summary>
    ''' 將usr(輸入資料)寫入msExcel_workbook(目標excel)的spec(名稱管理員)
    ''' </summary>
    ''' <param name="usr"></param>
    ''' <param name="spec"></param>
    ''' <param name="msExcel_workbook"></param>
    Overloads Sub excelWriteIn(usr As String, spec As String, msExcel_workbook As Excel.Workbook)
        If usr IsNot "" Then
            returnError_specName = spec '錯誤回報
            msExcel_workbook.Names.Item(spec).RefersToRange.Value = usr

            JobMaker_Form.ResultOutput_TextBox.Text += ($"名稱管理員:{spec} / 值:{usr} 寫入成功{vbCrLf}")
        Else
            returnError_specName = spec '錯誤回報

            JobMaker_Form.ResultOutput_TextBox.Text += ($"名稱管理員:{spec} / 值:{usr} 是空值寫入失敗{vbCrLf}")
        End If
    End Sub
    ''' <summary>
    ''' 將chkbox有打勾的項目和usr(輸入資料)寫入msExcel_workbook(目標excel)的spec(名稱管理員)
    ''' </summary>
    ''' <param name="usr"></param>
    ''' <param name="spec"></param>
    ''' <param name="chkbox"></param>
    ''' <param name="msExcel_workbook"></param>
    Overloads Sub excelWriteIn(usr As String, spec As String, chkbox As CheckBox, msExcel_workbook As Excel.Workbook)
        If usr IsNot "" And chkbox.Checked Then
            returnError_specName = spec '錯誤回報
            msExcel_workbook.Names.Item(spec).RefersToRange.Value = usr

            JobMaker_Form.ResultOutput_TextBox.Text += ($"打勾框:{chkbox.Name} 狀態:{chkbox.CheckState} / 名稱管理員:{spec} / 值:{usr} 寫入成功{vbCrLf}")
        Else
            returnError_specName = spec '錯誤回報

            If usr = "" Then
                JobMaker_Form.ResultOutput_TextBox.Text += ($"名稱管理員:{spec} / 值:{usr} 是空值寫入失敗{vbCrLf}")
            End If
            If chkbox.Checked = CheckState.Unchecked Then
                JobMaker_Form.ResultOutput_TextBox.Text += ($"打勾框:{chkbox.Name} 狀態:{chkbox.CheckState} 寫入失敗{vbCrLf}")
            End If
        End If
    End Sub
    ''' <summary>
    ''' 將radbox有打勾的項目和usr(輸入資料)寫入msExcel_workbook(目標excel)的spec(名稱管理員)
    ''' </summary>
    ''' <param name="usr"></param>
    ''' <param name="spec"></param>
    ''' <param name="radbox"></param>
    ''' <param name="msExcel_workbook"></param>
    Overloads Sub excelWriteIn(usr As String, spec As String, radbox As RadioButton, msExcel_workbook As Excel.Workbook)
        If usr IsNot "" And radbox.Checked Then
            returnError_specName = spec '錯誤回報
            msExcel_workbook.Names.Item(spec).RefersToRange.Value = usr

            JobMaker_Form.ResultOutput_TextBox.Text +=
                ($"打勾框:{radbox.Name} 狀態:{radbox.Checked} / 名稱管理員:{spec} / 值:{usr} 寫入成功{vbCrLf}")
        Else
            returnError_specName = spec '錯誤回報

            If usr = "" Then
                JobMaker_Form.ResultOutput_TextBox.Text += ($"名稱管理員:{spec} / 值:{usr} 是空值寫入失敗{vbCrLf}")
            End If
            If radbox.Checked = CheckState.Unchecked Then
                JobMaker_Form.ResultOutput_TextBox.Text += ($"打勾框:{radbox.Name} 狀態:{radbox.Checked} 寫入失敗{vbCrLf}")
            End If
        End If
    End Sub
    ''' <summary>
    ''' 沒打勾的CheckBox寫入資料.usr=要輸入的資料/spec=名稱管理員/chkbox=未打勾的CheckBox
    ''' </summary>
    ''' <param name="usr"> 輸入資料 </param>
    ''' <param name="spec"> 名稱管理員 </param>
    ''' <param name="chkbox"></param>
    ''' <param name="msExcel_workbook"></param>
    Sub excelWriteIn_ForReverseState(usr As String, spec As String, chkbox As CheckBox, msExcel_workbook As Excel.Workbook)
        If usr IsNot "" And Not chkbox.Checked Then
            returnError_specName = spec '錯誤回報
            msExcel_workbook.Names.Item(spec).RefersToRange.Value = usr

            JobMaker_Form.ResultOutput_TextBox.Text += ($"打勾框:{chkbox.Name} 狀態:{chkbox.CheckState} / 名稱管理員:{spec} / 值:{usr} 寫入成功{vbCrLf}")
        Else
            returnError_specName = spec '錯誤回報

            If usr = "" Then
                JobMaker_Form.ResultOutput_TextBox.Text += ($"名稱管理員:{spec} / 值:{usr} 是空值寫入失敗{vbCrLf}")
            End If
            If chkbox.Checked = CheckState.Unchecked Then
                JobMaker_Form.ResultOutput_TextBox.Text += ($"打勾框:{chkbox.Name} 狀態:{chkbox.Checked} 寫入失敗{vbCrLf}")
            End If

        End If
    End Sub

    '寫入excel內的方法
    ''' <summary>
    ''' 將 名稱管理員chkboxName(為圖形) 寫入workbook中的sheetPageName分頁名稱中
    ''' </summary>
    ''' <param name="chkboxName"> excel中checkBox的圖形名稱 </param>
    ''' <param name="sheetPageName"> excel中分頁名稱 </param>
    ''' <param name="msExcel_workbook"> workbook名稱 </param>
    Overloads Sub chkboxWriteIn(chkboxName As String, sheetPageName As String, msExcel_workbook As Excel.Workbook)
        If chkboxName IsNot "" Then
            msExcel_workbook.Sheets(sheetPageName).CheckBoxes(chkboxName).value = True

            JobMaker_Form.ResultOutput_TextBox.Text += ($"圖形名稱:{chkboxName} / 分頁名稱:{sheetPageName} 打勾成功{vbCrLf}")
        Else
            JobMaker_Form.ResultOutput_TextBox.Text += ($"圖形名稱:{chkboxName} / 分頁名稱:{sheetPageName} 打勾失敗{vbCrLf}")
        End If
    End Sub




    ''' <summary>
    ''' 判斷CheckList的主項目中Yes or No的選項哪個被勾選，被勾選的會回傳名稱管理員
    ''' </summary>
    ''' <param name="rdbtn_n"> RadioButton 為 NO 的 </param>
    ''' <param name="rdbtn_y"> RadioButton 為 YES 的 </param>
    ''' <param name="chk_QN_name"> RadioButton NO的名稱管理員名字 </param>
    ''' <param name="chk_QY_name"> RadioButton YES的名稱管理員名字 </param>
    ''' <returns></returns>
    Overloads Function chkBoxStateRead(rdbtn_n As RadioButton, rdbtn_y As RadioButton, chk_QN_name As String, chk_QY_name As String) As String
        If rdbtn_n.Checked Then
            chkBoxStateRead = chk_QN_name
        ElseIf rdbtn_y.Checked Then
            chkBoxStateRead = chk_QY_name
        End If

        Return chkBoxStateRead
    End Function
    ''' <summary>
    ''' 判斷CheckList的主項目中單一選項被勾選，勾選的回傳名稱管理員
    ''' </summary>
    ''' <param name="rdbtn"> CheckBox 被打勾的 </param>
    ''' <param name="chk_draw_name"> 回傳CheckList中CheckBox的名稱管理員 </param>
    ''' <returns></returns>
    Overloads Function chkBoxStateRead(rdbtn As CheckBox, chk_draw_name As String)
        If rdbtn.Checked Then
            Return chk_draw_name
        End If
    End Function





    ''' <summary>
    ''' [數字 月 轉換成英文]
    ''' </summary>
    ''' <returns></returns>
    Private Function monthTransfer_sub() As String
        Dim mon As Integer
        mon = JobMaker_Form.Basic_DrawDate_DateTimePicker.Value.Month
        Select Case mon
            Case 1
                Return "Jan"
            Case 2
                Return "Feb"
            Case 3
                Return "Mar"
            Case 4
                Return "Apr"
            Case 5
                Return "May"
            Case 6
                Return "Jun"
            Case 7
                Return "Jul"
            Case 8
                Return "Aug"
            Case 9
                Return "Sep"
            Case 10
                Return "Oct"
            Case 11
                Return "Nov"
            Case 12
                Return "Dec"
        End Select

    End Function


    Private Sub If79xID_Input(id As String, str As Boolean)
        'Dim if79_row, if79_col, i_n As Integer
        'Dim if79_sheet As String

        'if79_row = msExcel_workbook.Names.Item(id).RefersToRange.Row
        'if79_col = msExcel_workbook.Names.Item(id).RefersToRange.Column
        'if79_sheet = msExcel_workbook.Names.Item(id).RefersToRange.Worksheet.Name
        'i_n = 1
        'Do
        '    i_n = i_n + 1
        '    For i = 1 To i_n
        '        If msExcel_workbook.Worksheets(if79_sheet).Cells(if79_row + i, if79_col).value <> "" Then
        '            Exit Do
        '        End If
        '    Next
        'Loop
        'For j = 0 To i_n - 1
        '    For k = 0 To 2
        '        msExcel_workbook.Worksheets(if79_sheet).Cells(if79_row + j, if79_col + k).font.Strikethrough = str
        '    Next
        'Next
        'if79_row = 0
        'if79_col = 0
        'if79_sheet = ""
        'i_n = 1

    End Sub
End Class
