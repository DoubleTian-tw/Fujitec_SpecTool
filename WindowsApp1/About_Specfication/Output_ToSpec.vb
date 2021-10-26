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

    Public Sub Spec_FinalCheck(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application)
        Dim spec_item As Spec_Item = New Spec_Item()
        spec_item.ini_specTW_AllControler()


        Dim mCtrlNameForError As String = ""
        Dim mPanelNameForError As String = ""
        Try
            '全部仕樣確認表欄與列
            '項目
            Dim finalCheck_item_col As Integer =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_Item)
            Dim finalCheck_item_row As Integer =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_Item)
            '有無
            Dim finalCheck_state_col As Integer =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_State)
            Dim finalCheck_state_row As Integer =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_State)
            '仕樣
            Dim finalCheck_Spec_col As Integer =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.FinalCheck_Spec)
            Dim finalCheck_Spec_row As Integer =
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
                                                                     get_NameManager.FinalCheck_Item,
                                                                     item_countRow,
                                                                     item_number)
                            getMathOnExcel.
                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                     get_NameManager.FinalCheck_State,
                                                                     item_countRow,
                                                                     "O")

                            'Label取代Panel後的名稱，例如:A_Panel > A_Label
                            Dim afterReplaceTitle_Label As String =
                                spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mPanel, ctrlTypeName_Panel, ctrlTypeName_Label)
                            getMathOnExcel.
                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                     get_NameManager.FinalCheck_Spec,
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
                                                                                 get_NameManager.FinalCheck_State,
                                                                                 item_countRow,
                                                                                 "O")

                                        getMathOnExcel.
                                            setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                 get_NameManager.FinalCheck_Spec,
                                                                                 item_countRow,
                                                                                 getLabelText)
                                    ElseIf TypeOf (mCtrlContent) Is TextBox And mCtrlContent.Text <> "" And mCtrlContent.Enabled = True Then
                                        '其他TextBox
                                        item_countRow += 1

                                        Dim nameAfterReplace_ChkBox As String =
                                            spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mCtrlContent, ctrlTypeName_TextBox, ctrlTypeName_CheckBox)
                                        Dim nameAfterReplace_Label As String =
                                            spec_item.repalce_replaceName_to_myCtrlType_inMyCtrl(mCtrlContent, ctrlTypeName_TextBox, ctrlTypeName_Label)

                                        '如果控制項為CheckBox時的狀態，僅打勾的才輸出 -----
                                        Dim is_ChkBox_checked As Boolean =
                                            spec_item.getRelace_ChkBoxState_onPanel(nameAfterReplace_ChkBox, mPanel)
                                        '----- 如果控制項為CheckBox時的狀態，僅打勾的才輸出

                                        Dim getLabelText As String =
                                            spec_item.getRelace_NameText_onPanel(nameAfterReplace_Label, mPanel)

                                        If is_ChkBox_checked = True Then
                                            getMathOnExcel.
                                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                     get_NameManager.FinalCheck_State,
                                                                                     item_countRow,
                                                                                     "O")
                                            Dim getChkBoxText As String =
                                                spec_item.getRelace_NameText_onPanel(nameAfterReplace_ChkBox, mPanel)

                                            getMathOnExcel.
                                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                     get_NameManager.FinalCheck_Spec,
                                                                                     item_countRow,
                                                                                     $"{getChkBoxText} : {mCtrlContent.Text}")
                                        End If

                                        If getLabelText <> "" And mCtrlContent.Text <> "" Then
                                            getMathOnExcel.
                                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                     get_NameManager.FinalCheck_State,
                                                                                     item_countRow,
                                                                                     "O")

                                            getMathOnExcel.
                                                setValue_to_Cells_addBelow_onWorksht(msExcel_workbook,
                                                                                     get_NameManager.FinalCheck_Spec,
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
            Dim usrInput_arr As New ArrayList
            Dim usr_Local As String = JobMaker_Form.Basic_Local_ComboBox.Name
            usrInput_arr.Add(usr_Local)

            Dim usr_JobNo_New As String = JobMaker_Form.Basic_JobNoNew_TextBox.Name
            usrInput_arr.Add(usr_JobNo_New)

            Dim usr_JobNo_Old As String = JobMaker_Form.Basic_JobNoOld_TextBox.Name
            usrInput_arr.Add(usr_JobNo_Old)

            Dim usr_JobNo_Mod As String = JobMaker_Form.Basic_JobNoMOD_TextBox.Name
            usrInput_arr.Add(usr_JobNo_Mod)

            Dim usr_JobName As String = JobMaker_Form.Basic_JobName_TextBox.Name
            usrInput_arr.Add(usr_JobName)

            Dim usr_Designer, usr_Checker, usr_Approver As String
            Dim usr_drawDate As String = JobMaker_Form.Basic_DrawDate_DateTimePicker.Name
            usrInput_arr.Add(usr_drawDate)


            usr_Designer = ""
            If JobMaker_Form.Basic_DesingerChinese_ComboBox.Text <> "" Then
                usr_Designer =
                        JobMaker_Form.Basic_DesingerChinese_ComboBox.Name '設計者中文
            ElseIf JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text <> "" Then
                usr_Designer =
                        JobMaker_Form.Basic_DesingerEnglish_ComboBox.Name '設計者英文
            End If
            usrInput_arr.Add(usr_Designer)

            usr_Checker = ""
            If JobMaker_Form.Basic_CheckerChinese_ComboBox.Text <> "" Then
                usr_Checker =
                        JobMaker_Form.Basic_CheckerChinese_ComboBox.Name '檢查者中文
            ElseIf JobMaker_Form.Basic_DesingerChinese_ComboBox.Text <> "" Then
                usr_Checker =
                        JobMaker_Form.Basic_CheckerEnglish_ComboBox.Name '檢查者英文
            End If
            usrInput_arr.Add(usr_Checker)


            usr_Approver = ""
            If JobMaker_Form.Basic_ApproverChinese_ComboBox.Text <> "" Then
                usr_Approver =
                        JobMaker_Form.Basic_ApproverChinese_ComboBox.Name '承認者中文
            ElseIf JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text <> "" Then
                usr_Approver =
                        JobMaker_Form.Basic_ApproverEnglish_ComboBox.Name '承認者英文
            End If
            usrInput_arr.Add(usr_Approver)


            For Each i_str In usrInput_arr
                'Try
                Select Case i_str
                    Case usr_Local
                        Dim local As String = JobMaker_Form.Basic_Local_ComboBox.Text
                        Dim temp_local As String
                        temp_local = ""
                        Select Case local
                            Case "新竹"
                                temp_local = get_NameManager.DWG_HsinChu
                            Case "台南"
                                temp_local = get_NameManager.DWG_Tainan
                            Case "台中"
                                temp_local = get_NameManager.DWG_Taichung
                            Case "台北"
                                temp_local = get_NameManager.DWG_Taipei
                            Case "高雄"
                                temp_local = get_NameManager.DWG_Kaohsiung
                            Case "桃園"
                                temp_local = get_NameManager.DWG_Taoyuan
                        End Select

                        excelWriteIn("1",
                                     temp_local,
                                     msExcel_workbook)
                        'New工番號
                    Case usr_JobNo_New
                        excelWriteIn(JobMaker_Form.Basic_JobNoNew_TextBox.Text,
                                    get_NameManager.JOBNO,
                                    msExcel_workbook)
                            'Old工番號
                    Case usr_JobNo_Old
                        excelWriteIn(JobMaker_Form.Basic_JobNoOld_TextBox.Text,
                                    get_NameManager.JOBNO_OLD,
                                    msExcel_workbook)
                            'Mod工番號
                    Case usr_JobNo_Mod
                        excelWriteIn(JobMaker_Form.Basic_JobNoMOD_TextBox.Text,
                                    get_NameManager.JOBNO_MOD,
                                    msExcel_workbook)
                        '工番名
                    Case usr_JobName
                        excelWriteIn(JobMaker_Form.Basic_JobName_TextBox.Text,
                                     get_NameManager.JOBNAME,
                                     msExcel_workbook)
                        '設計者
                    Case usr_Designer
                        Dim temp_designer As String = ""
                        If JobMaker_Form.Basic_DesingerChinese_ComboBox.Text <> "" Then
                            temp_designer = JobMaker_Form.Basic_DesingerChinese_ComboBox.Text
                        ElseIf JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text <> "" Then
                            temp_designer = JobMaker_Form.Basic_DesingerEnglish_ComboBox.Text
                        End If

                        excelWriteIn(temp_designer,
                                    get_NameManager.DESIGENED,
                                    msExcel_workbook)
                        '審查者
                    Case usr_Checker
                        Dim temp_checker As String = ""
                        If JobMaker_Form.Basic_CheckerChinese_ComboBox.Text <> "" Then
                            temp_checker = JobMaker_Form.Basic_CheckerChinese_ComboBox.Text
                        ElseIf JobMaker_Form.Basic_CheckerEnglish_ComboBox.Text <> "" Then
                            temp_checker = JobMaker_Form.Basic_CheckerEnglish_ComboBox.Text
                        End If

                        excelWriteIn(temp_checker,
                                    get_NameManager.CHECKED,
                                    msExcel_workbook)
                        '承認者
                    Case usr_Approver
                        Dim temp_approver As String = ""
                        If JobMaker_Form.Basic_ApproverChinese_ComboBox.Text <> "" Then
                            temp_approver = JobMaker_Form.Basic_ApproverChinese_ComboBox.Text
                        ElseIf JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text <> "" Then
                            temp_approver = JobMaker_Form.Basic_ApproverEnglish_ComboBox.Text
                        End If

                        excelWriteIn(temp_approver,
                                     get_NameManager.APPROVED,
                                     msExcel_workbook)

                        '作圖日
                    Case usr_drawDate
                        Dim usr_drawDate_val As String
                        usr_drawDate_val =
                                $"{monthTransfer_sub()}.{JobMaker_Form.Basic_DrawDate_DateTimePicker.Value.Day}.{JobMaker_Form.Basic_DrawDate_DateTimePicker.Value.Year}" 'Date出圖時間

                        excelWriteIn(usr_drawDate_val,
                                     get_NameManager.DRAW_DATE,
                                     msExcel_workbook)
                End Select
                'Catch ex As Exception
                'JobMaker_Form.ResultOutput_TextBox.Text +=
                '    ($"<{JobMaker_Form.JMFileCho_Spec_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_str}>{vbCrLf}")
                'JobMaker_Form.ResultOutput_TextBox.Text +=
                '$"----------------------------------"
                'End Try
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
    Public Sub Spec_ChkList_Std(msExcel_workbook As Excel.Workbook, msExcel_app As Excel.Application) '基本分頁內容
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


            '儲存每一個使用者輸入的值
            usrInput_arr = {usr_JobNo_New, usr_JobNo_Old, usr_JobNo_MOD, usr_JobName,
                            usr_Designer, usr_Checker, usr_Approver, usr_Local,
                            usr_DrawDate}

            '輸入相對應的基本值
            For Each i_str In usrInput_arr
                If i_str <> "" Then
                    'Try
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
                    'Catch ex As Exception
                    'JobMaker_Form.ResultOutput_TextBox.Text +=
                    '    ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_str}>{vbCrLf}")
                    'End Try
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
            Dim usr_ChkList_JobName As String = get_NameManager.ChkList_JOBNAME
            Dim usr_ChkList_JobNo As String = get_NameManager.ChkList_JOBNO
            '--------------------------------------------------------------------------- Check List > CheckList_P1 > 名稱管理員 

            ' Check List > CheckList_P1/P2 分頁名稱 ---------------------------------------------------------------------------
            Dim chkListP1_ShtName As String = get_NameManager.ChkList_P1_PageName
            Dim chkListP2_ShtName As String = get_NameManager.ChkList_P2_PageName
            '--------------------------------------------------------------------------- Check List > CheckList_P1/P2 分頁名稱 

            ' Check List > CheckList_P1 > 名稱管理員 ---------------------------------------------------------------------------
            Dim chkList_PA_ChkBox As String = get_NameManager.ChkList_PA_ChkBox
            Dim chkList_OS_ChkBox As String = get_NameManager.ChkList_OS_ChkBox
            Dim chkList_CFM_ChkBox As String = get_NameManager.ChkList_CFM_ChkBox
            Dim chkList_ELE_ChkBox As String = get_NameManager.ChkList_ELE_ChkBox
            '--------------------------------------------------------------------------- Check List > CheckList_P1 > 名稱管理員

            ' CheckList中日期的CheckBox ---------------------------------------------------------------------------------------
            Dim usr_nameManager_chkList_PA As String = chkBoxStateRead(JobMaker_Form.ChkList_PaSheet_CheckBox, chkList_PA_ChkBox)
            Dim usr_nameManager_chkList_OS As String = chkBoxStateRead(JobMaker_Form.ChkList_OS_CheckBox, chkList_OS_ChkBox)
            Dim usr_nameManager_chkList_CFM As String = chkBoxStateRead(JobMaker_Form.ChkList_Confirm_CheckBox, chkList_CFM_ChkBox)
            Dim usr_nameManager_chkList_ELE As String = chkBoxStateRead(JobMaker_Form.ChkList_Elec_CheckBox, chkList_ELE_ChkBox)
            '--------------------------------------------------------------------------------------- CheckList中日期的CheckBox 


            Dim usr_nameManager_chkList_Q1 As String =
                chkBoxStateRead(JobMaker_Form.ChkList_1_no_RadioButton, JobMaker_Form.ChkList_1_yes_RadioButton,
                                get_NameManager.ChkList_Q1No_ChkBox,
                                get_NameManager.ChkList_Q1Yes_ChkBox)
            Dim usr_nameManager_chkList_Q2 As String =
                chkBoxStateRead(JobMaker_Form.ChkList_2_no_RadioButton, JobMaker_Form.ChkList_2_yes_RadioButton,
                                get_NameManager.ChkList_Q2No_ChkBox,
                                get_NameManager.ChkList_Q2Yes_ChkBox)
            Dim usr_nameManager_chkList_Q3 As String =
                chkBoxStateRead(JobMaker_Form.ChkList_3_no_RadioButton, JobMaker_Form.ChkList_3_yes_RadioButton,
                                get_NameManager.ChkList_Q3No_ChkBox,
                                get_NameManager.ChkList_Q3Yes_ChkBox)

            ' 5.VONIC -------------------------------------------------------
            Dim usr_nameManager_chkList_Q5_chkBoxState As String = ""
            If JobMaker_Form.ChkList_5_no_RadioButton.Checked Then
                usr_nameManager_chkList_Q5_chkBoxState = get_NameManager.ChkList_Q5No_ChkBox
            ElseIf JobMaker_Form.ChkList_5_nstd_RadioButton.Checked Then
                usr_nameManager_chkList_Q5_chkBoxState = get_NameManager.ChkList_Q5NoStd_ChkBox
            ElseIf JobMaker_Form.ChkList_5_std_RadioButton.Checked Then
                usr_nameManager_chkList_Q5_chkBoxState = get_NameManager.ChkList_Q5Std_ChkBox
            End If
            '------------------------------------------------------- 5.VONIC 

            Dim usr_nameManager_chkList_Q6 As String =
                chkBoxStateRead(JobMaker_Form.ChkList_6_no_RadioButton, JobMaker_Form.ChkList_6_yes_RadioButton,
                                get_NameManager.ChkList_Q6No_ChkBox,
                                get_NameManager.ChkList_Q6Yes_ChkBox)

            Dim usr_nameManager_chkList_Q6_yes As String =
                chkBoxStateRead(JobMaker_Form.ChkList_6_yesChk_RadioButton, JobMaker_Form.ChkList_6_yesItem_RadioButton,
                                get_NameManager.ChkList_Q6YesChk_ChkBox,
                                get_NameManager.ChkList_Q6YesItem_ChkBox)

            Dim usr_nameManager_chkList_Q7 As String =
                chkBoxStateRead(JobMaker_Form.ChkList_7_no_RadioButton, JobMaker_Form.ChkList_7_yes_RadioButton,
                                get_NameManager.ChkList_Q7No_ChkBox,
                                get_NameManager.ChkList_Q7Yes_ChkBox)

            Dim usr_nameManager_chkList_Q8 As String =
                chkBoxStateRead(JobMaker_Form.ChkList_8_no_RadioButton, JobMaker_Form.ChkList_8_yes_RadioButton,
                                get_NameManager.ChkList_Q8No_ChkBox,
                                get_NameManager.ChkList_Q8Yes_ChkBox)

            Dim usr_nameManager_chkList_Q9 As String =
                chkBoxStateRead(JobMaker_Form.ChkList_9_no_RadioButton, JobMaker_Form.ChkList_9_yes_RadioButton,
                                get_NameManager.ChkList_Q9No_ChkBox,
                                get_NameManager.ChkList_Q9Yes_ChkBox)

            ' 程式變更/2.使用裝置　--------------------------------------------
            Dim usr_nameManager_prgm_2_Test As String = ""
            Dim usr_nameManager_prgm_2_COP As String = ""
            Dim usr_nameManager_prgm_2_Tower As String = ""
            Dim usr_nameManager_prgm_2_Other As String = ""
            If JobMaker_Form.PrmList_2_test_CheckBox.Checked Then
                usr_nameManager_prgm_2_Test = get_NameManager.ChkList_Prgm_2_Test_ChkBox
            ElseIf JobMaker_Form.PrmList_2_COP_CheckBox.Checked Then
                usr_nameManager_prgm_2_COP = get_NameManager.ChkList_Prgm_2_COP_ChkBox
            ElseIf JobMaker_Form.PrmList_2_Tower_CheckBox.Checked Then
                usr_nameManager_prgm_2_Tower = get_NameManager.ChkList_Prgm_2_Tower_ChkBox
            ElseIf JobMaker_Form.PrmList_2_Other_CheckBox.Checked Then
                usr_nameManager_prgm_2_Other = get_NameManager.ChkList_Prgm_2_Other_ChkBox
            End If
            '-------------------------------------------- 程式變更/2.使用裝置　

            ' 程式變更/3.檢查方法　--------------------------------------------
            Dim usr_nameManager_prgm_3_Debug As String = ""
            Dim usr_nameManager_prgm_3_Test As String = ""
            Dim usr_nameManager_prgm_3_CFM As String = ""
            Dim usr_nameManager_prgm_3_EXE As String = ""
            Dim usr_nameManager_prgm_3_Other As String = ""
            If JobMaker_Form.PrmList_3_debug_CheckBox.Checked Then
                usr_nameManager_prgm_3_Debug = get_NameManager.ChkList_Prgm_3_Debug_ChkBox
            ElseIf JobMaker_Form.PrmList_3_test_CheckBox.Checked Then
                usr_nameManager_prgm_3_Test = get_NameManager.ChkList_Prgm_3_Test_ChkBox
            ElseIf JobMaker_Form.PrmList_3_confirm_CheckBox.Checked Then
                usr_nameManager_prgm_3_CFM = get_NameManager.ChkList_Prgm_3_CFM_ChkBox
            ElseIf JobMaker_Form.PrmList_3_excute_CheckBox.Checked Then
                usr_nameManager_prgm_3_EXE = get_NameManager.ChkList_Prgm_3_EXE_ChkBox
            ElseIf JobMaker_Form.PrmList_3_other_Checkbox.Checked Then
                usr_nameManager_prgm_3_Other = get_NameManager.ChkList_Prgm_3_Other_ChkBox
            End If
            '-------------------------------------------- 程式變更/3.檢查方法　

            ' 程式變更/4.檢查結果　-------------------------------------------------------------------------------------------------------------
            Dim usr_nameManager_prgm_Auto As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no1_RadioButton, JobMaker_Form.PrmList_4_yes1_RadioButton,
                                get_NameManager.ChkList_Prgm_4_1No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_1Yes_ChkBox)

            Dim usr_nameManager_prgm_Output As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no2_RadioButton, JobMaker_Form.PrmList_4_yes2_RadioButton,
                                get_NameManager.ChkList_Prgm_4_2No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_2Yes_ChkBox)

            Dim usr_nameManager_prgm_INI As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no3_RadioButton, JobMaker_Form.PrmList_4_yes3_RadioButton,
                                get_NameManager.ChkList_Prgm_4_3No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_3Yes_ChkBox)

            Dim usr_nameManager_prgm_Case As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no4_RadioButton, JobMaker_Form.PrmList_4_yes4_RadioButton,
                                get_NameManager.ChkList_Prgm_4_4No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_4Yes_ChkBox)

            Dim usr_nameManager_prgm_IF As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no5_RadioButton, JobMaker_Form.PrmList_4_yes5_RadioButton,
                                get_NameManager.ChkList_Prgm_4_5No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_5Yes_ChkBox)

            Dim usr_nameManager_prgm_Loop As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no6_RadioButton, JobMaker_Form.PrmList_4_yes6_RadioButton,
                                get_NameManager.ChkList_Prgm_4_6No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_6Yes_ChkBox)

            Dim usr_nameManager_prgm_Range As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no7_RadioButton, JobMaker_Form.PrmList_4_yes7_RadioButton,
                                get_NameManager.ChkList_Prgm_4_7No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_7Yes_ChkBox)

            Dim usr_nameManager_prgm_Casting As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no8_RadioButton, JobMaker_Form.PrmList_4_yes8_RadioButton,
                                get_NameManager.ChkList_Prgm_4_8No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_8Yes_ChkBox)

            Dim usr_nameManager_prgm_0 As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no9_RadioButton, JobMaker_Form.PrmList_4_yes9_RadioButton,
                                get_NameManager.ChkList_Prgm_4_9No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_9Yes_ChkBox)

            Dim usr_nameManager_prgm_Count As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no10_RadioButton, JobMaker_Form.PrmList_4_yes10_RadioButton,
                                get_NameManager.ChkList_Prgm_4_10No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_10Yes_ChkBox)

            Dim usr_nameManager_prgm_Address As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no11_RadioButton, JobMaker_Form.PrmList_4_yes11_RadioButton,
                                get_NameManager.ChkList_Prgm_4_11No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_11Yes_ChkBox)

            Dim usr_nameManager_prgm_Custom As String =
                chkBoxStateRead(JobMaker_Form.PrmList_4_no12_RadioButton, JobMaker_Form.PrmList_4_yes12_RadioButton,
                                get_NameManager.ChkList_Prgm_4_12No_ChkBox,
                                get_NameManager.ChkList_Prgm_4_12Yes_ChkBox)
            '------------------------------------------------------------------------------------------------------------- 程式變更/4.檢查結果　
            '取得各check box content中的名稱
            Dim usr_chkList_PA_year As String = get_NameManager.ChkList_PA_Year
            Dim usr_chkList_PA_month As String = get_NameManager.ChkList_PA_Month
            Dim usr_chkList_PA_date As String = get_NameManager.ChkList_PA_Day
            Dim usr_chkList_OS_year As String = get_NameManager.ChkList_OS_Year
            Dim usr_chkList_OS_month As String = get_NameManager.ChkList_OS_Month
            Dim usr_chkList_OS_date As String = get_NameManager.ChkList_OS_Day
            Dim usr_chkList_CFM_year As String = get_NameManager.ChkList_CFM_Year
            Dim usr_chkList_CFM_month As String = get_NameManager.ChkList_CFM_Month
            Dim usr_chkList_CFM_date As String = get_NameManager.ChkList_CFM_Day
            Dim usr_chkList_ELE_year As String = get_NameManager.ChkList_ELE_Year
            Dim usr_chkList_ELE_month As String = get_NameManager.ChkList_ELE_Month
            Dim usr_chkList_ELE_date As String = get_NameManager.ChkList_ELE_Day

            Dim usr_chkList_Q1_YCont As String = JobMaker_Form.ChkList_1_yes_Content_TextBox.Name     '檢查項目1  有   討論內容
            Dim usr_chkList_Q1_YResult As String = JobMaker_Form.ChkList_1_yes_result_TextBox.Name    '檢查項目1  有   結果
            Dim usr_chkList_Q2_YCont As String = JobMaker_Form.ChkList_2_yes_Content_TextBox.Name     '檢查項目2  有   討論結果
            Dim usr_chkList_Q2_YResult As String = JobMaker_Form.ChkList_2_yes_Result_TextBox.Name    '檢查項目2  有   討論結果
            Dim usr_chkList_Q3_YMan As String = JobMaker_Form.ChkList_3_yes_Man_TextBox.Name          '檢查項目3  有   討論者
            Dim usr_chkList_Q3_YCont As String = JobMaker_Form.ChkList_3_yes_Content_TextBox.Name     '檢查項目3  有   討論內容
            Dim usr_chkList_Q3_YResult As String = JobMaker_Form.ChkList_3_yes_Result_TextBox.Name    '檢查項目3  有   討論結果
            Dim usr_chkList_Q4_MMIC As String = JobMaker_Form.ChkList_4_ObjName_TextBox.Name          '檢查項目4  有   MMIC
            Dim usr_chkList_Q4_SV As String = JobMaker_Form.ChkList_4_SV_TextBox.Name                 '檢查項目4  有   SV
            Dim usr_chkList_Q5_StdCont As String = JobMaker_Form.ChkList_5_std_RadioButton.Name       '檢查項目5  有   標準內容
            Dim usr_chkList_Q5_nStdCont As String = JobMaker_Form.ChkList_5_nstd_RadioButton.Name     '檢查項目5  有   工直內容
            Dim usr_chkList_Q6_YCont As String = JobMaker_Form.ChkList_6_yes_Content_TextBox.Name     '檢查項目6  有   檢驗項目
            Dim usr_chkList_Q7_YCont As String = JobMaker_Form.ChkList_7_yes1_content_TextBox.Name    '檢查項目7  有   文書No
            Dim usr_prgm_Reason As String = JobMaker_Form.PrmList_1_reason_TextBox.Name               '程式變更理由    
            Dim usr_prgm_2_testCont As String = JobMaker_Form.PrmList_2_COP_TextBox.Name              '程式變更        測試裝置
            Dim usr_prgm_2_CopCont As String = JobMaker_Form.PrmList_2_test_TextBox.Name              '程式變更理由     控制盤 
            Dim usr_prgm_2_TowerCont As String = JobMaker_Form.PrmList_2_tower_TextBox.Name           '程式變更理由     研修測試塔
            Dim usr_prgm_2_OtherCont As String = JobMaker_Form.PrmList_2_other_TextBox.Name           '程式變更理由     其他  
            Dim usr_prgm_3_otherCont As String = JobMaker_Form.PrmList_3_other_TextBox.Name           '程式變更理由     其他
            Dim usr_prgm_4_testCont As String = JobMaker_Form.PrmList_4_content12_TextBox.Name        '程式變更理由     測試內容 


            Dim usrChkList_arr, usrPrgm_arr As String()
            usrChkList_arr = {usr_ChkList_JobName, usr_ChkList_JobNo,
                              usr_nameManager_chkList_Q1, usr_chkList_Q1_YCont, usr_chkList_Q1_YResult,
                              usr_nameManager_chkList_Q2, usr_chkList_Q2_YCont, usr_chkList_Q2_YResult,
                              usr_nameManager_chkList_Q3, usr_chkList_Q3_YMan, usr_chkList_Q3_YCont, usr_chkList_Q3_YResult,
                              usr_chkList_Q4_MMIC, usr_chkList_Q4_SV,
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

            'Try
            '輸入相對應的check list值
            If JobMaker_Form.Use_ChkList_CheckBox.CheckState Then
                For Each i_chkListStr In usrChkList_arr
                    If i_chkListStr <> "" Then
                        'Try
                        Select Case i_chkListStr
                            '基本資料
                            Case usr_ChkList_JobName
                                excelWriteIn(JobMaker_Form.Basic_JobName_TextBox.Text,
                                             usr_ChkList_JobName,
                                             msExcel_workbook)
                            Case usr_ChkList_JobNo
                                excelWriteIn(JobMaker_Form.Basic_JobNoNew_TextBox.Text,
                                             usr_ChkList_JobNo,
                                             msExcel_workbook)
                            Case usr_nameManager_chkList_Q1
                                chkboxWriteIn(usr_nameManager_chkList_Q1,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_Q2
                                chkboxWriteIn(usr_nameManager_chkList_Q2,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_Q3
                                chkboxWriteIn(usr_nameManager_chkList_Q3,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_Q5_chkBoxState
                                chkboxWriteIn(usr_nameManager_chkList_Q5_chkBoxState,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_Q6
                                chkboxWriteIn(usr_nameManager_chkList_Q6,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_Q6_yes
                                chkboxWriteIn(usr_nameManager_chkList_Q6_yes,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_Q7
                                chkboxWriteIn(usr_nameManager_chkList_Q7,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_Q8
                                chkboxWriteIn(usr_nameManager_chkList_Q8,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_Q9
                                chkboxWriteIn(usr_nameManager_chkList_Q9,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_PA
                                chkboxWriteIn(usr_nameManager_chkList_PA,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_OS
                                chkboxWriteIn(usr_nameManager_chkList_OS,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_CFM
                                chkboxWriteIn(usr_nameManager_chkList_CFM,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            Case usr_nameManager_chkList_ELE
                                chkboxWriteIn(usr_nameManager_chkList_ELE,
                                                chkListP1_ShtName,
                                                msExcel_workbook)
                            'PA/OS/確認圖/電器的年月日
                            Case usr_chkList_PA_year
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_PaSheet_DateTimePicker.Value.Year.ToString(),
                                                            get_NameManager.ChkList_PA_Year,
                                                            JobMaker_Form.ChkList_PaSheet_CheckBox,
                                                            msExcel_workbook)                                                                '                          
                            Case usr_chkList_PA_month
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_PaSheet_DateTimePicker.Value.Month.ToString(),
                                                             get_NameManager.ChkList_PA_Month,
                                                             JobMaker_Form.ChkList_PaSheet_CheckBox,
                                                             msExcel_workbook)
                            Case usr_chkList_PA_date
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_PaSheet_DateTimePicker.Value.Day.ToString(),
                                                             get_NameManager.ChkList_PA_Day,
                                                             JobMaker_Form.ChkList_PaSheet_CheckBox,
                                                             msExcel_workbook)
                            Case usr_chkList_OS_year
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_OS_DateTimePicker.Value.Year.ToString(),
                                                             get_NameManager.ChkList_OS_Year,
                                                             JobMaker_Form.ChkList_OS_CheckBox,
                                                             msExcel_workbook)

                            Case usr_chkList_OS_month
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_OS_DateTimePicker.Value.Month.ToString(),
                                                             get_NameManager.ChkList_OS_Month,
                                                             JobMaker_Form.ChkList_OS_CheckBox,
                                                             msExcel_workbook)
                            Case usr_chkList_OS_date
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_OS_DateTimePicker.Value.Day.ToString(),
                                                             get_NameManager.ChkList_OS_Day,
                                                             JobMaker_Form.ChkList_OS_CheckBox,
                                                             msExcel_workbook)
                            Case usr_chkList_CFM_year
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Year.ToString(),
                                                             get_NameManager.ChkList_CFM_Year,
                                                             JobMaker_Form.ChkList_Confirm_CheckBox,
                                                             msExcel_workbook)
                            Case usr_chkList_CFM_month
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Month.ToString(),
                                                             get_NameManager.ChkList_CFM_Month,
                                                             JobMaker_Form.ChkList_Confirm_CheckBox,
                                                             msExcel_workbook)
                            Case usr_chkList_CFM_date
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_Confirm_DateTimePicker.Value.Day.ToString(),
                                                             get_NameManager.ChkList_CFM_Day,
                                                             JobMaker_Form.ChkList_Confirm_CheckBox,
                                                             msExcel_workbook)

                            Case usr_chkList_ELE_year
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_Elec_DateTimePicker.Value.Year.ToString(),
                                                             get_NameManager.ChkList_ELE_Year,
                                                             JobMaker_Form.ChkList_Elec_CheckBox,
                                                             msExcel_workbook)

                            Case usr_chkList_ELE_month
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_Elec_DateTimePicker.Value.Month.ToString(),
                                                             get_NameManager.ChkList_ELE_Month,
                                                             JobMaker_Form.ChkList_Elec_CheckBox,
                                                             msExcel_workbook)

                            Case usr_chkList_ELE_date
                                excelWriteIn_ForReverseState(JobMaker_Form.ChkList_Elec_DateTimePicker.Value.Day.ToString(),
                                                             get_NameManager.ChkList_ELE_Day,
                                                             JobMaker_Form.ChkList_Elec_CheckBox,
                                                             msExcel_workbook)


                            'Textbox內容寫入
                            Case usr_chkList_Q1_YCont
                                'CheckList > 1.主式樣有無不清楚 > 討論內容
                                excelWriteIn(JobMaker_Form.ChkList_1_yes_Content_TextBox.Text,
                                             get_NameManager.ChkList_Q1Yes_Content,
                                             JobMaker_Form.ChkList_1_yes_RadioButton,
                                             msExcel_workbook)
                            Case usr_chkList_Q1_YResult
                                'CheckList > 1.主式樣有無不清楚 > 結果
                                excelWriteIn(JobMaker_Form.ChkList_1_yes_result_TextBox.Text,
                                             get_NameManager.ChkList_Q1Yes_Result,
                                             JobMaker_Form.ChkList_1_yes_RadioButton,
                                             msExcel_workbook)
                            Case usr_chkList_Q2_YCont
                                'CheckList > 2.有沒有發生問題 > 指出內容
                                excelWriteIn(JobMaker_Form.ChkList_2_yes_Content_TextBox.Text,
                                             get_NameManager.ChkList_Q2Yes_Content,
                                             JobMaker_Form.ChkList_2_yes_RadioButton,
                                             msExcel_workbook)
                            Case usr_chkList_Q2_YResult
                                'CheckList > 2.有沒有發生問題 > 結果
                                excelWriteIn(JobMaker_Form.ChkList_2_yes_Result_TextBox.Text,
                                             get_NameManager.ChkList_Q2Yes_Result,
                                             JobMaker_Form.ChkList_2_yes_RadioButton,
                                             msExcel_workbook)
                            Case usr_chkList_Q3_YMan
                                'CheckList > 3.電氣圖有沒有不清楚 > 討論者
                                excelWriteIn(JobMaker_Form.ChkList_3_yes_Man_TextBox.Text,
                                             get_NameManager.ChkList_Q3Yes_Man,
                                             JobMaker_Form.ChkList_3_yes_RadioButton,
                                             msExcel_workbook)
                            Case usr_chkList_Q3_YCont
                                'CheckList > 3.電氣圖有沒有不清楚 > 內容
                                excelWriteIn(JobMaker_Form.ChkList_3_yes_Content_TextBox.Text,
                                             get_NameManager.ChkList_Q3Yes_Content,
                                             JobMaker_Form.ChkList_3_yes_RadioButton,
                                             msExcel_workbook)
                            Case usr_chkList_Q3_YResult
                                'CheckList > 3.電氣圖有沒有不清楚 > 結論
                                excelWriteIn(JobMaker_Form.ChkList_3_yes_Result_TextBox.Text,
                                             get_NameManager.ChkList_Q3Yes_Result,
                                             JobMaker_Form.ChkList_3_yes_RadioButton,
                                             msExcel_workbook)
                            Case usr_chkList_Q4_MMIC
                                'CheckList > 4.MMIC > 內容
                                Dim dyCrtl As New DynamicControlName
                                dynamicControl_writeInExcel_CheckList_Prgm(JobMaker_Form.MMIC_MR_NumericUpDown.Value,
                                                                           get_NameManager.ChkList_Q4MMIC,
                                                                           get_NameManager.ChkList_Q4MmicBase,
                                                                           JobMaker_Form.MMIC_MR_Panel,
                                                                           dyCrtl.mmicBase_CarNo, dyCrtl.mmicBase_ObjName, dyCrtl.mmicBase_ObjNameBase,
                                                                           msExcel_workbook)
                            Case usr_chkList_Q4_SV
                                'CheckList > 4.SV > 內容
                                Dim dyCrtl As New DynamicControlName
                                dynamicControl_writeInExcel_CheckList_Prgm(JobMaker_Form.MMIC_SV_NumericUpDown.Value,
                                                                           get_NameManager.ChkList_Q4SV,
                                                                           get_NameManager.ChkList_Q4SVmicBase,
                                                                           JobMaker_Form.MMIC_SV_Panel,
                                                                           dyCrtl.svBase_CarNo, dyCrtl.svBase_ObjName, dyCrtl.svBase_ObjNameBase,
                                                                           msExcel_workbook)
                            Case usr_chkList_Q5_StdCont
                                'CheckList > 5.VONIC > 標準內容
                                If JobMaker_Form.ChkList_5_std_RadioButton.Checked Then
                                    dynamicControl_writeInExcel_CheckList_VD10(JobMaker_Form.MMIC_VD10_NumericUpDown.Value,
                                                                               get_NameManager.ChkList_Q5Std_Content,
                                                                               JobMaker_Form.MMIC_VD10_Panel,
                                                                               msExcel_workbook)
                                End If
                            Case usr_chkList_Q5_nStdCont
                                'CheckList > 5.VONIC > 工直內容
                                If JobMaker_Form.ChkList_5_nstd_RadioButton.Checked Then

                                    dynamicControl_writeInExcel_CheckList_VD10(JobMaker_Form.MMIC_VD10_NumericUpDown.Value,
                                                                get_NameManager.ChkList_Q5nStd_Content,
                                                                JobMaker_Form.MMIC_VD10_Panel,
                                                                msExcel_workbook)
                                End If
                                'CheckList > 6.執行動作確認 > 檢驗項目內容
                                excelWriteIn(JobMaker_Form.ChkList_6_yes_Content_TextBox.Text,
                                             get_NameManager.ChkList_Q6Yes_Content,
                                             JobMaker_Form.ChkList_6_yes_RadioButton,
                                             msExcel_workbook)

                            Case usr_chkList_Q7_YCont
                                'CheckList > 7.參考資料 > 文書NO
                                excelWriteIn(JobMaker_Form.ChkList_7_yes1_content_TextBox.Text,
                                             get_NameManager.ChkList_Q7Yes_Content,
                                             JobMaker_Form.ChkList_7_yes_RadioButton,
                                             msExcel_workbook)

                        End Select
                        'Catch ex As Exception
                        '    JobMaker_Form.ResultOutput_TextBox.Text += ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_chkListStr}>{vbCrLf}")
                        'End Try
                    End If
                Next
            Else
                    JobMaker_Form.ResultFailOutput_TextBox.Text = ($"<提醒> Check List 分頁未輸出，原因:分頁未打勾")
                    JobMaker_Form.JobMaker_TabControl.SelectedTab = JobMaker_Form.CheckList2_TabPage
                    Dim basic_result As DialogResult = MsgBox(($"「{JobMaker_Form.CheckList2_TabPage.Text}」未使用是否重來?"), vbYesNo)
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
                        'Try
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
                                             get_NameManager.ChkList_Prgm_1_reason,
                                             msExcel_workbook)

                            Case usr_prgm_2_testCont
                                'Check List > 程式變更 > 2-1測試裝置
                                excelWriteIn(JobMaker_Form.PrmList_2_COP_TextBox.Text,
                                             get_NameManager.ChkList_Prgm_2_Test_Content,
                                             JobMaker_Form.PrmList_2_test_CheckBox,
                                             msExcel_workbook)
                            Case usr_prgm_2_CopCont
                                'Check List > 程式變更 > 2-2控制盤
                                excelWriteIn(JobMaker_Form.PrmList_2_test_TextBox.Text,
                                             get_NameManager.ChkList_Prgm_2_COP_Content,
                                             JobMaker_Form.PrmList_2_COP_CheckBox,
                                             msExcel_workbook)
                            Case usr_prgm_2_TowerCont
                                'Check List > 程式變更 > 2-3研修測試塔
                                excelWriteIn(JobMaker_Form.PrmList_2_tower_TextBox.Text,
                                             get_NameManager.ChkList_Prgm_2_Tower_Content,
                                             JobMaker_Form.PrmList_2_Tower_CheckBox,
                                             msExcel_workbook)
                            Case usr_prgm_2_OtherCont
                                'Check List > 程式變更 > 2-4 其他
                                excelWriteIn(JobMaker_Form.PrmList_2_other_TextBox.Text,
                                             get_NameManager.ChkList_Prgm_2_Other_Content,
                                             JobMaker_Form.PrmList_2_Other_CheckBox,
                                             msExcel_workbook)
                            Case usr_prgm_3_otherCont
                                'Check List > 程式變更 > 3-1 其他
                                excelWriteIn(JobMaker_Form.PrmList_3_other_TextBox.Text,
                                             get_NameManager.ChkList_Prgm_3_OtherContent,
                                             JobMaker_Form.PrmList_3_other_Checkbox,
                                             msExcel_workbook)
                            Case usr_prgm_4_testCont
                                'Check List > 程式變更 > 4 測試內容
                                excelWriteIn(JobMaker_Form.PrmList_4_content12_TextBox.Text,
                                             get_NameManager.ChkList_Prgm_4_TestContent,
                                             msExcel_workbook)
                        End Select
                        'Catch ex As Exception
                        '    JobMaker_Form.ResultOutput_TextBox.Text += ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_prgmStr}>{vbCrLf}")
                        'End Try
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
                'Catch ex As Exception
                '    MsgBox(ex.Message)
                'End Try

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
        Dim prk_Row, prk_Col, temp_prk_Row As Integer

        '取得 名稱管理員specName 的Row
        Dim startCell_Row As Integer =
            msExcel_workbook.Names.Item(specName).RefersToRange.Row '號機名是第n行

        '取得 名稱管理員specName 的Col
        Dim startCell_Col As Integer =
            msExcel_workbook.Names.Item(specName).RefersToRange.Column '號機名是第n行
        '取得 目前使用的worksheet名稱
        Dim startWorksheet_name As Integer = msExcel_workbook.Names.Item(specName).RefersToRange.Worksheet.Name

        '取得 名稱管理員specName Range的頭例如A4的4
        Dim startRange_Row As Integer = startCell_Row
        Dim startRange_Col As Integer =
            getMathOnExcel.convertColumn_fromIntToString(startCell_Col)
        '取得該合併儲存格的數量
        Dim merge_num As Integer =
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

    Private Sub dynamicControl_writeInExcel_CheckList_Prgm(mNumericUpDown_num As Integer,
                                                           insertRow_specName As String, notInsertRow_specName As String,
                                                           mPanel As Control,
                                                           CarNo As String, ObjName As String, ObjNameBase As String,
                                                           msExcel_workbook As Excel.Workbook)
        Dim startWorksheet_name As String = Nothing
        Dim merge_num As Integer = Nothing

        insertRow(mNumericUpDown_num, insertRow_specName, msExcel_workbook, startWorksheet_name, merge_num)

        '檢查Panel中有幾個控制項就跑幾次
        Dim dyCtrl As New DynamicControlName
        Dim outputText(mNumericUpDown_num - 1) As String
        Dim outputText_Base(mNumericUpDown_num - 1) As String
        Dim prk_insertCol As Integer = msExcel_workbook.Names.Item(insertRow_specName).RefersToRange.Column '行
        Dim prk_insertRow As Integer = msExcel_workbook.Names.Item(insertRow_specName).RefersToRange.Row '列
        Dim prk_Col As Integer = msExcel_workbook.Names.Item(notInsertRow_specName).RefersToRange.Column '行
        Dim prk_Row As Integer = msExcel_workbook.Names.Item(notInsertRow_specName).RefersToRange.Row '列

        For Each tempCtrl As Control In mPanel.Controls '填入電梯的相關資訊
            For lift_i = 1 To mNumericUpDown_num
                '檢查控制項名稱是否符合需求的(dyCtrl_Array)
                If tempCtrl.Name = $"{CarNo}_{lift_i}" Then
                    outputText(lift_i - 1) += $"{tempCtrl.Text}:"
                ElseIf tempCtrl.Name = $"{ObjName}_{lift_i}" Then
                    outputText(lift_i - 1) += tempCtrl.Text
                ElseIf tempCtrl.Name = $"{ObjNameBase}_{lift_i}" Then
                    outputText_Base(lift_i - 1) = $"(BASE:{tempCtrl.Text})"
                End If
            Next
        Next
        '取得欄、行，每執行完一次就會更新"行"的值 --------------------------------------------------
        For lift_i = 1 To mNumericUpDown_num
            msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_insertRow, prk_insertCol).Value = outputText(lift_i - 1)
            msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_Row, prk_Col).Value = outputText_Base(lift_i - 1)
            prk_insertRow += merge_num
            prk_Row += merge_num
        Next
        '-------------------------------------------------- 取得欄、行，每執行完一次就會更新"行"的值 
    End Sub
    Private Sub dynamicControl_writeInExcel_CheckList_VD10(mNumericUpDown_num As Integer, specName As String,
                                                             mPanel As Control,
                                                             msExcel_workbook As Excel.Workbook)
        Dim startWorksheet_name As String = Nothing
        Dim merge_num As Integer = Nothing
        insertRow(mNumericUpDown_num, specName, msExcel_workbook, startWorksheet_name, merge_num)

        '檢查Panel中有幾個控制項就跑幾次
        Dim dyCtrl As New DynamicControlName
        Dim outputText(mNumericUpDown_num - 1) As String
        Dim prk_Col As Integer = msExcel_workbook.Names.Item(specName).RefersToRange.Column '行
        Dim prk_Row As Integer = msExcel_workbook.Names.Item(specName).RefersToRange.Row '列

        For Each tempCtrl As Control In mPanel.Controls '填入電梯的相關資訊
            For lift_i = 1 To mNumericUpDown_num
                '檢查控制項名稱是否符合需求的(dyCtrl_Array)
                If tempCtrl.Name = $"{dyCtrl.vd10Base_CarNo}_{lift_i}" Then
                    outputText(lift_i - 1) += $"{tempCtrl.Text}:"
                ElseIf tempCtrl.Name = $"{dyCtrl.vd10Base_ObjName}_{lift_i}" Then
                    outputText(lift_i - 1) += tempCtrl.Text
                End If
            Next
        Next
        '取得欄、行，每執行完一次就會更新"行"的值 --------------------------------------------------
        For lift_i = 1 To mNumericUpDown_num
            msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_Row, prk_Col).Value = outputText(lift_i - 1)
            prk_Row += merge_num
        Next
        '-------------------------------------------------- 取得欄、行，每執行完一次就會更新"行"的值 
    End Sub

    Private Shared Sub insertRow(mNumericUpDown_num As Integer, specName As String,
                                 msExcel_workbook As Excel.Workbook,
                                 ByRef startWorksheet_name As String, ByRef merge_num As Integer)
        '取得 名稱管理員specName 的Row
        Dim startCell_Row As Integer =
            msExcel_workbook.Names.Item(specName).RefersToRange.Row '號機名是第n行
        '取得 名稱管理員specName 的Col
        Dim startCell_Col As Integer =
            msExcel_workbook.Names.Item(specName).RefersToRange.Column '號機名是第n行
        '取得 目前使用的worksheet名稱
        startWorksheet_name = msExcel_workbook.Names.Item(specName).RefersToRange.Worksheet.Name

        '取得 名稱管理員specName Range的頭例如A4的4
        Dim startRange_Row As Integer = startCell_Row
        Dim startRange_Col As Integer = getMathOnExcel.convertColumn_fromIntToString(startCell_Col)
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
                                                        specName_Array As ArrayList,
                                                        mPanel As Control,
                                                        dyCtrl_ArrayCount As Integer, dyCtrl_Array As Array,
                                                        msExcel_workbook As Excel.Workbook)

        Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData

        Dim prk_Row, prk_Col, temp_prk_Row As Integer

        '取得 名稱管理員specName 的Row
        Dim startCell_Row As Integer =
            getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, specName)
        '取得 名稱管理員specName 的Col
        Dim startCell_Col As Integer =
            getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, specName)
        '取得 目前使用WorkSheet的名稱
        Dim startWorksheet_name As String =
            getMathOnExcel.getWorksheetName_fromNameManager(msExcel_workbook, specName)

        '取得 名稱管理員specName Range的頭例如A4的4
        Dim startRange_Row As Integer = startCell_Row
        '取得 名稱管理員specName Range的尾例如A4的A
        Dim startRange_Col As String =
            getMathOnExcel.convertColumn_fromIntToString(startCell_Col)

        '取得該合併儲存格的數量
        Dim merge_num As Integer =
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

                        If tempCtrl.Name = $"{JobMaker_Form.Spec_TopFL_ComboBox.Name}_{lift_i}" Then
                            For Each realFL As Control In mPanel.Controls
                                If realFL.Name = $"{JobMaker_Form.Spec_TopFL_Real_ComboBox.Name}_{lift_i}" Then
                                    tempCtrlText = $"{tempCtrl.Text} {realFL.Text}"
                                End If
                            Next
                        ElseIf tempCtrl.Name = $"{JobMaker_Form.Spec_BtmFL_ComboBox.Name}_{lift_i}" Then
                            For Each realFL As Control In mPanel.Controls
                                If realFL.Name = $"{JobMaker_Form.Spec_BtmFL_Real_ComboBox.Name}_{lift_i}" Then
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
        mNumericUpDown1_num = CInt(mNumericUpDown1.Value)
        mNumericUpDown2_num = CInt(mNumericUpDown2.Value)
        If mNumericUpDown1_num > mNumericUpDown2_num Then
            mNumeric_Max = mNumericUpDown1_num
        ElseIf mNumericUpDown1_num < mNumericUpDown2_num Then
            mNumeric_Max = mNumericUpDown2_num
        Else
            mNumeric_Max = mNumericUpDown1_num
        End If
        '----------------------------------- 針對flashRom 與 EEPROM 比較 取最大值 



        Dim prk_Row, prk_Col, temp_prk_Row As Integer

        '取得 名稱管理員specName 的Row
        Dim startCell_Row As Integer =
            getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, specName)
        '取得 名稱管理員specName 的Col
        Dim startCell_Col As Integer =
            getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, specName)
        '取得 目前WorkSheet名稱
        Dim startWorksheet_name As String =
            getMathOnExcel.getWorksheetName_fromNameManager(msExcel_workbook, specName)

        '取得 名稱管理員specName Range的頭例如A4的4
        Dim startRange_Row As Integer = startCell_Row
        '取得 名稱管理員specName Range的尾例如A4的A
        Dim startRange_Col As Integer =
           getMathOnExcel.convertColumn_fromIntToString(startCell_Col)

        '取得該合併儲存格的數量
        Dim merge_num As Integer =
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

        For Each tempCtrl1 As Control In mPanel1.Controls '填入電梯的相關資訊
            For lift_i As Integer = 1 To mNumericUpDown1_num
                For lift_j As Integer = 1 To dyCtrl1_ArrayCount
                    Console.WriteLine(tempCtrl1.Name)
                    If tempCtrl1.Name = $"{dyCtrl1_Array(lift_j - 1)}_{lift_i}" Then
                        prk_Col = msExcel_workbook.Names.Item(specName1_Array(lift_j - 1)).RefersToRange.Column '行
                        prk_Row = msExcel_workbook.Names.Item(specName1_Array(lift_j - 1)).RefersToRange.Row '列
                        prk_Row += lift_i * merge_num

                        msExcel_workbook.Worksheets(startWorksheet_name).Cells(prk_Row, prk_Col).Value = tempCtrl1.Text
                        Console.WriteLine($"({prk_Row},{prk_Col})={tempCtrl1.Text}")
                        prk_Row = temp_prk_Row
                    End If
                Next
            Next
        Next

        prk_Row = 0
        prk_Col = 0

        For Each tempCtrl2 As Control In mPanel2.Controls
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
                Dim usrInput_arr As New ArrayList
                Dim spec_car_name As String = get_NameManager.SPEC_CAR_NAME
                usrInput_arr.Add(spec_car_name)

                Dim spec_car_no As String = get_NameManager.SPEC_CAR_NO
                usrInput_arr.Add(spec_car_no)

                Dim spec_car_ope As String = get_NameManager.SPEC_CAR_OPE
                usrInput_arr.Add(spec_car_ope)

                Dim spec_car_topfl As String = get_NameManager.SPEC_CAR_TOPFL
                usrInput_arr.Add(spec_car_topfl)

                Dim spec_car_btmfl As String = get_NameManager.SPEC_CAR_BTMFL
                usrInput_arr.Add(spec_car_btmfl)

                Dim spec_car_stop As String = get_NameManager.SPEC_CAR_STOP
                usrInput_arr.Add(spec_car_stop)

                Dim spec_car_speed As String = get_NameManager.SPEC_CAR_SPEED
                usrInput_arr.Add(spec_car_speed)

                Dim spec_car_flname As String = get_NameManager.SPEC_CAR_FLNAME
                usrInput_arr.Add(spec_car_flname)

                'Spec 基本
                Dim dyCtrlName As DynamicControlName = New DynamicControlName
                dyCtrlName.JobMaker_LiftInfo()
                dynamicControl_writeInExcel_SpecBasic(JobMaker_Form.Spec_LiftNum_NumericUpDown.Value,
                                                      get_NameManager.SPEC_CAR_NAME,
                                                      usrInput_arr,
                                                      JobMaker_Form.SpecBasic_LiftItem_Dynamic_Panel,
                                                      dyCtrlName.JobMaker_LiftInfoName_Array.Count,
                                                      dyCtrlName.JobMaker_LiftInfoName_Array,
                                                      msExcel_workbook)


                Dim usrInput_arr2 As New ArrayList

                Dim spec_car_machine_type As String = get_NameManager.SPEC_CAR_MACHINE_TYPE
                usrInput_arr2.Add(spec_car_machine_type)
                Dim spec_car_control_way As String = get_NameManager.SPEC_CAR_CONTROL_WAY
                usrInput_arr2.Add(spec_car_control_way)

                For Each item In usrInput_arr2
                    If item <> "" Then
                        Try
                            Select Case item
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
                                    Dim JM_PURPOSE As String() = {get_NameManager.SPEC_CAR_PURPOSE}

                                    dynamicControl_writeInExcel_MMIC(JobMaker_Form.Spec_MachineType_NumericUpDown, JobMaker_Form.Spec_Purpose_NumericUpDown,
                                                                     spec_car_control_way,
                                                                     JM_CONTROL_WAY, JM_PURPOSE,
                                                                     JobMaker_Form.Spec_ControlWay_Panel, JobMaker_Form.Spec_Purpose_Panel,
                                                                     {dyCtrlName.Spec_ControlWay_ComboBox}.Count, {dyCtrlName.Spec_ControlWay_ComboBox},
                                                                     {dyCtrlName.Spec_Purpose_ComboBox}.Count, {dyCtrlName.Spec_Purpose_ComboBox},
                                                                     msExcel_workbook)
                            End Select
                        Catch ex As Exception
                            JobMaker_Form.ResultFailOutput_TextBox.Text +=
                                ($"<{JobMaker_Form.JMFileCho_Spec_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{item}>{vbCrLf}")
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
            MsgBox($"Spec_SPEC_Basic funciton error : {ex.Message}")
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

            Dim usrInput_arr As New ArrayList
            Dim spec_operation_type As String = get_NameManager.SPEC_OPERATION_TYPE '操作方式
            usrInput_arr.Add(spec_operation_type)

            Dim usr_Spec_MachineType As String = JobMaker_Form.Spec_Base_ComboBox.Name
            usrInput_arr.Add(usr_Spec_MachineType)

            Dim usr_Spec_Auto_DR As String = JobMaker_Form.Spec_DRAuto_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Auto_DR)

            Dim usr_Spec_Cancell_call As String = JobMaker_Form.Spec_CancellCall_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Cancell_call)

            Dim usr_Spec_Lamp_Chk As String = JobMaker_Form.Spec_LampChk_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Lamp_Chk)

            Dim usr_Spec_Cancell_Behind As String = JobMaker_Form.Spec_CancellBehind_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Cancell_Behind)

            Dim usr_Spec_Auto_Fan As String = JobMaker_Form.Spec_AutoFan_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Auto_Fan)

            Dim usr_Spec_CC_Cancel As String = JobMaker_Form.Spec_CCCancell_ComboBox.Name
            usrInput_arr.Add(usr_Spec_CC_Cancel)

            Dim usr_Spec_Auto_Pass As String = JobMaker_Form.Spec_AutoPass_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Auto_Pass)

            Dim usr_Spec_Operation As String = JobMaker_Form.Spec_Operation_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Operation)

            Dim usr_Spec_Install_Ope As String = JobMaker_Form.Spec_install_ope_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Install_Ope)

            Dim usr_Spec_Indep_Ope As String = JobMaker_Form.Spec_Indep_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Indep_Ope)

            Dim usr_Spec_UCMP As String = JobMaker_Form.Spec_UCMP_ComboBox.Name
            usrInput_arr.Add(usr_Spec_UCMP)

            Dim usr_Spec_HIN_CPI As String = JobMaker_Form.Spec_HinCpi_ComboBox.Name
            usrInput_arr.Add(usr_Spec_HIN_CPI)

            Dim usr_Spec_Fire_Ope As String = JobMaker_Form.Spec_Fire_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Fire_Ope)

            Dim usr_Spec_Fireman As String = JobMaker_Form.Spec_Fireman_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Fireman)

            Dim usr_Spec_Parking As String = JobMaker_Form.Spec_Parking_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Parking)

            Dim usr_Spec_Seismic As String = JobMaker_Form.Spec_Seismic_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Seismic)

            Dim usr_Spec_CPI As String = JobMaker_Form.Spec_CPI_ComboBox.Name
            usrInput_arr.Add(usr_Spec_CPI)

            Dim usr_Spec_Car_Gong As String = JobMaker_Form.Spec_CarGong_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Car_Gong)

            Dim usr_Spec_Hall_Gong As String = JobMaker_Form.Spec_HallGong_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Hall_Gong)

            Dim usr_Spec_HPI As String = JobMaker_Form.Spec_HPIMsg_ComboBox.Name
            usrInput_arr.Add(usr_Spec_HPI)

            Dim usr_Spec_Dr_Hold As String = JobMaker_Form.Spec_DrHold_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Dr_Hold)

            Dim usr_Spec_CRD As String = JobMaker_Form.Spec_CRD_ComboBox.Name
            usrInput_arr.Add(usr_Spec_CRD)

            Dim usr_Spec_Emer_Power As String = JobMaker_Form.Spec_Emer_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Emer_Power)

            Dim usr_Spec_Landic As String = JobMaker_Form.Spec_Landic_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Landic)

            Dim usr_Spec_MLF_Return As String = JobMaker_Form.Spec_MFLReturn_ComboBox.Name
            usrInput_arr.Add(usr_Spec_MLF_Return)

            Dim usr_Spec_Vonic As String = JobMaker_Form.Spec_Vonic_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Vonic)

            Dim usr_Spec_WCOB As String = JobMaker_Form.Spec_WCOB_ComboBox.Name
            usrInput_arr.Add(usr_Spec_WCOB)

            Dim usr_Spec_Elvic As String = JobMaker_Form.Spec_Elvic_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Elvic)

            Dim usr_Spec_HLL As String = JobMaker_Form.Spec_HLL_ComboBox.Name
            usrInput_arr.Add(usr_Spec_HLL)

            Dim usr_Spec_ATT As String = JobMaker_Form.Spec_ATT_ComboBox.Name
            usrInput_arr.Add(usr_Spec_ATT)

            Dim usr_Spec_Flood As String = JobMaker_Form.Spec_Flood_ComboBox.Name
            usrInput_arr.Add(usr_Spec_Flood)

            Dim usr_Spec_LS1M As String = JobMaker_Form.Spec_LS1M_ComboBox.Name
            usrInput_arr.Add(usr_Spec_LS1M)

            Dim usr_Spec_PRU As String = JobMaker_Form.Spec_PRU_ComboBox.Name
            usrInput_arr.Add(usr_Spec_PRU)

            Dim usr_Spec_FrontRear_DR As String = JobMaker_Form.Spec_FrontRearDr_ComboBox.Name
            usrInput_arr.Add(usr_Spec_FrontRear_DR)

            Dim usr_Spec_OpeSw As String = JobMaker_Form.Spec_OpeSw_ComboBox.Name
            usrInput_arr.Add(usr_Spec_OpeSw)

            Dim usr_Spec_WTB As String = JobMaker_Form.Spec_WTB_ComboBox.Name
            usrInput_arr.Add(usr_Spec_WTB)

            Dim with_val As String =
                msExcel_workbook.Names.Item(get_NameManager.SetTable_RESULT_WITH).RefersToRange.Value '取得 有 內的文字內容
            Dim without_val As String =
                msExcel_workbook.Names.Item(get_NameManager.SetTable_RESULT_WITHOUT).RefersToRange.Value '取得 無 內的文字內容
            Dim no_val As String =
                msExcel_workbook.Names.Item(get_NameManager.SetTable_NO).RefersToRange.Value '取得 訊號NO 內的文字內容
            Dim nc_val As String =
                msExcel_workbook.Names.Item(get_NameManager.SetTable_NC).RefersToRange.Value '取得 訊號NC 內的文字內容
            For Each item In usrInput_arr
                If item <> "" Then
                    'Try
                    Select Case item
                            '操作方式 ------------------------------------------------------------------------------------------------------
                        Case spec_operation_type
                            If JobMaker_Form.Spec_LiftNum_NumericUpDown.Value >= 1 Then
                                Dim mdir As New Dictionary(Of String, String)
                                For Each items As Control In JobMaker_Form.SpecBasic_LiftItem_Dynamic_Panel.Controls
                                    For i As Integer = 1 To LiftNum
                                        If items.Name = $"{JobMaker_Form.Spec_Control_ComboBox.Name}_{i}" Then
                                            If mdir.ContainsKey(items.Text) Then
                                                For Each items_2 As Control In JobMaker_Form.SpecBasic_LiftItem_Dynamic_Panel.Controls
                                                    If items_2.Name = $"{JobMaker_Form.Spec_LiftName_TextBox.Name}_{i}" Then
                                                        mdir(items.Text) = mdir.Item(items.Text) & $",{items_2.Text}"
                                                        Exit For
                                                    End If
                                                Next
                                            Else
                                                For Each items_2 As Control In JobMaker_Form.SpecBasic_LiftItem_Dynamic_Panel.Controls
                                                    If items_2.Name = $"{JobMaker_Form.Spec_LiftName_TextBox.Name}_{i}" Then
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
                                excelWriteIn(outputText,
                                             get_NameManager.SPEC_OPERATION_TYPE,
                                             msExcel_workbook)
                            End If
                            '------------------------------------------------------------------------------------------------------ 操作方式 
                            ' 開門時限自動調節 ------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Auto_DR
                            excelWriteIn(JobMaker_Form.Spec_DRAuto_ComboBox.Text,
                                         get_NameManager.SPEC_AUTO_DR,
                                         msExcel_workbook)
                            If JobMaker_Form.Spec_DRAuto_ComboBox.Text = get_NameManager.TB_O Then
                                Dim spec_auto_dr_photoeye As String = get_NameManager.SPEC_AUTO_DR_PHOTOEYE
                                Dim spec_auto_dr_safety As String = get_NameManager.SPEC_AUTO_DR_SAFETY

                                Dim pho_val As String =
                                    msExcel_workbook.Names.Item(spec_auto_dr_photoeye
                                                                ).RefersToRange.Value '取得 光電裝置 內的文字內容
                                Dim safety_val As String =
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
                                         get_NameManager.SPEC_CANCELL_CALL,
                                         msExcel_workbook)
                            If JobMaker_Form.Spec_CancellCall_ComboBox.Text = get_NameManager.TB_O Then
                                Dim spec_cancell_call_scob As String =
                                    get_NameManager.SPEC_CANCELL_CALL_SCOB
                                Dim spec_cancell_call_six As String =
                                    get_NameManager.SPEC_CANCELL_CALL_SIX

                                Dim scob_val As String =
                                    msExcel_workbook.Names.Item(spec_cancell_call_scob
                                                                ).RefersToRange.Value '取得 SCOB 內的文字內容
                                Dim six_val As String =
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
                                         get_NameManager.SPEC_CANCELL_BEHIND,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 逆呼無效

                            ' 燈點檢模式 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Lamp_Chk
                            excelWriteIn(JobMaker_Form.Spec_LampChk_ComboBox.Text,
                                         get_NameManager.SPEC_LAMP_CHK,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 燈點檢模式


                            ' 風扇連動 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Auto_Fan
                            excelWriteIn(JobMaker_Form.Spec_AutoFan_ComboBox.Text,
                                         get_NameManager.SPEC_AUTO_FAN,
                                         msExcel_workbook)
                            If JobMaker_Form.Spec_ION_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                '離子除菌
                                Dim spec_auto_fan_ion As String =
                                    get_NameManager.SPEC_AUTO_FAN_ION

                                msExcel_workbook.Names.Item(spec_auto_fan_ion).RefersToRange.Cells.Font.Strikethrough = True

                                Dim ion_row As Integer =
                                    getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, spec_auto_fan_ion)
                                Dim ion_col As Integer =
                                    getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, spec_auto_fan_ion)

                                Dim spec_shtName As String =
                                    getMathOnExcel.getWorksheetName_fromNameManager(msExcel_workbook, spec_auto_fan_ion)
                                For i = 0 To 3
                                    msExcel_workbook.Worksheets(spec_shtName).Cells(ion_row + i, ion_col).font.Strikethrough = True
                                Next
                            End If

                            '重要設定 ion
                            Dim ion_val As String = ""
                            If JobMaker_Form.Spec_AutoFan_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                ion_val = get_NameManager.TB_WITHOUT
                            Else
                                If JobMaker_Form.Spec_ION_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                    ion_val = get_NameManager.TB_WITH
                                    msExcel_workbook.Names.Item(get_NameManager.IMPORTANT_FAN_CONTENT
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                ElseIf JobMaker_Form.Spec_ION_ComboBox.Text = get_NameManager.TB_WITH Then
                                    ion_val = get_NameManager.TB_WITH & "(ION)"
                                End If
                            End If

                            excelWriteIn(ion_val,
                                         get_NameManager.IMPORTANT_FAN,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 風扇連動


                            ' 車廂呼叫取消機能 ----------------------------------------------------------------------------------------------------
                        Case usr_Spec_CC_Cancel
                            excelWriteIn(JobMaker_Form.Spec_CCCancell_ComboBox.Text,
                                         get_NameManager.SPEC_CC_CANCEL,
                                         msExcel_workbook)
                            '---------------------------------------------------------------------------------------------------- 車廂呼叫取消機能

                            ' 自動滿員通過 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Auto_Pass
                            excelWriteIn(JobMaker_Form.Spec_AutoPass_ComboBox.Text,
                                         get_NameManager.SPEC_AUTO_PASS,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 自動滿員通過

                            ' 拒付運轉 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Install_Ope
                            excelWriteIn(JobMaker_Form.Spec_install_ope_ComboBox.Text,
                                         get_NameManager.SPEC_INSTALL_OPE,
                                         msExcel_workbook)'------------------------------------------------------------------------------------------------------ 拒付運轉

                            ' 專用運轉 -----------------------------------------------------------------------------------------------------
                        Case usr_Spec_Indep_Ope
                            excelWriteIn(JobMaker_Form.Spec_Indep_ComboBox.Text,
                                         get_NameManager.SPEC_INDEP_OPE,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 專用運轉

                            ' 戶開行走保護裝置 --------------------------------------------------------------------------------------------------
                        Case usr_Spec_UCMP
                            excelWriteIn(JobMaker_Form.Spec_UCMP_ComboBox.Text,
                                         get_NameManager.SPEC_UCMP,
                                         msExcel_workbook)
                            '--------------------------------------------------------------------------------------------------- 戶開行走保護裝置

                            ' HIN CPI --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_HIN_CPI
                            excelWriteIn(JobMaker_Form.Spec_HinCpi_ComboBox.Text,
                                         get_NameManager.SPEC_HIN_CPI,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ HIN CPI

                            ' 火災 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Fire_Ope
                            excelWriteIn(JobMaker_Form.Spec_Fire_ComboBox.Text,
                                         get_NameManager.SPEC_FIRE_OPE,
                                         msExcel_workbook)

                            '避難階
                            msExcel_workbook.Names.Item(get_NameManager.SetTable_ESCAPE_FL
                                                        ).RefersToRange.Cells.Value = JobMaker_Form.Spec_EscapeFL_TextBox.Text

                            If JobMaker_Form.Spec_Fire_ComboBox.Text = get_NameManager.TB_O Then
                                Dim spec_fire_ope_signal As String =
                                    get_NameManager.SPEC_FIRE_OPE_SIGNAL

                                Dim signal_val As String =
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
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_FIRE_ONLY).RefersToRange.Cells.Value =
                                        $"(Only {JobMaker_Form.Spec_Fire_Only_TextBox.Text})"
                                End If
                            End If
                            '------------------------------------------------------------------------------------------------------ 火災

                            ' 消防梯 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Fireman
                            excelWriteIn(JobMaker_Form.Spec_Fireman_ComboBox.Text,
                                         get_NameManager.SPEC_FIREMAN,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_Fireman_ComboBox.Text = get_NameManager.TB_O And
                               JobMaker_Form.Spec_Fireman_Only_CheckBox.Checked Then
                                msExcel_workbook.Names.Item(get_NameManager.SetTable_ESCAPE_FL_ONLY).RefersToRange.Cells.Value =
                                    $"(Only {JobMaker_Form.Spec_Fireman_Only_TextBox.Text})"
                            End If

                            '-----------------------------------------------------------------------------------------------------  消防梯

                            ' 停車階運轉 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Parking
                            excelWriteIn(JobMaker_Form.Spec_Parking_ComboBox.Text,
                                         get_NameManager.SPEC_PARKING,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_Parking_ComboBox.Text = get_NameManager.TB_O Then
                                Dim spec_pk_cmd1 As String = get_NameManager.SPEC_PK_CMD1
                                Dim spec_pk_cmd2 As String = get_NameManager.SPEC_PK_CMD2

                                Dim cmd1 As String =
                                    msExcel_workbook.Names.Item(spec_pk_cmd1).RefersToRange.Value '取得 cmd 內的文字內容

                                Dim cmd2 As String =
                                    msExcel_workbook.Names.Item(spec_pk_cmd2).RefersToRange.Value '取得 cmd 內的文字內容

                                Dim elv_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_PK_ELVIC).RefersToRange.Value '取得 elvic 內的文字內容

                                Dim wtb_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_PK_WTB).RefersToRange.Value '取得 WTB 內的文字內容

                                Dim cob_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_PK_COB).RefersToRange.Value '取得 COB 內的文字內容

                                Dim dro_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_PK_DROPEN).RefersToRange.Value '取得 OPEN 內的文字內容

                                Dim hal_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_PK_SW).RefersToRange.Value '取得 SW 內的文字內容

                                Dim drc_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_PK_DRCLOSE).RefersToRange.Value '取得 CLOSE 內的文字內容
                                msExcel_workbook.Names.Item(get_NameManager.SetTable_PARKING_FL).RefersToRange.Cells.Value =
                                    JobMaker_Form.Spec_Parking_FL_TextBox.Text

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
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_PARKING_ONLY).RefersToRange.Cells.Value =
                                        $"(Only {JobMaker_Form.Spec_Parking_Only_TextBox.Text})"
                                End If
                            End If
                            '------------------------------------------------------------------------------------------------------ 停車階運轉

                            ' 地震管制運轉 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Seismic
                            excelWriteIn(JobMaker_Form.Spec_Seismic_ComboBox.Text,
                                         get_NameManager.SPEC_SEISMIC,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_Seismic_ComboBox.Text = get_NameManager.TB_O Then
                                '地震管制Only
                                If JobMaker_Form.Spec_Seismic_Only_CheckBox.Checked Then
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_Seismic_ONLY).RefersToRange.Cells.Value =
                                        $"(Only {JobMaker_Form.Spec_Seismic_Only_TextBox.Text})"
                                End If

                                '地震管制 感知器Only ------------------------------------------
                                If JobMaker_Form.Spec_SeismicSensor_Only_CheckBox.Checked Then
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_Seismic_SENSOR_ONLY).RefersToRange.Cells.Value =
                                        $"(Only {JobMaker_Form.Spec_SeismicSensor_Only_TextBox.Text})"
                                End If
                                msExcel_workbook.Names.Item(get_NameManager.SetTable_Seismic_SENSOR).RefersToRange.Cells.Value =
                                    JobMaker_Form.Spec_SeismicSensor_ComboBox.Text
                                '------------------------------------------ 地震管制 感知器Only 

                                '地震管制 自動解除開關Only ------------------------------------
                                If JobMaker_Form.Spec_SeismicSW_Only_CheckBox.Checked Then
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_SeismicSW_ONLY).RefersToRange.Cells.Value =
                                        $"(Only {JobMaker_Form.Spec_SeismicSW_Only_TextBox.Text})"
                                End If

                                If JobMaker_Form.Spec_SeismicSW_ComboBox.Text = get_NameManager.TB_X Then
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_SeismicSW_WITH
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                Else
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_SeismicSW_WITHOUT
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                End If
                                '------------------------------------ 地震管制 自動解除開關Only 
                            End If
                            '------------------------------------------------------------------------------------------------------ 地震管制運轉

                            ' 車廂管制運轉燈 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_CPI
                            excelWriteIn(JobMaker_Form.Spec_CPI_ComboBox.Text,
                                         get_NameManager.SPEC_CPI,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_CPI_ComboBox.Text = get_NameManager.TB_O Then

                                Dim cpiEmr_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_CPI_EMER
                                                                ).RefersToRange.Value '取得 管制 運轉燈內的文字內容

                                Dim cpiFm_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_CPI_FM
                                                                ).RefersToRange.Value '取得 緊急 運轉燈內的文字內容
                                Dim cpiOlt_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_CPI_OLT).RefersToRange.Value '取得 滿載 運轉燈內的文字內容

                                '車廂管制燈-地震
                                If JobMaker_Form.Spec_CpiSeismic_ComboBox.Text = get_NameManager.TB_X Then
                                    Dim sei_val As String =
                                           msExcel_workbook.Names.Item(get_NameManager.SetTable_CPI_SEISMIC
                                                                       ).RefersToRange.Cells.Value '取得地震時的文字內容
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_CPI_EMER
                                                                ).RefersToRange.Characters(InStr(cpiEmr_val, sei_val), Len(sei_val)).
                                                                Font.Strikethrough = True
                                End If
                                '車廂管制燈-火災
                                If JobMaker_Form.Spec_CpiFire_ComboBox.Text = get_NameManager.TB_X Then
                                    Dim fire_val As String =
                                        msExcel_workbook.Names.Item(get_NameManager.SetTable_CPI_FIRE
                                                                    ).RefersToRange.Cells.Value '取得火災時的文字內容
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_CPI_EMER
                                                                ).RefersToRange.Characters(InStr(cpiEmr_val, fire_val), Len(fire_val)).
                                                                Font.Strikethrough = True
                                End If
                                '車廂管制燈-自家發
                                If JobMaker_Form.Spec_CpiEmer_ComboBox.Text = get_NameManager.TB_X Then
                                    Dim emerP_val As String =
                                        msExcel_workbook.Names.Item(get_NameManager.SetTable_CPI_EMER
                                                                    ).RefersToRange.Cells.Value '取得自家發時的文字內容
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_CPI_EMER
                                                                ).RefersToRange.Characters(InStr(cpiEmr_val, emerP_val), Len(emerP_val)).
                                                                Font.Strikethrough = True
                                End If
                                '車廂管制燈-緊急
                                If JobMaker_Form.Spec_CpiFM_ComboBox.Text = get_NameManager.TB_X Then
                                    Dim fm_val As String =
                                        msExcel_workbook.Names.Item(get_NameManager.SetTable_CPI_FIREMAN
                                                                    ).RefersToRange.Cells.Value '取得緊急時的文字內容

                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_CPI_FM
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                End If
                                '車廂管制燈-滿載
                                If JobMaker_Form.Spec_CpiOLT_ComboBox.Text = get_NameManager.TB_X Then
                                    Dim olt_val As String =
                                        msExcel_workbook.Names.Item(get_NameManager.SetTable_CPI_OLT
                                                                    ).RefersToRange.Cells.Value '取得超載時的文字內容
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_CPI_OLT
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                End If
                                '車廂管制燈-滿載Only
                                If JobMaker_Form.Spec_CpiOLT_Only_CheckBox.Checked Then
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CPI_OLT_ONLY).RefersToRange.Cells.Value =
                                        $"(Only {JobMaker_Form.Spec_CpiOLT_Only_TextBox.Text})"
                                End If
                            End If
                            '------------------------------------------------------------------------------------------------------ 車廂管制運轉燈

                            ' 車廂上到著鈴 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Car_Gong
                            excelWriteIn(JobMaker_Form.Spec_CarGong_ComboBox.Text,
                                         get_NameManager.SPEC_CAR_GONG,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_CarGong_ComboBox.Text = get_NameManager.TB_O Then

                                Dim spec_car_gong_pos As String = get_NameManager.SPEC_CAR_GONG_POS
                                Dim spec_car_gong_cartop As String = get_NameManager.SPEC_CAR_GONG_CARTOP
                                Dim spec_car_gong_cartopbtm As String = get_NameManager.SPEC_CAR_GONG_CARTOPBTM
                                Dim spec_car_gong_cob As String = get_NameManager.SPEC_CAR_GONG_COB
                                Dim spec_car_gong_vonic As String = get_NameManager.SPEC_CAR_GONG_VONIC

                                Dim settable_car_top As String = get_NameManager.SetTable_CAR_TOP
                                Dim settable_car_top_btm As String = get_NameManager.SetTable_CAR_TOP_BTM
                                Dim settable_car_cob As String = get_NameManager.SetTable_CAR_COB
                                Dim settable_car_vonic As String = get_NameManager.SetTable_CAR_VONIC

                                Dim pos_val As String =
                                    msExcel_workbook.Names.Item(spec_car_gong_pos).RefersToRange.Cells.Value '取得 位置 的文字內容
                                Dim carTop_val As String =
                                    msExcel_workbook.Names.Item(settable_car_top).RefersToRange.Cells.Value '取得 車廂上 的文字內容
                                Dim carTopBtm_val As String =
                                    msExcel_workbook.Names.Item(settable_car_top_btm).RefersToRange.Cells.Value '取得 車廂上下 的文字內容
                                Dim cob_val As String =
                                    msExcel_workbook.Names.Item(settable_car_cob).RefersToRange.Cells.Value '取得 COB 的文字內容
                                Dim inVonic_val As String =
                                    msExcel_workbook.Names.Item(settable_car_vonic).RefersToRange.Cells.Value '取得 VONIC 的文字內容

                                'Car 車廂上
                                If JobMaker_Form.Spec_CarGong_Top_CheckBox.Checked = False And
                                   JobMaker_Form.Spec_CarGong_Top_TextBox.Text = get_NameManager.TB_CarTop Then
                                    '無
                                    msExcel_workbook.Names.Item(spec_car_gong_pos
                                                                ).RefersToRange.Characters(InStr(pos_val, carTop_val), Len(carTop_val)).
                                                                Font.Strikethrough = True
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
                                         get_NameManager.SPEC_HALL_GONG,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 乘場到著鈴

                            ' 乘場信號文字 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_HPI
                            excelWriteIn(JobMaker_Form.Spec_HPIMsg_ComboBox.Text,
                                         get_NameManager.SPEC_HPI,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_HPIMsg_ComboBox.Text = get_NameManager.TB_O Then
                                Dim spec_hpi_msg As String = get_NameManager.SPEC_HPI_MSG
                                Dim spec_hpi_main As String = get_NameManager.SPEC_HPI_MAIN

                                Dim halMsg_val As String =
                                    msExcel_workbook.Names.Item(spec_hpi_msg).RefersToRange.Cells.Value '取得 乘場燈 的文字內容
                                Dim halMain_val As String =
                                    msExcel_workbook.Names.Item(spec_hpi_main).RefersToRange.Cells.Value '取得 保養中 的文字內容
                                Dim olt_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_HALL_OLT
                                                                ).RefersToRange.Cells.Value '取得 滿載 的文字內容
                                Dim main_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_HALL_MAIN
                                                                ).RefersToRange.Cells.Value '取得 保養 的文字內容
                                Dim indep_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_HALL_INDEP
                                                                ).RefersToRange.Cells.Value '取得 專用 的文字內容
                                Dim fm_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_HALL_FM
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
                                         get_NameManager.SPEC_DR_HOLD,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 開門延長

                            ' 刷卡機 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_CRD
                            excelWriteIn(JobMaker_Form.Spec_CRD_ComboBox.Text,
                                         get_NameManager.SPEC_CRD,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_CRD_ComboBox.Text = get_NameManager.TB_O Then

                                Dim type_all_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_TYPE_ALL
                                                                ).RefersToRange.Cells.Value '取得 全層管制 的文字內容
                                Dim type_notall_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_TYPE_NOTALL
                                                                ).RefersToRange.Cells.Value '取得 分層管制 的文字內容
                                Dim crd_Y_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_SPEC_Y
                                                                ).RefersToRange.Cells.Value '取得 式樣有 的文字內容
                                Dim crd_N_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_SPEC_N
                                                                ).RefersToRange.Cells.Value '取得 式樣無 的文字內容
                                Dim rvs_crd_Y As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_RVS_CALL_Y
                                                                ).RefersToRange.Cells.Value '取得 逆呼有 的文字內容
                                Dim rvs_crd_N_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_RVS_CALL_N
                                                                ).RefersToRange.Cells.Value '取得 逆呼無 的文字內容
                                Dim anti_crd_Y_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_ANTI_Y
                                                                ).RefersToRange.Cells.Value '取得 嬉戲有 的文字內容
                                Dim anti_crd_N_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_ANTI_N
                                                                ).RefersToRange.Cells.Value '取得 嬉戲無 的文字內容
                                Dim time_crd_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_CRD_TIME_SET
                                                                ).RefersToRange.Cells.Value '取得 時間 的文字內容

                                Dim spec_crd_type As String = get_NameManager.SPEC_CRD_TYPE
                                Dim spec_crd_rgl4_y As String = get_NameManager.SPEC_CRD_RGL4_Y
                                Dim spec_crd_rgl4_n As String = get_NameManager.SPEC_CRD_RGL4_N
                                Dim spec_crd_rgl5_y As String = get_NameManager.SPEC_CRD_RGL5_Y
                                Dim spec_crd_rgl5_n As String = get_NameManager.SPEC_CRD_RGL5_N

                                Dim crd_type As String =
                                    msExcel_workbook.Names.Item(spec_crd_type).RefersToRange.Cells.Value '分層或全層

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
                                         get_NameManager.SPEC_EMER_POWER,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_Emer_ComboBox.Text = get_NameManager.TB_O Then

                                '自家發Signal --------------------------------------------------------------------------------
                                Dim spec_emer_signal As String = get_NameManager.SPEC_EMER_SIGNAL
                                Dim sig_val As String =
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
                                excelWriteIn(JobMaker_Form.Spec_EmerCapacity_NumericUpDown.Value,
                                             get_NameManager.SPEC_EMER_CAPCITY,
                                             msExcel_workbook)
                                '---------------------------------------------------------------------------------- 自家發容量 

                                '自家發入力點 -------------------------------------------------------------------------------
                                excelWriteIn(JobMaker_Form.Spec_EmerInput_ComboBox.Text,
                                             get_NameManager.SPEC_EMER_INPUT,
                                             msExcel_workbook)
                                '------------------------------------------------------------------------------- 自家發入力點 

                                '自家發Address -----------------------------------------------------------------------------
                                excelWriteIn(JobMaker_Form.Spec_EmerAddress_ComboBox.Text,
                                             get_NameManager.SPEC_EMER_ADDRESS,
                                             msExcel_workbook)
                                '----------------------------------------------------------------------------- 自家發Address 

                                '自家發Group -----------------------------------------------------------------------------

                                Dim JM_Spec_Emer As String() =
                                    {get_NameManager.SPEC_EMER_POWER_GROUP,
                                     get_NameManager.SPEC_EMER_POWER_CarName,
                                     get_NameManager.SPEC_EMER_POWER_EscapeFL,
                                     get_NameManager.SPEC_EMER_POWER_RETURN,
                                     get_NameManager.SPEC_EMER_POWER_CONTINUE
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
                                         get_NameManager.SPEC_LANDIC,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ Landic

                            ' 基準階賦歸 -----------------------------------------------------------------------------------------------------
                        Case usr_Spec_MLF_Return
                            excelWriteIn(JobMaker_Form.Spec_MFLReturn_ComboBox.Text,
                                         get_NameManager.SPEC_MLF_RETURN,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_MFLReturn_ComboBox.Text = get_NameManager.TB_O Then
                                msExcel_workbook.Names.Item(get_NameManager.SetTable_MAIN_FL).RefersToRange.Cells.Value =
                                    JobMaker_Form.Spec_MFLReturn_FL_TextBox.Text & "階"
                            End If
                            '------------------------------------------------------------------------------------------------------ 基準階賦歸

                            ' VONIC --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Vonic
                            excelWriteIn(JobMaker_Form.Spec_Vonic_ComboBox.Text,
                                         get_NameManager.SPEC_VONIC,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_Vonic_ComboBox.Text = get_NameManager.TB_O Then
                                If JobMaker_Form.Spec_Vonic_standard_ComboBox.Text = get_NameManager.TB_O Then
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_VONIC_NSTD_C
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_VONIC_NSTD_E
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_VONIC_STD_C
                                                                ).RefersToRange.Cells.Font.Strikethrough = False
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_VONIC_STD_E
                                                                ).RefersToRange.Cells.Font.Strikethrough = False
                                Else
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_VONIC_NSTD_C
                                                                ).RefersToRange.Cells.Font.Strikethrough = False
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_VONIC_NSTD_E
                                                                ).RefersToRange.Cells.Font.Strikethrough = False
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_VONIC_STD_C
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_VONIC_STD_E
                                                                ).RefersToRange.Cells.Font.Strikethrough = True
                                End If
                            End If

                            '------------------------------------------------------------------------------------------------------ VONIC

                            ' 殘障COB --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_WCOB
                            excelWriteIn(JobMaker_Form.Spec_WCOB_ComboBox.Text,
                                         get_NameManager.SPEC_WCOB,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_WCOB_ComboBox.Text = get_NameManager.TB_O Then


                                Dim spec_sub_wcob As String = get_NameManager.SPEC_WCOB_SUB
                                Dim spec_whb_bz As String = get_NameManager.SPEC_WCOB_BZ
                                Dim spec_whb_ring As String = get_NameManager.SPEC_WCOB_RING

                                Dim bz_Y As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_WCOB_BZ_Y
                                                                ).RefersToRange.Cells.Value '取得 BZ有 的文字內容
                                Dim bz_N As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_WCOB_BZ_N
                                                                ).RefersToRange.Cells.Value '取得 BZ無 的文字內容
                                Dim ring_Y As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_WCOB_RING_Y
                                                                ).RefersToRange.Cells.Value '取得 RING有 的文字內容
                                Dim ring_N As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_WCOB_RING_N
                                                                ).RefersToRange.Cells.Value '取得 RING無 的文字內容
                                Dim sub_wcob_val As String =
                                    msExcel_workbook.Names.Item(spec_sub_wcob).RefersToRange.Cells.Value
                                Dim bz_val As String =
                                    msExcel_workbook.Names.Item(spec_whb_bz).RefersToRange.Cells.Value
                                Dim ring_val As String =
                                    msExcel_workbook.Names.Item(spec_whb_ring).RefersToRange.Cells.Value

                                If JobMaker_Form.Spec_WCOB_only_CheckBox.Checked Then 'COB Only
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_WCOB_ONLY).RefersToRange.Cells.Value =
                                        $"(Only {JobMaker_Form.Spec_WCOB_only_TextBox.Text})"
                                End If
                                If JobMaker_Form.Spec_WSCOB_only_CheckBox.Checked Then 'SCOB Only
                                    msExcel_workbook.Names.Item(get_NameManager.SetTable_WSCOB_ONLY).RefersToRange.Cells.Value =
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
                                If JobMaker_Form.Spec_WCOB_Ring_ComboBox.Text = "鳴動" Then '鳴動
                                    msExcel_workbook.Names.Item(spec_whb_bz
                                                                ).RefersToRange.Characters(InStr(bz_val, bz_N), Len(bz_N)).
                                                                Font.Strikethrough = True
                                    msExcel_workbook.Names.Item(spec_whb_ring
                                                                ).RefersToRange.Characters(InStr(ring_val, ring_N), Len(ring_N)).
                                                                Font.Strikethrough = True
                                ElseIf JobMaker_Form.Spec_WCOB_Ring_ComboBox.Text = "不鳴動" Then '不鳴動
                                    msExcel_workbook.Names.Item(spec_whb_bz
                                                                ).RefersToRange.Characters(InStr(bz_val, bz_Y), Len(bz_Y)).
                                                                Font.Strikethrough = True
                                    msExcel_workbook.Names.Item(spec_whb_ring
                                                               ).RefersToRange.Characters(InStr(ring_val, ring_Y), Len(ring_Y)).
                                                               Font.Strikethrough = True
                                End If
                            End If
                            '------------------------------------------------------------------------------------------------------ 殘障HB

                            ' ELVIC --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_Elvic
                            excelWriteIn(JobMaker_Form.Spec_Elvic_ComboBox.Text,
                                         get_NameManager.SPEC_ELVIC,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_Elvic_ComboBox.Text = get_NameManager.TB_O Then
                                Dim elv_ele_grp As CheckBox() =
                                                                {JobMaker_Form.Spec_Elvic_Parking_CheckBox, JobMaker_Form.Spec_Elvic_VIP_CheckBox,
                                                                 JobMaker_Form.Spec_Elvic_Indep_CheckBox, JobMaker_Form.Spec_Elvic_FloorLockOut_CheckBox,
                                                                 JobMaker_Form.Spec_Elvic_Express_CheckBox, JobMaker_Form.Spec_Elvic_ReturnFL_CheckBox
                                                                }
                                Dim elv_grp_grp As CheckBox() =
                                                                    {JobMaker_Form.Spec_Elvic_Traffic_Peak_CheckBox,
                                                                     JobMaker_Form.Spec_Elvic_MainFL_CheckBox, JobMaker_Form.Spec_Elvic_FloorLockOut_GR_CheckBox,
                                                                     JobMaker_Form.Spec_Elvic_Zoning_CheckBox, JobMaker_Form.Spec_Elvic_CarCall_CheckBox
                                                                    }
                                Dim elv_other_grp As CheckBox() =
                                                                      {JobMaker_Form.Spec_Elvic_Fire_CheckBox, JobMaker_Form.Spec_Elvic_Wavic_CheckBox,
                                                                       JobMaker_Form.Spec_Elvic_CRD_CheckBox
                                                                      }
                                Dim num_grp As String() = {"①", "②", "③", "④", "⑤", "⑥"}
                                Dim sh_name As String =
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_ELVIC).RefersToRange.Worksheet.Name
                                Dim elv_Row As Integer =
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_ELVIC_CMD).RefersToRange.Row '號機名是第n行

                                Dim elv_Col As Integer =
                                    msExcel_workbook.Names.Item(get_NameManager.SPEC_ELVIC_CMD).RefersToRange.Column '號機名是第n列
                                Dim first_i As Integer = 0
                                '第一大象
                                For i = 1 To elv_ele_grp.Count
                                    If elv_ele_grp(i - 1).Checked Then
                                        msExcel_workbook.Names.Item(get_NameManager.SPEC_ELVIC_CMD).RefersToRange.Cells.Value =
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
                                            Dim trafficString As String = ""
                                            If JobMaker_Form.Spec_Elvic_Traffic_UpPeak_CheckBox.Checked Then
                                                trafficString &= JobMaker_Form.Spec_Elvic_Traffic_UpPeak_CheckBox.Text
                                            End If
                                            If JobMaker_Form.Spec_Elvic_Traffic_DownPeak_CheckBox.Checked Then
                                                trafficString &= JobMaker_Form.Spec_Elvic_Traffic_DownPeak_CheckBox.Text
                                            End If
                                            If JobMaker_Form.Spec_Elvic_Traffic_Lunch_CheckBox.Checked Then
                                                trafficString &= JobMaker_Form.Spec_Elvic_Traffic_Lunch_CheckBox.Text
                                            End If

                                            If trafficString <> "" Then
                                                msExcel_workbook.Worksheets(sh_name).Cells(elv_Row, elv_Col).Value =
                                                    "   " & trafficString
                                                elv_Row = elv_Row + 1
                                                msExcel_workbook.Worksheets(sh_name).Range($"{elv_Row}:{elv_Row}").Insert
                                            End If
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
                                         get_NameManager.SPEC_HLL,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 乘場廳燈

                            ' 運轉手盤 --------------------------------------------------------------------------------------------------------
                        Case usr_Spec_ATT
                            excelWriteIn(JobMaker_Form.Spec_ATT_ComboBox.Text,
                                         get_NameManager.SPEC_ATT,
                                         msExcel_workbook)
'------------------------------------------------------------------------------------------------------ 運轉手盤 

                            ' 浸水管制運轉 -----------------------------------------------------------------------------------------------------
                        Case usr_Spec_Flood
                            excelWriteIn(JobMaker_Form.Spec_Flood_ComboBox.Text,
                                         get_NameManager.SPEC_FLOOD,
                                         msExcel_workbook)

                            msExcel_workbook.Names.Item(get_NameManager.SetTable_FLOOD_FL).RefersToRange.Cells.Value =
                                JobMaker_Form.Spec_Flood_FL_TextBox.Text
                            '------------------------------------------------------------------------------------------------------ 浸水管制運轉

                            ' LS1M ------------------------------------------------------------------------------------------------------
                        Case usr_Spec_LS1M
                            excelWriteIn(JobMaker_Form.Spec_LS1M_ComboBox.Text,
                                         get_NameManager.SPEC_LS1M,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ LS1M

                            ' 電力回升 ------------------------------------------------------------------------------------------------------
                        Case usr_Spec_PRU
                            excelWriteIn(JobMaker_Form.Spec_PRU_ComboBox.Text,
                                         get_NameManager.SPEC_PRU,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 電力回升

                            ' 正背門 -----------------------------------------------------------------------------------------------------
                        Case usr_Spec_FrontRear_DR
                            excelWriteIn(JobMaker_Form.Spec_FrontRearDr_ComboBox.Text,
                                         get_NameManager.SPEC_FRONT_REAR_DR,
                                         msExcel_workbook)
                            '------------------------------------------------------------------------------------------------------ 正背門

                            ' 單群控切換 -------------------------------------------------------------------------------------------------
                        Case usr_Spec_OpeSw
                            excelWriteIn(JobMaker_Form.Spec_OpeSw_ComboBox.Text,
                                         get_NameManager.SPEC_OPE_SW,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_OpeSw_ComboBox.Text = get_NameManager.TB_O Then
                                excelWriteIn(JobMaker_Form.Spec_OpeSw_DevicePos_TextBox.Text,
                                             get_NameManager.SetTable_OpeSW_Content,
                                             msExcel_workbook)
                                excelWriteIn(JobMaker_Form.Spec_OpeSw_InputPos_ComboBox.Text,
                                             get_NameManager.SPEC_OPE_SW_POS,
                                             msExcel_workbook)
                                excelWriteIn(JobMaker_Form.Spec_OpeSw_InputAddress_TextBox.Text,
                                             get_NameManager.SPEC_OPE_SW_ADDRESS,
                                             msExcel_workbook)
                            End If
                            '---------------------------------------------------------------------------------------------- 單群控切換
                        Case usr_Spec_WTB
                            excelWriteIn(JobMaker_Form.Spec_WTB_ComboBox.Text,
                                         get_NameManager.SPEC_WTB,
                                         msExcel_workbook)

                            If JobMaker_Form.Spec_WTB_ComboBox.Text = get_NameManager.TB_O Then
                                '故障燈
                                If JobMaker_Form.Spec_WTB_Error_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_ERROR)
                                End If
                                '休止燈
                                If JobMaker_Form.Spec_WTB_Stop_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_STOP)
                                End If
                                '消防運轉燈
                                If JobMaker_Form.Spec_WTB_FM_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_FM)
                                End If
                                '地震管制燈
                                If JobMaker_Form.Spec_WTB_EQ_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_EQ)
                                End If
                                '專用運轉燈
                                If JobMaker_Form.Spec_WTB_Indep_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_INDEP)
                                End If
                                '正常運轉燈
                                If JobMaker_Form.Spec_WTB_Normal_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_NORMAL)
                                End If
                                '緊急電源燈
                                If JobMaker_Form.Spec_WTB_Urgent_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_URGENT)
                                End If
                                '火災燈
                                If JobMaker_Form.Spec_WTB_FO_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_FO)
                                End If
                                '自家發燈
                                If JobMaker_Form.Spec_WTB_EmerPow_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_EMERPOWER)
                                End If
                                '警示燈
                                If JobMaker_Form.Spec_WTB_Alart_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_ALART)
                                End If
                                '地震強度燈
                                If JobMaker_Form.Spec_WTB_EQMac_ComboBox.Text = get_NameManager.TB_X Then
                                    getMathOnExcel.strikeThrough_allText_onWorkSht(msExcel_workbook, get_NameManager.SPEC_WTB_EQMAC)
                                End If
                            End If
                    End Select
                    'Catch ex As Exception
                    '    JobMaker_Form.ResultFailOutput_TextBox.Text +=
                    '        ($"<{JobMaker_Form.JMFileCho_Spec_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_TWSpec_str}>{vbCrLf}")
                    'End Try
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
            Dim usr_IMP_MachineType As String =
                JobMaker_Form.Imp_MachineRoom_ComboBox.Name
            'Dim usr_IMP_FAN  As String=
            '    JobMaker_Form.Imp_FAN_ComboBox.Name
            Dim usr_IMP_OverBalance As String =
                JobMaker_Form.Imp_OverBalance_ComboBox.Name
            Dim usr_IMP_WHB As String =
                JobMaker_Form.Imp_WHB_ComboBox.Name
            Dim usr_IMP_DoorType As String =
                JobMaker_Form.Imp_DoorType_TextBox.Name
            Dim usr_IMP_HIN As String =
                JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls.Count


            Dim usrInput_IMP_arr() As String = {usr_IMP_MachineType, usr_IMP_OverBalance,
                                                usr_IMP_WHB, usr_IMP_DoorType, usr_IMP_HIN}
            Dim i_ImpStr As String

            '輸入相對應的check list值
            For Each i_ImpStr In usrInput_IMP_arr
                If i_ImpStr <> "" Then
                    Try
                        Select Case i_ImpStr

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
                                             get_NameManager.IMPORTANT_BALANCE,
                                             msExcel_workbook)
                            Case usr_IMP_WHB
                                excelWriteIn(JobMaker_Form.Imp_WHB_ComboBox.Text,
                                             get_NameManager.IMPORTANT_WCOB,
                                             msExcel_workbook)
                            Case usr_IMP_DoorType
                                excelWriteIn(JobMaker_Form.Imp_DoorType_TextBox.Text,
                                             get_NameManager.IMPORTANT_DOOR,
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
        If JobMaker_Form.HallIndicator_FlowLayoutPanel.Controls.Count <> 0 Then
            Dim Imp_HIN_FL_Col As Integer =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.IMPORTANT_HIN_FL)
            Dim Imp_HIN_FL_Row As Integer =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.IMPORTANT_HIN_FL)
            Dim Imp_HIN_Col As Integer =
                getMathOnExcel.getCol_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.IMPORTANT_HIN)
            Dim Imp_HIN_Row As Integer =
                getMathOnExcel.getRow_fromNameManager_typeIsCell(msExcel_workbook, get_NameManager.IMPORTANT_HIN)

            Dim Imp_SheetName As String =
                getMathOnExcel.getWorksheetName_fromNameManager(msExcel_workbook, get_NameManager.IMPORTANT_HIN_FL)

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
            Dim with_val As String =
                getMathOnExcel.getValue_formNameManager(msExcel_workbook, get_NameManager.SetTable_RESULT_WITH)
            Dim without_val As String =
                getMathOnExcel.getValue_formNameManager(msExcel_workbook, get_NameManager.SetTable_RESULT_WITHOUT)

            Dim usr_MMIC_MachineType As String =
                get_NameManager.MMIC_MACHINE_TYPE
            Dim usr_MMIC_FLEX As String =
                get_NameManager.MMIC_OPERATION
            Dim usr_SV_FLEX_N As String =
                get_NameManager.MMIC_FLEX_N_SV

            Dim usr_MMIC_CP43x As String =
                JobMaker_Form.MMIC_MR_CP43x_ComboBox.Name
            Dim usr_MMIC_CarObj As String =
                JobMaker_Form.MMIC_MR_NumericUpDown.Name
            Dim usr_MMIC_E_BASE As String =
                JobMaker_Form.MMIC_MR_EBase_ComboBox.Name
            Dim usr_MMIC_E_CarObj As String =
                JobMaker_Form.MMIC_MR_E_NumericUpDown.Name

            Dim usr_SV_CarObj As String =
                JobMaker_Form.MMIC_SV_NumericUpDown.Name
            Dim usr_SV_E_BASE As String =
                JobMaker_Form.MMIC_SV_EBase_ComboBox.Name
            Dim usr_SV_E_CarObj As String =
                JobMaker_Form.MMIC_SV_E_NumericUpDown.Name

            Dim usr_VD10_ROM_Device As String =
                JobMaker_Form.MMIC_VD10_ROM_ComboBox.Name
            Dim usr_VD10_Quantity As String =
                JobMaker_Form.MMIC_VD10_Quantity_ComboBox.Name
            Dim usr_VD10_CarObj As String =
                JobMaker_Form.MMIC_VD10_NumericUpDown.Name

            Dim usrInput_MMIC_arr() As String = {usr_MMIC_MachineType, usr_MMIC_FLEX,
                                                 usr_MMIC_CP43x, usr_MMIC_CarObj,
                                                 usr_MMIC_E_BASE, usr_MMIC_E_CarObj,
                                                 usr_SV_CarObj, usr_SV_E_CarObj,
                                                 usr_SV_E_BASE,
                                                 usr_VD10_CarObj,
                                                 usr_VD10_ROM_Device, usr_VD10_Quantity}
            Dim dyCtrlName As DynamicControlName = New DynamicControlName
            dyCtrlName.JobMaker_MMICInfo()

            '輸入相對應的MMIC值
            For Each i_mmicStr As String In usrInput_MMIC_arr
                If i_mmicStr <> "" Then
                    Try
                        Select Case i_mmicStr
                            Case usr_MMIC_MachineType
                                '[機種]
                                'Dim JM_MACHINE_TYPE As String() = {usr_MMIC_MachineType}
                                'Dim spec_stored As Spec_StoredJobData = New Spec_StoredJobData
                                'dynamicControl_writeInExcel(JobMaker_Form.Spec_MachineType_NumericUpDown.Value,
                                '                            usr_MMIC_MachineType,
                                '                            JM_MACHINE_TYPE,
                                '                            JobMaker_Form.Spec_MachineType_Panel,
                                '                            spec_stored.LoadStored_PanelType.SingleLayer_Panel,
                                '                            {dyCtrlName.Spec_MachineType_ComboBox}.Count,
                                '                            {dyCtrlName.Spec_MachineType_ComboBox},
                                '                            msExcel_workbook)
                            Case usr_MMIC_FLEX
                                '[FLEX-N幾百]
                                returnError_specName = usr_MMIC_FLEX
                                Dim JM_MMIC_FLEX As String() = {usr_MMIC_FLEX}
                                Dim JM_MACHINE_TYPE As String() = {usr_MMIC_MachineType}

                                dynamicControl_writeInExcel_MMIC(JobMaker_Form.Spec_FLEX_N_NumericUpDown, JobMaker_Form.Spec_MachineType_NumericUpDown,
                                                                 usr_MMIC_FLEX,
                                                                 JM_MMIC_FLEX, JM_MACHINE_TYPE,
                                                                 JobMaker_Form.Spec_FLEX_N_Panel, JobMaker_Form.Spec_MachineType_Panel,
                                                                 {dyCtrlName.Spec_FLEX_N_ComboBox}.Count, {dyCtrlName.Spec_FLEX_N_ComboBox},
                                                                 {dyCtrlName.Spec_MachineType_ComboBox}.Count, {dyCtrlName.Spec_MachineType_ComboBox},
                                                                 msExcel_workbook)
                            Case usr_SV_FLEX_N
                                excelWriteIn(JobMaker_Form.MMIC_FLEX_N_ComboBox.Text,
                                             get_NameManager.MMIC_FLEX_N_SV,
                                             msExcel_workbook)
                            Case usr_MMIC_CP43x
                                '[MR-MMIC > 有無CP43]
                                returnError_specName = get_NameManager.MMIC_CP43x
                                Dim cp43x_val As String =
                                    msExcel_workbook.Names.Item(get_NameManager.MMIC_CP43x).RefersToRange.Value '取得 有 內的文字內容

                                If JobMaker_Form.MMIC_MR_CP43x_ComboBox.Text = get_NameManager.TB_WITHOUT Then
                                    msExcel_workbook.Names.Item(get_NameManager.MMIC_CP43x
                                                                ).RefersToRange.Characters(InStr(cp43x_val, with_val), Len(with_val)).
                                                                Font.Strikethrough = True
                                Else
                                    msExcel_workbook.Names.Item(get_NameManager.MMIC_CP43x
                                                                ).RefersToRange.Characters(InStr(cp43x_val, without_val), Len(without_val)).
                                                                Font.Strikethrough = True
                                End If


                            Case usr_MMIC_E_BASE
                                '[MR-MMIC > EEPROM DATA > BASE]
                                returnError_specName = get_NameManager.MMIC_EBase
                                excelWriteIn(JobMaker_Form.MMIC_MR_EBase_ComboBox.Text,
                                             get_NameManager.MMIC_EBase,
                                             msExcel_workbook)

                            Case usr_MMIC_E_CarObj
                                '[MR-MMIC > EEPROM DATA > 自動生成控制項]
                            Case usr_SV_E_BASE
                                '[SV > EEPROM DATA > BASE]
                                returnError_specName = get_NameManager.SV_EBase
                                excelWriteIn(JobMaker_Form.MMIC_SV_EBase_ComboBox.Text,
                                             get_NameManager.SV_EBase,
                                             msExcel_workbook)
                            Case usr_SV_E_CarObj
                                '[SV > 自動生成控制項]
                            Case usr_VD10_ROM_Device
                                '[VD10 > ROM DEVICE]
                                returnError_specName = get_NameManager.VONIC_ROM_Device
                                excelWriteIn(JobMaker_Form.MMIC_VD10_ROM_ComboBox.Text,
                                             get_NameManager.VONIC_ROM_Device,
                                             msExcel_workbook)
                            Case usr_VD10_Quantity
                                '[VD10 > QUANTITY 幾片]
                                returnError_specName = get_NameManager.VONIC_Quantity
                                excelWriteIn(JobMaker_Form.MMIC_VD10_Quantity_ComboBox.Text,
                                             get_NameManager.VONIC_Quantity,
                                             msExcel_workbook)
                            Case usr_MMIC_CarObj
                                '[MR-MMIC > 自動生成控制項]
                                returnError_specName = usr_MMIC_CarObj
                                Dim JM_MMIC_MR As String() = {get_NameManager.MMIC_CarNo,
                                                              get_NameManager.MMIC_CarObj}

                                Dim JM_MMIC_MR_E As String() = {get_NameManager.MMIC_ECarNo,
                                                                get_NameManager.MMIC_ECarObj}

                                dynamicControl_writeInExcel_MMIC(JobMaker_Form.MMIC_MR_NumericUpDown, JobMaker_Form.MMIC_MR_E_NumericUpDown,
                                                                 get_NameManager.MMIC_CarNo,
                                                                 JM_MMIC_MR, JM_MMIC_MR_E,
                                                                 JobMaker_Form.MMIC_MR_Panel, JobMaker_Form.MMIC_MR_E_Panel,
                                                                 dyCtrlName.JobMaker_MMIC_Mr_InfoName_Array.Count, dyCtrlName.JobMaker_MMIC_Mr_InfoName_Array,
                                                                 dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array.Count, dyCtrlName.JobMaker_MMIC_MrEBase_InfoName_Array,
                                                                 msExcel_workbook)
                            Case usr_SV_CarObj
                                '[SV > 自動生成控制項]
                                returnError_specName = usr_SV_CarObj
                                Dim JM_MMIC_SV As String() = {get_NameManager.SV_CarNo,
                                                              get_NameManager.SV_CarObj}

                                Dim JM_MMIC_SV_E As String() = {get_NameManager.SV_ECarNo,
                                                                get_NameManager.SV_ECarObj}


                                dynamicControl_writeInExcel_MMIC(JobMaker_Form.MMIC_SV_NumericUpDown, JobMaker_Form.MMIC_SV_E_NumericUpDown,
                                                                 get_NameManager.SV_CarNo,
                                                                 JM_MMIC_SV, JM_MMIC_SV_E,
                                                                 JobMaker_Form.MMIC_SV_Panel, JobMaker_Form.MMIC_SV_E_Panel,
                                                                 dyCtrlName.JobMaker_MMIC_Sv_InfoName_Array.Count, dyCtrlName.JobMaker_MMIC_Sv_InfoName_Array,
                                                                 dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array.Count, dyCtrlName.JobMaker_MMIC_SvEBase_InfoName_Array,
                                                                 msExcel_workbook)

                            Case usr_VD10_CarObj
                                '[SV > 自動生成控制項]
                                returnError_specName = usr_VD10_CarObj
                                Dim JM_MMIC_VONIC As String() = {get_NameManager.VONIC_CarNo,
                                                                 get_NameManager.VONIC_CarObj}
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
                            ($"<{JobMaker_Form.JMFileCho_ChkList_TextBox.Text}.xls>中無名稱管理員:<{returnError_specName}>{vbCrLf}上述設定值為:<{i_mmicStr}>{vbCrLf},{ex.Message}{vbCrLf}")
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
        Try
            If usr IsNot "" Then
                returnError_specName = spec '錯誤回報
                msExcel_workbook.Names.Item(spec).RefersToRange.Value = usr

                JobMaker_Form.ResultOutput_TextBox.Text += ($"名稱管理員:{spec} / 值:{usr} 寫入成功{vbCrLf}")
            Else
                returnError_specName = spec '錯誤回報
                JobMaker_Form.ResultOutput_TextBox.Text += ($"名稱管理員:{spec} / 值:{usr} 是空值寫入失敗{vbCrLf}")
            End If
        Catch ex As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"[Error] 名稱管理員:{spec}"
            JobMaker_Form.ResultFailOutput_TextBox.Text += ex.Message
            writeIntoErrorInfoTxt($"[Error] 名稱管理員:{spec}   {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' 將chkbox有打勾的項目和usr(輸入資料)寫入msExcel_workbook(目標excel)的spec(名稱管理員)
    ''' </summary>
    ''' <param name="usr"></param>
    ''' <param name="spec"></param>
    ''' <param name="chkbox"></param>
    ''' <param name="msExcel_workbook"></param>
    Overloads Sub excelWriteIn(usr As String, spec As String, chkbox As CheckBox, msExcel_workbook As Excel.Workbook)
        Try
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
        Catch ex As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"[Error] 名稱管理員:{spec}"
            JobMaker_Form.ResultFailOutput_TextBox.Text += ex.Message
            writeIntoErrorInfoTxt($"[Error] 名稱管理員:{spec}   {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' 將radbox有打勾的項目和usr(輸入資料)寫入msExcel_workbook(目標excel)的spec(名稱管理員)
    ''' </summary>
    ''' <param name="usr"></param>
    ''' <param name="spec"></param>
    ''' <param name="radbox"></param>
    ''' <param name="msExcel_workbook"></param>
    Overloads Sub excelWriteIn(usr As String, spec As String, radbox As RadioButton, msExcel_workbook As Excel.Workbook)
        Try
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
        Catch ex As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"[Error] 名稱管理員:{spec}"
            JobMaker_Form.ResultFailOutput_TextBox.Text += ex.Message
            writeIntoErrorInfoTxt($"[Error] 名稱管理員:{spec}   {ex.Message}")
        End Try
    End Sub
    ''' <summary>
    ''' 沒打勾的CheckBox寫入資料
    ''' </summary>
    ''' <param name="usr"> 輸入資料 </param>
    ''' <param name="spec"> 名稱管理員 </param>
    ''' <param name="chkbox">未打勾的CheckBox</param>
    ''' <param name="msExcel_workbook"></param>
    Sub excelWriteIn_ForReverseState(usr As String, spec As String, chkbox As CheckBox, msExcel_workbook As Excel.Workbook)
        Try
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
        Catch ex As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"[Error] 名稱管理員:{spec}"
            JobMaker_Form.ResultFailOutput_TextBox.Text += ex.Message
            writeIntoErrorInfoTxt($"[Error] 名稱管理員:{spec}   {ex.Message}")
        End Try
    End Sub

    '寫入excel內的方法
    ''' <summary>
    ''' 將 名稱管理員chkboxName(為圖形) 寫入workbook中的sheetPageName分頁名稱中
    ''' </summary>
    ''' <param name="chkboxName"> excel中checkBox的圖形名稱 </param>
    ''' <param name="sheetPageName"> excel中分頁名稱 </param>
    ''' <param name="msExcel_workbook"> workbook名稱 </param>
    Overloads Sub chkboxWriteIn(chkboxName As String, sheetPageName As String, msExcel_workbook As Excel.Workbook)
        Try
            If chkboxName IsNot "" Then
                msExcel_workbook.Sheets(sheetPageName).CheckBoxes(chkboxName).value = True

                JobMaker_Form.ResultOutput_TextBox.Text += ($"圖形名稱:{chkboxName} / 分頁名稱:{sheetPageName} 打勾成功{vbCrLf}")
            Else
                JobMaker_Form.ResultOutput_TextBox.Text += ($"圖形名稱:{chkboxName} / 分頁名稱:{sheetPageName} 打勾失敗{vbCrLf}")
            End If
        Catch ex As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"[Error] 圖形名稱:{chkboxName}"
            JobMaker_Form.ResultFailOutput_TextBox.Text += ex.Message
            writeIntoErrorInfoTxt($"[Error] 圖形名稱:{chkboxName}   {ex.Message}")
        End Try
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
    Overloads Function chkBoxStateRead(rdbtn As CheckBox, chk_draw_name As String) As String
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

End Class
