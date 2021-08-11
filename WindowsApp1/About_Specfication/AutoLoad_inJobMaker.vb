Imports Microsoft.Office.Interop

Module AutoLoad_inJobMaker

    ''' <summary>
    ''' ㄧ頁標準55列
    ''' </summary>
    Const std_page_row As Integer = 55

    ''' <summary>
    ''' 共三頁Key1~Key3
    ''' </summary>
    Const std_page_num As Integer = 3

    ''' <summary>
    ''' 取得分頁名字
    ''' </summary>
    Dim sheetName As String()
    ''' <summary>
    ''' 取得X台號機
    ''' </summary>
    Dim liftNum As Integer
    ''' <summary>
    ''' 儲存號機名 例如:#1~#4
    ''' </summary>
    Dim liftNo_Array As String() = {""}


    Private Sub getSheetName(msExcel_workbook As Excel.Workbook)
        Dim sheet_count As String = msExcel_workbook.Worksheets.Count

        ReDim sheetName(sheet_count - 1)
        For ws As Integer = 1 To sheet_count
            sheetName(ws - 1) = msExcel_workbook.Sheets(ws).Name
        Next
    End Sub
    ''' <summary>
    ''' 取得號機數量
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    Private Sub getLiftNum(msExcel_workbook As Excel.Workbook)
        Dim row_add, col_add As Integer
        Dim row_current, col_current As Integer '目前ROW & COL

        For Each ws In sheetName
            Do
                row_current = 1 + row_add
                col_current = 1 + col_add
                '從Cell(1,1)開始往下讀取，如果超過五列是空值就往下一欄讀取
                row_add += 1
                If row_add > 5 And getMathOnExcel.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current) = "" Then
                    col_add += 1
                    row_add = 1
                    row_current = 1
                End If

                If getMathOnExcel.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current) = " 號機" Then
                    Dim title_colName_afterConvert, content_colName_afterConvert As String
                    '取得Title轉換成文字的欄位名稱，例如:號機
                    title_colName_afterConvert = getMathOnExcel.convertColumn_fromIntToString(col_current)
                    Dim title_mergeNum, content_mergeNum As Integer
                    '取得Title目前合併儲存格的欄位數量，例如:號機有兩欄
                    title_mergeNum = getMathOnExcel.getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, ws, row_current, title_colName_afterConvert)
                    '往右加到內容處，例如:號機往右+2，變成號機名
                    col_current = col_current + title_mergeNum

                    '取得Content轉換成文字的欄位名稱，例如:TW-1111~TW-XXXX
                    content_colName_afterConvert = getMathOnExcel.convertColumn_fromIntToString(col_current)
                    content_mergeNum = getMathOnExcel.getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, ws, row_current, content_colName_afterConvert)

                    Do
                        '#1~#X
                        If getMathOnExcel.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current + 1, col_current) <> "" Then
                            ReDim Preserve liftNo_Array(liftNum)
                            liftNo_Array(liftNum) =
                                getMathOnExcel.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current + 1, col_current)
                            liftNum += 1
                        Else
                            Exit Do
                        End If
                        col_current += content_mergeNum
                    Loop
                End If

            Loop Until row_current > std_page_row Or liftNum > 0
        Next
    End Sub


    Private Sub setValueToTextbox(msExcel_workbook As Excel.Workbook, mTitleText As String, mTextBox As Control)
        Dim row_add, col_add As Integer
        Dim row_current, col_current As Integer '目前ROW & COL的位置

        '儲存每一個title的內容，例如:所在場地的內容是 台灣 有 4台，則 allContent_ofEachTitle_forCompare(0~3)="台灣"
        Dim allContent_ofEachTitle_forCompare(liftNum - 1) As String

        Dim allContent_ofEachTitle_forOutput As String() = {""} '儲存每一個開頭的號機內容，例如:現場所在地->#1:台灣 #2-4:香港的 {台灣,香港}
        Dim sameLiftContent_forOutput As String() = {""} '儲存相同號機的內容，例如:#1:台灣 #2-4:香港 的 {#1,#2-4}
        For Each ws In sheetName
            For num_i As Integer = 1 To std_page_num
                col_add = 1
                Do
                    row_current = 1 + (std_page_row * (num_i - 1)) + row_add
                    col_current = 1 + col_add
                    '從Cell(1,1)開始往下讀取，如果超過五列是空值就往下一欄讀取
                    row_add += 1
                    If row_add > 5 And getMathOnExcel.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current) = "" Then
                        col_add += 1
                        row_add = 1
                        row_current = 1
                    End If

                    If getMathOnExcel.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current) = mTitleText Then
                        Dim title_colName_afterConvert, content_colName_afterConvert As String
                        '取得Title轉換成文字的欄位名稱，例如:號機
                        title_colName_afterConvert = getMathOnExcel.convertColumn_fromIntToString(col_current)
                        Dim title_mergeNum, content_mergeNum As Integer
                        '取得Title目前合併儲存格的欄位數量，例如:號機有兩欄
                        title_mergeNum = getMathOnExcel.getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, ws, row_current, title_colName_afterConvert)
                        '往右加到內容處，例如:號機往右+2，變成號機名
                        col_current = col_current + title_mergeNum

                        '取得Content轉換成文字的欄位名稱，例如:Tw-1111
                        content_colName_afterConvert = getMathOnExcel.convertColumn_fromIntToString(col_current)
                        content_mergeNum = getMathOnExcel.getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, ws, row_current, content_colName_afterConvert)

                        Dim output_count As Integer
                        For i As Integer = 1 To liftNum
                            allContent_ofEachTitle_forCompare(i - 1) = getMathOnExcel.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current)
                            col_current += content_mergeNum

                            For Each output In allContent_ofEachTitle_forOutput
                                If allContent_ofEachTitle_forCompare(i - 1).Equals(output) Then
                                    For same_i As Integer = 1 To allContent_ofEachTitle_forOutput.Count
                                        If allContent_ofEachTitle_forOutput(same_i - 1).Equals(output) Then
                                            sameLiftContent_forOutput(same_i - 1) += $"{liftNo_Array(i - 1)},"
                                        End If
                                    Next
                                    Exit For
                                Else
                                    ReDim Preserve allContent_ofEachTitle_forOutput(output_count)
                                    ReDim Preserve sameLiftContent_forOutput(output_count)
                                    allContent_ofEachTitle_forOutput(output_count) = allContent_ofEachTitle_forCompare(i - 1)
                                    sameLiftContent_forOutput(output_count) = $"{liftNo_Array(i - 1)},"
                                    output_count += 1
                                End If
                            Next
                        Next

                        For output As Integer = 1 To allContent_ofEachTitle_forOutput.Count
                            If output < allContent_ofEachTitle_forOutput.Count Then
                                mTextBox.Text +=
                                    $"{sameLiftContent_forOutput(output - 1)} : {allContent_ofEachTitle_forOutput(output - 1)} / "
                            Else
                                mTextBox.Text +=
                                    $"{sameLiftContent_forOutput(output - 1)} : {allContent_ofEachTitle_forOutput(output - 1)}"
                            End If
                        Next

                        Exit Sub
                    End If
                Loop Until row_current > num_i * std_page_row
            Next
        Next
    End Sub

    Public Sub readData_fromExcel(msExcel_workbook As Excel.Workbook)

        'getSheetName(msExcel_workbook)
        'getLiftNum(msExcel_workbook)

        'Dim eepData_textbox As TextBox =
        '    {JobMaker_Form.EepData_MachineRoom_TextBox, JobMaker_Form.EepData_Speed_TextBox,
        '     JobMaker_Form.EepData_Capactity_TextBox, JobMaker_Form.EepData_TopFL_TextBox}
        'setValueToTextbox(msExcel_workbook, , eepData_textbox)
        'For Each mTabCtrl As Control In JobMaker_Form.EepData_TabControl.Controls
        '    For Each grpCtrl As Control In mTabCtrl.Controls
        '        For Each Ctrl As Control In grpCtrl.Controls
        '            If TypeOf (Ctrl) Is TextBox Then
        '                'MsgBox(Ctrl.Name)
        '            End If
        '        Next
        '    Next
        'Next

        For Each Ctrl As Control In JobMaker_Form.EepData_Page1_GroupBox.Controls
            If TypeOf (Ctrl) Is TextBox Then
                MsgBox(Ctrl.Name)
            End If
        Next
    End Sub


End Module
