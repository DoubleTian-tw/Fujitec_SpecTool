Imports Microsoft.Office.Interop



Public Class AutoLoad_intoJobMaker

    Const std_page_row As Integer = 55 'ㄧ頁標準55列
    Const std_page_num As Integer = 3 'ㄧ頁標準55列


    Public Sub readData_fromExcel(msExcel_workbook As Excel.Workbook)
        Dim excelMath As getMath_onExcel = New getMath_onExcel

        Dim sheet_count As String = msExcel_workbook.Worksheets.Count
        Dim sheetName(sheet_count - 1) As String
        For ws As Integer = 1 To sheet_count
            sheetName(ws - 1) = msExcel_workbook.Sheets(ws).Name
            'MsgBox(sheetName(ws))
        Next


        Dim row_add, col_add As Integer
        Dim row_current, col_current As Integer '目前ROW & COL
        Dim liftNum As Integer '取得X台號機
        Dim liftNo_Array As String() = {""} '儲存號機名 例如:#1~#4
        For Each ws In sheetName
            Do
                row_current = 1 + row_add
                col_current = 1 + col_add
                '從Cell(1,1)開始往下讀取，如果超過五列是空值就往下一欄讀取
                row_add += 1
                If row_add > 5 And excelMath.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current) = "" Then
                    col_add += 1
                    row_add = 1
                    row_current = 1
                End If

                If excelMath.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current) = " 號機" Then
                    Dim title_colName_afterConvert, content_colName_afterConvert As String
                    '取得Title轉換成文字的欄位名稱，例如:號機
                    title_colName_afterConvert = excelMath.convertColumn_fromIntToString(col_current)
                    Dim title_mergeNum, content_mergeNum As Integer
                    '取得Title目前合併儲存格的欄位數量，例如:號機有兩欄
                    title_mergeNum = excelMath.getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, ws, row_current, title_colName_afterConvert)
                    '往右加到內容處，例如:號機往右+2，變成號機名
                    col_current = col_current + title_mergeNum

                    '取得Content轉換成文字的欄位名稱，例如:TW-1111~TW-XXXX
                    content_colName_afterConvert = excelMath.convertColumn_fromIntToString(col_current)
                    content_mergeNum = excelMath.getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, ws, row_current, content_colName_afterConvert)

                    Do
                        '#1~#X
                        If excelMath.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current + 1, col_current) <> "" Then
                            ReDim Preserve liftNo_Array(liftNum)
                            liftNo_Array(liftNum) =
                                excelMath.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current + 1, col_current)
                            liftNum += 1
                        Else
                            Exit Do
                        End If
                        col_current += content_mergeNum
                    Loop
                End If

                'If liftNum > 0 Then
                '    Exit Do
                'End If
            Loop Until row_current > std_page_row Or liftNum > 0
        Next

        '儲存每一個title的內容，例如:所在場地的內容是 台灣 * 4台，則 allContent_ofEachTitle(0..3)=台灣
        Dim allContent_ofEachTitle_forCompare(liftNum - 1) As String

        Dim allContent_ofEachTitle_forOutput As String() = {""}
        Dim sameLiftContent_forOutput As String() = {""}
        For Each ws In sheetName
            For num_i As Integer = 1 To std_page_num
                col_add = 1
                Do
                    row_current = 1 + (std_page_row * (num_i - 1)) + row_add
                    col_current = 1 + col_add
                    '從Cell(1,1)開始往下讀取，如果超過五列是空值就往下一欄讀取
                    row_add += 1
                    If row_add > 5 And excelMath.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current) = "" Then
                        col_add += 1
                        row_add = 1
                        row_current = 1
                    End If

                    'MsgBox(excelMath.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current))
                    If excelMath.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current) = "定格速度(m/min)" Then
                        Dim title_colName_afterConvert, content_colName_afterConvert As String
                        '取得Title轉換成文字的欄位名稱，例如:號機
                        title_colName_afterConvert = excelMath.convertColumn_fromIntToString(col_current)
                        Dim title_mergeNum, content_mergeNum As Integer
                        '取得Title目前合併儲存格的欄位數量，例如:號機有兩欄
                        title_mergeNum = excelMath.getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, ws, row_current, title_colName_afterConvert)
                        '往右加到內容處，例如:號機往右+2，變成號機名
                        col_current = col_current + title_mergeNum

                        '取得Content轉換成文字的欄位名稱，例如:Tw-1111
                        content_colName_afterConvert = excelMath.convertColumn_fromIntToString(col_current)
                        content_mergeNum = excelMath.getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook, ws, row_current, content_colName_afterConvert)

                        Dim output_count As Integer
                        For i As Integer = 1 To liftNum
                            allContent_ofEachTitle_forCompare(i - 1) = excelMath.getValue_byRowCol_formWorksheet(msExcel_workbook, ws, row_current, col_current)
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
                                JobMaker_Form.EepData_Speed_TextBox.Text +=
                                    $"{sameLiftContent_forOutput(output - 1)} : {allContent_ofEachTitle_forOutput(output - 1)} / "
                            Else
                                JobMaker_Form.EepData_Speed_TextBox.Text +=
                                    $"{sameLiftContent_forOutput(output - 1)} : {allContent_ofEachTitle_forOutput(output - 1)}"
                            End If
                        Next


                    End If
                Loop Until row_current > num_i * std_page_row
            Next
        Next
    End Sub
End Class
