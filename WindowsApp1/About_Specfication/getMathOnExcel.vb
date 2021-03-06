Imports Microsoft.Office.Interop

Module getMathOnExcel
    ''' <summary>
    ''' 取得目標Cell的值
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="ws">worksheet name</param>
    ''' <param name="row">row</param>
    ''' <param name="col">column</param>
    ''' <returns></returns>
    Public Function getValue_byRowCol_fromWorksheet(msExcel_workbook As Excel.Workbook, ws As String,
                                                   row As Integer, col As Integer) As String
        getValue_byRowCol_fromWorksheet = ""
        Try
            Return msExcel_workbook.Worksheets(ws).cells(row, col).value
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getValue_byRowCol_fromWorksheet",
                                              $"<{ws}>分頁，取<Cell({row},{col})>的值時發生錯誤", ex)
        End Try

    End Function
    ''' <summary>
    ''' 取得指定 nameManager(名稱管理員) 的 value(值) , i.g 名稱管理員叫 Apple 取得值 蘋果
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <returns></returns>
    Public Function getValue_fromNameManager(msExcel_workbook As Excel.Workbook, nameManager As String) As String
        getValue_fromNameManager = ""
        Try
            Return msExcel_workbook.Names.Item(nameManager).RefersToRange.Value
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getValue_fromNameManager",
                                              $"<{nameManager}>取值時發生錯誤", ex)
        End Try
    End Function
    ''' <summary>
    ''' 取得指定 nameManager(名稱管理員) 的 column(欄) , 類型為Cell , i.g 名稱管理員叫 Apple 欄位在(1,2) 回傳1
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <returns></returns>
    Public Function getCol_fromNameManager_typeIsCell(msExcel_workbook As Excel.Workbook, nameManager As String) As Integer
        getCol_fromNameManager_typeIsCell = 0
        Try
            Return msExcel_workbook.Names.Item(nameManager).RefersToRange.Column '欄
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getCol_fromNameManager_typeIsCell",
                                             $"<{nameManager}>取欄位(column)時發生錯誤", ex)
        End Try
    End Function
    ''' <summary>
    ''' 取得指定 nameManager(名稱管理員) 的 row(列) , 類型為Cell , i.g 名稱管理員叫 Apple 欄位在(1,2) 回傳1
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <returns></returns>
    Public Function getRow_fromNameManager_typeIsCell(msExcel_workbook As Excel.Workbook, nameManager As String) As Integer
        getRow_fromNameManager_typeIsCell = 0
        Try
            Return msExcel_workbook.Names.Item(nameManager).RefersToRange.Row '列
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getRow_fromNameManager_typeIsCell",
                                              $"<{nameManager}>取列位(Row)時發生錯誤", ex)
        End Try
    End Function

    ''' <summary>
    ''' Return Col number from NameManager on Excel , type is Range . ig A4 > Return to 4 (欄)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager"></param>
    ''' <returns></returns>
    Public Function getCol_fromNameManager_typeIsRange(msExcel_workbook As Excel.Workbook, nameManager As String) As String
        getCol_fromNameManager_typeIsRange = ""
        Try
            Return msExcel_workbook.Names.Item(nameManager).RefersToRange.AddressLocal(False, False).Last.ToString()
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getRow_fromNameManager_typeIsCell",
                                              $"<{nameManager}>取欄位(Col)時發生錯誤(A4 > Return to 4 (欄))", ex)
        End Try
    End Function
    ''' <summary>
    ''' Return Row number from NameManager on Excel , type is Range . ig A4 > Return to A (列)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager"></param>
    ''' <returns></returns>
    Public Function getRow_fromNameManager_typeIsRange(msExcel_workbook As Excel.Workbook, nameManager As String) As String
        getRow_fromNameManager_typeIsRange = ""
        Try
            Return msExcel_workbook.Names.Item(nameManager).RefersToRange.AddressLocal(False, False).First.ToString()
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getRow_fromNameManager_typeIsRange",
                                              $"將<{nameManager}>取列位(Row)時發生錯誤(A4 > Return to A (列))", ex)
        End Try
    End Function

    ''' <summary>
    ''' Return 合併儲存格的Row列數量 , 類型為Range
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="msExcel_worksheetsName"></param>
    ''' <param name="range_row">row類型為Range</param>
    ''' <param name="range_col">column類型為Range</param>
    ''' <returns></returns>
    Public Function getRowCount_ifRangeIsMerge_onWorkShts(msExcel_workbook As Excel.Workbook,
                                                          msExcel_worksheetsName As String,
                                                          range_row As String, range_col As String) As Integer
        getRowCount_ifRangeIsMerge_onWorkShts = 0
        Try
            Return msExcel_workbook.Worksheets(msExcel_worksheetsName).range(range_col & range_row).MergeArea.Rows.Count

        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getRowCount_ifRangeIsMerge_onWorkShts",
                                              $"取得 合併儲存格的<{range_row}>Row列的數量 時發生錯誤", ex)

        End Try
    End Function
    ''' <summary>
    ''' Return 合併儲存格的col列數量 , 類型為Range
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="msExcel_worksheetsName"></param>
    ''' <param name="range_row">row類型為Range</param>
    ''' <param name="range_col">column類型為Range</param>
    ''' <returns></returns>
    Public Function getColCount_ifRangeIsMerge_onWorkShts(msExcel_workbook As Excel.Workbook,
                                                          msExcel_worksheetsName As String,
                                                          range_row As String, range_col As String) As Integer
        getColCount_ifRangeIsMerge_onWorkShts = 0
        Try

            Return msExcel_workbook.Worksheets(msExcel_worksheetsName).range(range_col & range_row).MergeArea.Columns.Count
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getColCount_ifRangeIsMerge_onWorkShts",
                                              $"取得合併儲存格的<{range_col}>col列的數量 時發生錯誤", ex)

        End Try
    End Function

    ''' <summary>
    ''' 取得目前 nameManager(名稱管理員)所在的 分頁名稱
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <returns></returns>
    Public Function getWorksheetName_fromNameManager(msExcel_workbook As Excel.Workbook, nameManager As String) As String
        getWorksheetName_fromNameManager = ""
        Try
            Return msExcel_workbook.Names.Item(nameManager).RefersToRange.Worksheet.Name
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.getWorksheetName_fromNameManager",
                                              $"取得目前 <{nameManager}>(名稱管理員)所在的 分頁名稱 時發生錯誤", ex)

        End Try
    End Function

    ''' <summary>
    ''' 為 nameManager(目標儲存格) 設定 value(值)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <param name="value">要輸入的值</param>
    Public Sub setValue_to_Cells_onWorksht(msExcel_workbook As Excel.Workbook,
                                           nameManager As String,
                                           value As String)
        Dim mWorksheet_Name As String = getWorksheetName_fromNameManager(msExcel_workbook, nameManager)
        Dim nameManager_Row As Integer = getRow_fromNameManager_typeIsCell(msExcel_workbook, nameManager)
        Dim nameManager_Col As Integer = getCol_fromNameManager_typeIsCell(msExcel_workbook, nameManager)

        Try
            msExcel_workbook.Worksheets(mWorksheet_Name).Cells(nameManager_Row, nameManager_Col).Value = value
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.setValue_to_Cells_onWorksht",
                                              $"為 <{nameManager}>(目標儲存格) 設定 value(值) 時發生錯誤", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 為 nameManager(目標儲存格) 設定 value(值)，並依序往下(列)儲存 > Row + i
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <param name="i">loop</param>
    ''' <param name="mValue">要輸入的值</param>
    Public Sub setValue_to_Cells_addBelow_onWorksht(msExcel_workbook As Excel.Workbook,
                                                    nameManager As String,
                                                    i As Integer,
                                                    mValue As String)
        Dim mWorksheet_Name As String = getWorksheetName_fromNameManager(msExcel_workbook, nameManager)
        Dim nameManager_Row As Integer = getRow_fromNameManager_typeIsCell(msExcel_workbook, nameManager)
        Dim nameManager_Col As Integer = getCol_fromNameManager_typeIsCell(msExcel_workbook, nameManager)
        Try
            msExcel_workbook.Worksheets(mWorksheet_Name).Cells(nameManager_Row + i, nameManager_Col).Value = mValue
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.setValue_to_Cells_addBelow_onWorksht",
                                              $"為 <{nameManager}>(目標儲存格) 設定 value(值)，並依序往下(列)儲存 > Row + i 時發生錯誤", ex)

        End Try

    End Sub

    ''' <summary>
    ''' 從nameManager取得當前worksheet name，並依照Cell(row,col)設定value
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager"></param>
    ''' <param name="mRow"></param>
    ''' <param name="mCol"></param>
    ''' <param name="mValue"></param>
    Public Sub setValue_to_RowCol_onWorksht(msExcel_workbook As Excel.Workbook,
                                            nameManager As String,
                                            mRow As Integer, mCol As Integer,
                                            mValue As String)
        Dim mWorksheet_Name As String = getWorksheetName_fromNameManager(msExcel_workbook, nameManager)
        Try
            msExcel_workbook.Worksheets(mWorksheet_Name).Cells(mRow, mCol).Value = mValue
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.setValue_to_RowCol_onWorksht",
                                              $"為 <Cell{mRow},{mCol}>(目標儲存格) 設定 value(值) 發生錯誤", ex)

        End Try

    End Sub
    ''' <summary>
    ''' 為 nameManager(目標儲存格) 設定 value(值)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager"></param>
    ''' <param name="mValue"></param>
    Public Sub setValue_to_nameManager_onWorksht(msExcel_workbook As Excel.Workbook,
                                                 nameManager As String,
                                                 mValue As String)
        Try
            msExcel_workbook.Names.Item(nameManager).RefersToRange.Cells.Value = mValue
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.setValue_to_nameManager_onWorksht",
                                              $"為 <{nameManager}>(目標儲存格) 設定 value(值)時發生錯誤", ex)

        End Try
    End Sub

    ''' <summary>
    ''' 每隔兩個項目就變更Cell顏色
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="current_Row">目前Row</param>
    ''' <param name="title_Col">標題Col</param>
    ''' <param name="item_number">目前第n項</param>
    Public Sub ChangeRangeColor_FinalCheck_onExcel(msExcel_workbook As Excel.Workbook,
                                                   current_Row As Integer, title_Col As Integer,
                                                   item_number As Integer)

        If item_number Mod 2 = 0 Then
            Dim get_NameManager As Spec_NameManager = New Spec_NameManager
            setColor_toRange_onWorkShts(msExcel_workbook,
                                        get_NameManager.FinalCheck_Item,
                                        current_Row, title_Col,
                                        RGB(208, 240, 255))
            setColor_toRange_onWorkShts(msExcel_workbook,
                                        get_NameManager.FinalCheck_Spec,
                                        current_Row, title_Col + 1,
                                        RGB(208, 240, 255))
            setColor_toRange_onWorkShts(msExcel_workbook,
                                        get_NameManager.FinalCheck_State,
                                        current_Row, title_Col + 2,
                                        RGB(208, 240, 255))
        End If
    End Sub


    ''' <summary>
    ''' 將 column(欄)從數字轉換成英文 , i.g 38欄 > AL欄
    ''' </summary>
    ''' <param name="getCol_fromInt">傳入column(欄),類型為 int(整數)</param>
    ''' <returns></returns>
    Public Function convertColumn_fromIntToString(getCol_fromInt As Integer) As String
        convertColumn_fromIntToString = ""
        Try
            Dim modulo As Integer
            While getCol_fromInt > 0
                modulo = (getCol_fromInt - 1) Mod 26
                convertColumn_fromIntToString = Convert.ToChar(65 + modulo).ToString() + convertColumn_fromIntToString
                getCol_fromInt = CInt((getCol_fromInt - modulo) / 26)
            End While

            Return convertColumn_fromIntToString
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.convertColumn_fromIntToString",
                                              $"將 <{getCol_fromInt}> column(欄)從數字轉換成英文 時發生錯誤", ex)
        End Try
    End Function

    ''' <summary>
    ''' 在指定的nameManager(spec)加上刪除線
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="spec">指定的nameManager</param>
    Public Sub strikeThrough_allText_onWorkSht(msExcel_workbook As Excel.Workbook, spec As String)
        Try
            msExcel_workbook.Names.Item(spec).RefersToRange.Cells.Font.Strikethrough = True
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.strikeThrough_allText_onWorkSht",
                                              $"<{spec}>文字刪除線 添加 時發生錯誤", ex)

        End Try
    End Sub
    ''' <summary>
    ''' 在指定的nameManager(spec)去除刪除線
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="spec">指定的nameManager</param>
    Public Sub NotStrikeThrough_allText_onWorkSht(msExcel_workbook As Excel.Workbook, spec As String)
        Try
            msExcel_workbook.Names.Item(spec).RefersToRange.Cells.Font.Strikethrough = False
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.NotStrikeThrough_allText_onWorkSht",
                                              $"<{spec}>文字刪除線 取消 時發生錯誤", ex)

        End Try
    End Sub

    ''' <summary>
    ''' 在指定的nameManager(spec)的全部文字(allString)內指定文字(partString)加上刪除線
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="spec"></param>
    ''' <param name="partString"></param>
    ''' <param name="allString"></param>
    Public Sub strikeThrough_partText_onWorkSht(msExcel_workbook As Excel.Workbook, spec As String,
                                                partString As String, allString As String)
        msExcel_workbook.Names.Item(spec).RefersToRange.Characters(InStr(allString, partString), Len(partString)
                                    ).Font.Strikethrough = True
    End Sub

    ''' <summary>
    ''' 取得Worksheet name並設定該Range的顏色
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="spec_name"></param>
    ''' <param name="range_row"></param>
    ''' <param name="range_col"></param>
    ''' <param name="setColor"></param>
    ''' <returns></returns>
    Public Function setColor_toRange_onWorkShts(msExcel_workbook As Excel.Workbook,
                                                spec_name As String,
                                                range_row As Integer, range_col As Integer,
                                                setColor As Integer) As Integer
        setColor_toRange_onWorkShts = 0
        Try
            Dim msWorkSheetName As String = getWorksheetName_fromNameManager(msExcel_workbook, spec_name)
            Dim col As String = convertColumn_fromIntToString(range_col)
            msExcel_workbook.Worksheets(msWorkSheetName).range(col & CStr(range_row)).Interior.Color = setColor
        Catch ex As Exception
            errorInfo.writeInfoError_errorMsg($"getMathOnExcel.setColor_toRange_onWorkShts",
                                              $"設定Excel Range Color <{spec_name}> 時發生錯誤", ex)

        End Try
    End Function
End Module
