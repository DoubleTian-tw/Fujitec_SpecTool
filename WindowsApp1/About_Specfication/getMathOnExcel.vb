Imports Microsoft.Office.Interop
Module getMathOnExcel
    Public Function getValue_byRowCol_formWorksheet(msExcel_workbook As Excel.Workbook, ws As String,
                                                   row As Integer, col As Integer) As String
        Return msExcel_workbook.Worksheets(ws).cells(row, col).value
    End Function
    ''' <summary>
    ''' 取得指定 nameManager(名稱管理員) 的 value(值) , i.g 名稱管理員叫 Apple 取得值 蘋果
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <returns></returns>
    Public Function getValue_formNameManager(msExcel_workbook As Excel.Workbook, nameManager As String) As String
        Return msExcel_workbook.Names.Item(nameManager).RefersToRange.Value
    End Function
    ''' <summary>
    ''' 取得指定 nameManager(名稱管理員) 的 column(欄) , 類型為Cell , i.g 名稱管理員叫 Apple 欄位在(1,2) 回傳1
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <returns></returns>
    Public Function getCol_fromNameManager_typeIsCell(msExcel_workbook As Excel.Workbook, nameManager As String) As Integer
        Return msExcel_workbook.Names.Item(nameManager).RefersToRange.Column '欄
    End Function
    ''' <summary>
    ''' 取得指定 nameManager(名稱管理員) 的 row(列) , 類型為Cell , i.g 名稱管理員叫 Apple 欄位在(1,2) 回傳1
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <returns></returns>
    Public Function getRow_fromNameManager_typeIsCell(msExcel_workbook As Excel.Workbook, nameManager As String) As Integer
        Return msExcel_workbook.Names.Item(nameManager).RefersToRange.Row '列
    End Function

    ''' <summary>
    ''' Return Col number from NameManager on Excel , type is Range . ig A4 > Return to 4 (欄)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager"></param>
    ''' <returns></returns>
    Public Function getCol_fromNameManager_typeIsRange(msExcel_workbook As Excel.Workbook, nameManager As String) As String
        Return msExcel_workbook.Names.Item(nameManager).RefersToRange.AddressLocal(False, False).Last.ToString()
    End Function
    ''' <summary>
    ''' Return Row number from NameManager on Excel , type is Range . ig A4 > Return to A (列)
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager"></param>
    ''' <returns></returns>
    Public Function getRow_fromNameManager_typeIsRange(msExcel_workbook As Excel.Workbook, nameManager As String) As String
        Return msExcel_workbook.Names.Item(nameManager).RefersToRange.AddressLocal(False, False).First.ToString()
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
        Return msExcel_workbook.Worksheets(msExcel_worksheetsName).range(range_col & range_row).MergeArea.Rows.Count
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
        Return msExcel_workbook.Worksheets(msExcel_worksheetsName).range(range_col & range_row).MergeArea.Columns.Count
    End Function

    ''' <summary>
    ''' 取得目前 nameManager(名稱管理員)所在的 分頁名稱
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <returns></returns>
    Public Function getWorksheetName_fromNameManager(msExcel_workbook As Excel.Workbook, nameManager As String) As String
        Return msExcel_workbook.Names.Item(nameManager).RefersToRange.Worksheet.Name
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

        msExcel_workbook.Worksheets(mWorksheet_Name).Cells(nameManager_Row, nameManager_Col).Value = value
    End Sub

    ''' <summary>
    ''' 為 nameManager(目標儲存格) 設定 value(值)，並依序往下(列)儲存 > Row + i
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="nameManager">名稱管理員</param>
    ''' <param name="i">loop</param>
    ''' <param name="value">要輸入的值</param>
    Public Sub setValue_to_Cells_addBelow_onWorksht(msExcel_workbook As Excel.Workbook,
                                                    nameManager As String,
                                                    i As Integer,
                                                    value As String)
        Dim mWorksheet_Name As String = getWorksheetName_fromNameManager(msExcel_workbook, nameManager)
        Dim nameManager_Row As Integer = getRow_fromNameManager_typeIsCell(msExcel_workbook, nameManager)
        Dim nameManager_Col As Integer = getCol_fromNameManager_typeIsCell(msExcel_workbook, nameManager)

        msExcel_workbook.Worksheets(mWorksheet_Name).Cells(nameManager_Row + i, nameManager_Col).Value = value

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
        msExcel_workbook.Names.Item(nameManager).RefersToRange.Cells.Value = mValue
    End Sub
    ''' <summary>
    ''' 將 column(欄)從數字轉換成英文 , i.g 38欄 > AL欄
    ''' </summary>
    ''' <param name="getCol_fromInt">傳入column(欄),類型為 int(整數)</param>
    ''' <returns></returns>
    Public Function convertColumn_fromIntToString(getCol_fromInt As Integer) As String
        Dim modulo As Integer

        While getCol_fromInt > 0
            modulo = (getCol_fromInt - 1) Mod 26
            convertColumn_fromIntToString = Convert.ToChar(65 + modulo).ToString() + convertColumn_fromIntToString
            getCol_fromInt = CInt((getCol_fromInt - modulo) / 26)
        End While

        Return convertColumn_fromIntToString
    End Function

    ''' <summary>
    ''' 在指定的nameManager(spec)加上刪除線
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="spec">指定的nameManager</param>
    Public Sub strikeThrough_allText_onWorkSht(msExcel_workbook As Excel.Workbook, spec As String)
        msExcel_workbook.Names.Item(spec).RefersToRange.Cells.Font.Strikethrough = True
    End Sub
    ''' <summary>
    ''' 在指定的nameManager(spec)去除刪除線
    ''' </summary>
    ''' <param name="msExcel_workbook"></param>
    ''' <param name="spec">指定的nameManager</param>
    Public Sub NotStrikeThrough_allText_onWorkSht(msExcel_workbook As Excel.Workbook, spec As String)
        msExcel_workbook.Names.Item(spec).RefersToRange.Cells.Font.Strikethrough = False
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
End Module
