Module ProgramAllName
    Public Const fileName_mainProgram As String = "MagicTool"
    Public Const fileName_updateProgram As String = "Update_magicTool"
    Public Const fileName_updateNoticeFile As String = "File_Update_Notice"
    Public Const fileName_SetFileIni As String = "SetFile.ini"
    Public Const fileName_SetFileIniBat As String = "SettingINI.bat"
    Public Const fileName_Manualpptx As String = "Manual.pptx"
    Public Const fileName_ErrorInfo As String = "errorInfo.txt"
    Public Const SQLite_ToolDBMS_Name As String = "Tool_Database.sqlite"
    Public Const SQLite_StdJobDataDBMS_Name As String = "Standard_StoredJobData.sqlite"

    ''' <summary>
    ''' 回傳組件名稱
    ''' </summary>
    ''' <returns></returns>
    Public Function get_assemblyName() As String
        Dim assemblyFullName As String
        Dim fullNameArr() As String
        assemblyFullName = GetType(MagicTool).Assembly.FullName
        fullNameArr = assemblyFullName.Split(",")
        If fullNameArr IsNot Nothing Then
            Return fullNameArr(0)
        Else
            Return ""
        End If
    End Function
End Module
