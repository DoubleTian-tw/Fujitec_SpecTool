Imports System.IO

Module errorInfo
    Dim errorTxtPath As String = $"{Application.StartupPath}\errorInfo.txt"

    ''' <summary>
    ''' create error info檔案
    ''' </summary>
    Public Sub createError_InfoTxt(msg As String)
        'Create errorInof.txt if it doesn't exist
        If File.Exists(errorTxtPath) = False Then
            Dim fs As FileStream = File.Create(errorTxtPath)
        End If

        'Add first error msg to the txt file
        Using ws As StreamWriter = File.AppendText(errorTxtPath)
            ws.WriteLine($"{Now} ==== {msg} ====")
        End Using
    End Sub

    ''' <summary>
    ''' 追加Error info類型的文字
    ''' </summary>
    ''' <param name="title">錯誤標題</param>
    Public Sub writeTitleIntoError_InfoTxt(title As String)
        Using ws As StreamWriter = File.AppendText(errorTxtPath)
            ws.WriteLine($"{Now} ==== {title} ====")
        End Using
    End Sub

    ''' <summary>
    ''' 追加Error info文字
    ''' </summary>
    ''' <param name="msg">錯誤訊息</param>
    Public Sub writeIntoError_InfoTxt(msg As String)
        'Add first error msg to the txt file
        Using ws As StreamWriter = File.AppendText(errorTxtPath)
            ws.WriteLine(msg)
        End Using
    End Sub
End Module
