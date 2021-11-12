Imports System.IO
'Imports Microsoft.Office.Interop
Module errorInfo
    Dim errorTxtPath As String = $"{Application.StartupPath}\errorInfo.txt"

    ''' <summary>
    ''' create error info檔案
    ''' </summary>
    Public Sub createError_InfoTxt(title As String)
        Try
            'Create errorInof.txt if it doesn't exist
            If File.Exists(errorTxtPath) = False Then
                Dim fs As FileStream = File.Create(errorTxtPath)
            End If

            'Add first error msg to the txt file
            Using ws As StreamWriter = File.AppendText(errorTxtPath)
                ws.WriteLine($"{Now} ==== {title} ===={vbCrLf}")
            End Using
        Catch ex As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"=== errorInfo.createError_InfoTxt ==={vbCrLf}{vbCrLf}"
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{ex.Message}{vbCrLf}{vbCrLf}"
        End Try
    End Sub

    ''' <summary>
    ''' 追加Error info類型的文字
    ''' </summary>
    ''' <param name="title">錯誤標題</param>
    Public Sub writeTitleIntoError_InfoTxt(title As String)
        Try
            Using ws As StreamWriter = File.AppendText(errorTxtPath)
                ws.WriteLine($"{Now} ==== {title} ===={vbCrLf}")
            End Using
        Catch ex As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"=== errorInfo.writeTitleIntoError_InfoTxt ==={vbCrLf}{vbCrLf}"
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{ex.Message}{vbCrLf}{vbCrLf}"
        End Try

    End Sub

    ''' <summary>
    ''' 追加Error info文字
    ''' </summary>
    ''' <param name="msg">錯誤訊息</param>
    Public Sub writeInfoError_InfoTxt(msg As String)
        'Add first error msg to the txt file
        Try
            Using ws As StreamWriter = File.AppendText(errorTxtPath)
                ws.WriteLine($"{msg}{vbCrLf}")
            End Using
        Catch ex As Exception
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"=== errorInfo.writeInfoError_InfoTxt ==={vbCrLf}{vbCrLf}"
            JobMaker_Form.ResultFailOutput_TextBox.Text += $"{ex.Message}{vbCrLf}{vbCrLf}"
        End Try

    End Sub

    Public Sub writeInfoError_errorMsg(title_errorMsg As String, content_errrorMsg As String,
                                       ex As Exception)
        writeTitleIntoError_InfoTxt($"{title_errorMsg}")
        writeInfoError_InfoTxt($"{content_errrorMsg}{vbCrLf}{ex.Message}")
        JobMaker_Form.ResultFailOutput_TextBox.Text += $"{content_errrorMsg}{vbCrLf}{ex.Message}{vbCrLf}"
    End Sub

    ''' <summary>
    ''' 輸出文字至TextBox中，並將插入符號保持在最下方
    ''' </summary>
    ''' <param name="outputText">要輸出的文字</param>
    Public Sub writeInfo_toTextBox_focusOnBelow(tb As TextBox, outputText As String)
        With tb
            .Text += $"{outputText}{vbCrLf}"
            .SelectionStart = .TextLength
            .ScrollToCaret()
        End With
    End Sub
End Module
