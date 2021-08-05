Imports Microsoft.VisualBasic

Public Class JB_FormSize
    Enum mySize
        Ini_Size
        Re_Size
    End Enum

    Public Sub get_Size(ByVal FS As mySize)
        Select Case mySize
            Case mySize.Ini_Size
                Return
            Case mySize.Re_Size

        End Select
    End Sub
End Class
