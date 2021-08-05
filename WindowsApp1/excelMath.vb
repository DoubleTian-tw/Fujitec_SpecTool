Public Class excelMath
    Public Function convertColumn_fromIntToString(getCol_fromInt As Integer) As String
        Dim modulo As Integer

        While getCol_fromInt > 0
            modulo = (getCol_fromInt - 1) Mod 26
            convertColumn_fromIntToString = Convert.ToChar(65 + modulo).ToString() + convertColumn_fromIntToString
            getCol_fromInt = CInt((getCol_fromInt - modulo) / 26)
        End While

        Return convertColumn_fromIntToString
    End Function
End Class
