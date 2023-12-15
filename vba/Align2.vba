Function Align2(A4 As String)
    If Val(A4) >= 10 Then
        Align2 = A4
    Else
        Align2 = Space(1) & A4
    End If
End Function