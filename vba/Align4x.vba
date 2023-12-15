Function Align4x(A4 As String)
    If Left$(A4, 3) = "-0," Then
        Align4x = Space(1) & A4: Exit Function
    ElseIf Val(A4) >= 10 Then
        Align4x = Space(1) & A4: Exit Function
    ElseIf Val(A4) >= 0 Then
        Align4x = Space(2) & A4: Exit Function
    ElseIf Val(A4) > -10 Then
        Align4x = Space(1) & A4: Exit Function
    Else
        Align4x = A4
    End If
End Function