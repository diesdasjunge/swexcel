Function Align4(A4 As String)
    If Left$(A4, 3) = "-0," Then
        Align4 = Space(2) & A4: Exit Function
    ElseIf Val(A4) >= 100 Then
        Align4 = Space(1) & A4: Exit Function
    ElseIf Val(A4) >= 10 Then
        Align4 = Space(2) & A4: Exit Function
    ElseIf Val(A4) >= 0 Then
        Align4 = Space(3) & A4: Exit Function
    ElseIf Val(A4) > -10 Then
        Align4 = Space(2) & A4: Exit Function
    ElseIf Val(A4) > -100 Then
        Align4 = Space(1) & A4: Exit Function
    Else
        Align4 = A4
    End If
End Function