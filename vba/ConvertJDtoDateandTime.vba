Sub ConvertJDtoDateandTime(XjdX As Double)
'Date is stored in JDDate$ and time is stored in JDTime$
    Dim fraction As Double, hh As Integer, mm As Integer, sd As Date
    Dim temp_JD As Double
    Dim t1$, y7 As Long, m7 As Long, d7 As Long, h7 As Double

    temp_JD = XjdX              'always GMT
    Call swe_revjul(temp_JD + 0.000001, 1, y7, m7, d7, h7)

    If MyiDate% = 0 Then
        JDDate$ = Format$(m7 & "/" & d7 & "/" & y7, "mm/dd/yyyy")
    ElseIf MyiDate% = 1 Then
        JDDate$ = Format$(m7 & "/" & d7 & "/" & y7, "dd/mm/yyyy")
    ElseIf MyiDate% = 2 Then
        JDDate$ = Format$(m7 & "/" & d7 & "/" & y7, "yyyy/mm/dd")
    End If
    
    JDTime$ = Format$(h7 / 24, "hh:nn:ss")

    If JDTime$ = "24:00:00" Then
        JDTime$ = "00:00:00"
        
        sd = JDDate$
        sd = sd + 1
        
        JDDate$ = sd
    End If
End Sub