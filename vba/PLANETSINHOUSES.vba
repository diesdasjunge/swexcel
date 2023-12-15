Sub PLANETSINHOUSES()
    Dim CorrPlanetsY As Double, x As Integer, y As Integer

    For x = 1 To 12
        For y = 1 To 15
            Mid$(PH$, y * 10 - 5, 10) = " HOUSE"
        
            CorrPlanetsY = Planets(y) + (1 / 36000)
            
            If x < 12 And HC(x) > HC(x + 1) Then GoTo PIH6
        
            If x = 12 And HC(x) > HC(1) Then GoTo PIH8
        
            If CorrPlanetsY >= HC(x) And CorrPlanetsY < HC(x + 1) And x < 12 Then Mid$(PH$, y * 10 - 9, 4) = Mid$(HCX1$, x * 4 - 3, 4): GoTo PIH9
        
            If CorrPlanetsY >= HC(x) And CorrPlanetsY < HC(1) And x = 12 Then Mid$(PH$, y * 10 - 9, 4) = Mid$(HCX1$, x * 4 - 3, 4)
        
            GoTo PIH9

PIH6:
            If (CorrPlanetsY >= HC(x) And CorrPlanetsY < 360) Or (CorrPlanetsY < HC(x + 1) And CorrPlanetsY >= 0) Then Mid$(PH$, y * 10 - 9, y * 10 - 6) = Mid$(HCX1$, x * 4 - 3, 4)
            GoTo PIH9

PIH8:
            If (CorrPlanetsY >= HC(x) And CorrPlanetsY < 360) Or (CorrPlanetsY < HC(1) And CorrPlanetsY >= 0) Then Mid$(PH$, y * 10 - 9, y * 10 - 6) = Mid$(HCX1$, x * 4 - 3, 4)
PIH9:
        Next y
    Next x
End Sub