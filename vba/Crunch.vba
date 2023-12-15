Public Function Crunch(x As Double) As Double
    If x >= 0 Then
        Crunch = x - Int(x / 360) * 360
    Else
        Crunch = 360 + (x - ((1 + Int(x / 360)) * 360))
    End If
End Function