Function Cart2Sph(x As Double, y As Double, z As Double) As Double
    Dim BETA As Double, R As Double, RHO As Double, ZETA As Double
    
    R = Sqr(x * x + y * y + z * z)
    RHO = Sqr(x * x + y * y)
    ZETA = 2 * DEGREES(Atn(y / (Abs(x) + Sqr(x * x + y * y))))
    
    If RHO = 0 And z > 0 Then
        BETA = 90
    ElseIf RHO = 0 And z = 0 Then
        BETA = 0
    ElseIf RHO = 0 And z < 0 Then
        BETA = -90
    ElseIf RHO <> 0 Then
        BETA = DEGREES(Atn(z / RHO))
    End If

    Cart2Sph = BETA
End Function