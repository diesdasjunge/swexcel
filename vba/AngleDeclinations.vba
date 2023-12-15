Sub AngleDeclinations()
    Dim x As Double, y As Double, z As Double
    
    OB = -Radians(epsilon)
    
    x = COSINE(Natal(14))
    y = Cos(-OB) * Sine(Natal(14))
    z = Sin(-OB) * Sine(Natal(14))
    
    NatalDecl(14) = Cart2Sph(x, y, z)
    
    x = COSINE(Natal(15))
    y = Cos(-OB) * Sine(Natal(15))
    z = Sin(-OB) * Sine(Natal(15))
    
    NatalDecl(15) = Cart2Sph(x, y, z)
End Sub