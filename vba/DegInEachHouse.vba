Sub DegInEachHouse()
    For I = 1 To 11: DegreesInEachHouse(I) = HC3(I + 1) - HC3(I)
    If DegreesInEachHouse(I) < 0 Then DegreesInEachHouse(I) = DegreesInEachHouse(I) + 360
    Next I
    DegreesInEachHouse(12) = HC3(1) - HC3(12)
    If DegreesInEachHouse(12) < 0 Then DegreesInEachHouse(12) = DegreesInEachHouse(12) + 360
End Sub