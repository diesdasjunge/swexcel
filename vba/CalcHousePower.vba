Sub CalcHousePower()
'find unoccupied house power for houses 1 - 11
    For x = 1 To 11: TEMP7 = Int(HC3(x) / 30) + 1
    TEMP8 = Int(HC3(x + 1) / 30) + 1
    If TEMP8 - TEMP7 < 0 Then TEMP8 = TEMP8 + 12
'do ruler of sign on house cusp
    If TEMP7 = 1 Then HousePower(x) = PlanetTotalPower(4) / 2
    If TEMP7 = 2 Then HousePower(x) = PlanetTotalPower(3) / 2
    If TEMP7 = 3 Then HousePower(x) = PlanetTotalPower(2) / 2
    If TEMP7 = 4 Then HousePower(x) = PlanetTotalPower(10) / 2
    If TEMP7 = 5 Then HousePower(x) = PlanetTotalPower(1) / 2
    If TEMP7 = 6 Then HousePower(x) = PlanetTotalPower(2) / 2
    If TEMP7 = 7 Then HousePower(x) = PlanetTotalPower(3) / 2
    If TEMP7 = 8 Then HousePower(x) = ((PlanetTotalPower(4) + PlanetTotalPower(9)) / 2) / 2
    If TEMP7 = 9 Then HousePower(x) = PlanetTotalPower(5) / 2
    If TEMP7 = 10 Then HousePower(x) = PlanetTotalPower(6) / 2
    If TEMP7 = 11 Then HousePower(x) = ((PlanetTotalPower(6) + PlanetTotalPower(7)) / 2) / 2
    If TEMP7 = 12 Then HousePower(x) = ((PlanetTotalPower(5) + PlanetTotalPower(8)) / 2) / 2
'do ruler of intercepted sign in the house
    If TEMP8 - TEMP7 >= 2 Then
CHP1:      TEMP7 = TEMP7 + 1: If TEMP7 = TEMP8 Then GoTo CHP2
       T7 = TEMP7: If T7 > 12 Then T7 = T7 - 12
       If T7 = 1 Then HousePower(x) = HousePower(x) + PlanetTotalPower(4) / 4
       If T7 = 2 Then HousePower(x) = HousePower(x) + PlanetTotalPower(3) / 4
       If T7 = 3 Then HousePower(x) = HousePower(x) + PlanetTotalPower(2) / 4
       If T7 = 4 Then HousePower(x) = HousePower(x) + PlanetTotalPower(10) / 4
       If T7 = 5 Then HousePower(x) = HousePower(x) + PlanetTotalPower(1) / 4
       If T7 = 6 Then HousePower(x) = HousePower(x) + PlanetTotalPower(2) / 4
       If T7 = 7 Then HousePower(x) = HousePower(x) + PlanetTotalPower(3) / 4
       If T7 = 8 Then HousePower(x) = HousePower(x) + ((PlanetTotalPower(4) + PlanetTotalPower(9)) / 2) / 4
       If T7 = 9 Then HousePower(x) = HousePower(x) + PlanetTotalPower(5) / 4
       If T7 = 10 Then HousePower(x) = HousePower(x) + PlanetTotalPower(6) / 4
       If T7 = 11 Then HousePower(x) = HousePower(x) + ((PlanetTotalPower(6) + PlanetTotalPower(7)) / 2) / 4
       If T7 = 12 Then HousePower(x) = HousePower(x) + ((PlanetTotalPower(5) + PlanetTotalPower(8)) / 2) / 4
       GoTo CHP1
    End If
CHP2:   Next x
'find unoccupied house power for house 12
    TEMP7 = Int(HC3(12) / 30) + 1: TEMP8 = Int(HC3(1) / 30) + 1
    If TEMP8 - TEMP7 < 0 Then TEMP8 = TEMP8 + 12
'do ruler of sign on house cusp
    If TEMP7 = 1 Then HousePower(12) = PlanetTotalPower(4) / 2
    If TEMP7 = 2 Then HousePower(12) = PlanetTotalPower(3) / 2
    If TEMP7 = 3 Then HousePower(12) = PlanetTotalPower(2) / 2
    If TEMP7 = 4 Then HousePower(12) = PlanetTotalPower(10) / 2
    If TEMP7 = 5 Then HousePower(12) = PlanetTotalPower(1) / 2
    If TEMP7 = 6 Then HousePower(12) = PlanetTotalPower(2) / 2
    If TEMP7 = 7 Then HousePower(12) = PlanetTotalPower(3) / 2
    If TEMP7 = 8 Then HousePower(12) = ((PlanetTotalPower(4) + PlanetTotalPower(9)) / 2) / 2
    If TEMP7 = 9 Then HousePower(12) = PlanetTotalPower(5) / 2
    If TEMP7 = 10 Then HousePower(12) = PlanetTotalPower(6) / 2
    If TEMP7 = 11 Then HousePower(12) = ((PlanetTotalPower(6) + PlanetTotalPower(7)) / 2) / 2
    If TEMP7 = 12 Then HousePower(12) = ((PlanetTotalPower(5) + PlanetTotalPower(8)) / 2) / 2
'do ruler of intercepted sign in the house
    If TEMP8 - TEMP7 >= 2 Then
CHP1X:     TEMP7 = TEMP7 + 1: If TEMP7 = TEMP8 Then GoTo CHP2X
       T7 = TEMP7: If T7 > 12 Then T7 = T7 - 12
       If T7 = 1 Then HousePower(12) = HousePower(12) + PlanetTotalPower(4) / 4
       If T7 = 2 Then HousePower(12) = HousePower(12) + PlanetTotalPower(3) / 4
       If T7 = 3 Then HousePower(12) = HousePower(12) + PlanetTotalPower(2) / 4
       If T7 = 4 Then HousePower(12) = HousePower(12) + PlanetTotalPower(10) / 4
       If T7 = 5 Then HousePower(12) = HousePower(12) + PlanetTotalPower(1) / 4
       If T7 = 6 Then HousePower(12) = HousePower(12) + PlanetTotalPower(2) / 4
       If T7 = 7 Then HousePower(12) = HousePower(12) + PlanetTotalPower(3) / 4
       If T7 = 8 Then HousePower(12) = HousePower(12) + ((PlanetTotalPower(4) + PlanetTotalPower(9)) / 2) / 4
       If T7 = 9 Then HousePower(12) = HousePower(12) + PlanetTotalPower(5) / 4
       If T7 = 10 Then HousePower(12) = HousePower(12) + PlanetTotalPower(6) / 4
       If T7 = 11 Then HousePower(12) = HousePower(12) + ((PlanetTotalPower(6) + PlanetTotalPower(7)) / 2) / 4
       If T7 = 12 Then HousePower(12) = HousePower(12) + ((PlanetTotalPower(5) + PlanetTotalPower(8)) / 2) / 4
       GoTo CHP1X
    End If
CHP2X:
'find unoccupied house harmony for houses 1 - 11
    For x = 1 To 11: TEMP7 = Int(HC3(x) / 30) + 1
    TEMP8 = Int(HC3(x + 1) / 30) + 1
    If TEMP8 - TEMP7 < 0 Then TEMP8 = TEMP8 + 12
'do ruler of sign on house cusp
    If TEMP7 = 1 Then HouseHarmony(x) = PlanetHarmony(4) / 2
    If TEMP7 = 2 Then HouseHarmony(x) = PlanetHarmony(3) / 2
    If TEMP7 = 3 Then HouseHarmony(x) = PlanetHarmony(2) / 2
    If TEMP7 = 4 Then HouseHarmony(x) = PlanetHarmony(10) / 2
    If TEMP7 = 5 Then HouseHarmony(x) = PlanetHarmony(1) / 2
    If TEMP7 = 6 Then HouseHarmony(x) = PlanetHarmony(2) / 2
    If TEMP7 = 7 Then HouseHarmony(x) = PlanetHarmony(3) / 2
    If TEMP7 = 8 Then HouseHarmony(x) = ((PlanetHarmony(4) + PlanetHarmony(9)) / 2) / 2
    If TEMP7 = 9 Then HouseHarmony(x) = PlanetHarmony(5) / 2
    If TEMP7 = 10 Then HouseHarmony(x) = PlanetHarmony(6) / 2
    If TEMP7 = 11 Then HouseHarmony(x) = ((PlanetHarmony(6) + PlanetHarmony(7)) / 2) / 2
    If TEMP7 = 12 Then HouseHarmony(x) = ((PlanetHarmony(5) + PlanetHarmony(8)) / 2) / 2
'do ruler of intercepted sign in the house
    If TEMP8 - TEMP7 >= 2 Then
CHP3:      TEMP7 = TEMP7 + 1: If TEMP7 = TEMP8 Then GoTo CHP4
       T7 = TEMP7: If T7 > 12 Then T7 = T7 - 12
       If T7 = 1 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(4) / 4
       If T7 = 2 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(3) / 4
       If T7 = 3 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(2) / 4
       If T7 = 4 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(10) / 4
       If T7 = 5 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(1) / 4
       If T7 = 6 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(2) / 4
       If T7 = 7 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(3) / 4
       If T7 = 8 Then HouseHarmony(x) = HouseHarmony(x) + ((PlanetHarmony(4) + PlanetHarmony(9)) / 2) / 4
       If T7 = 9 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(5) / 4
       If T7 = 10 Then HouseHarmony(x) = HouseHarmony(x) + PlanetHarmony(6) / 4
       If T7 = 11 Then HouseHarmony(x) = HouseHarmony(x) + ((PlanetHarmony(6) + PlanetHarmony(7)) / 2) / 4
       If T7 = 12 Then HouseHarmony(x) = HouseHarmony(x) + ((PlanetHarmony(5) + PlanetHarmony(8)) / 2) / 4
       GoTo CHP3
    End If
CHP4:   Next x
'find unoccupied house harmony for house 12
    TEMP7 = Int(HC3(12) / 30) + 1: TEMP8 = Int(HC3(1) / 30) + 1
    If TEMP8 - TEMP7 < 0 Then TEMP8 = TEMP8 + 12
'do ruler of sign on house cusp
    If TEMP7 = 1 Then HouseHarmony(12) = PlanetHarmony(4) / 2
    If TEMP7 = 2 Then HouseHarmony(12) = PlanetHarmony(3) / 2
    If TEMP7 = 3 Then HouseHarmony(12) = PlanetHarmony(2) / 2
    If TEMP7 = 4 Then HouseHarmony(12) = PlanetHarmony(10) / 2
    If TEMP7 = 5 Then HouseHarmony(12) = PlanetHarmony(1) / 2
    If TEMP7 = 6 Then HouseHarmony(12) = PlanetHarmony(2) / 2
    If TEMP7 = 7 Then HouseHarmony(12) = PlanetHarmony(3) / 2
    If TEMP7 = 8 Then HouseHarmony(12) = ((PlanetHarmony(4) + PlanetHarmony(9)) / 2) / 2
    If TEMP7 = 9 Then HouseHarmony(12) = PlanetHarmony(5) / 2
    If TEMP7 = 10 Then HouseHarmony(12) = PlanetHarmony(6) / 2
    If TEMP7 = 11 Then HouseHarmony(12) = ((PlanetHarmony(6) + PlanetHarmony(7)) / 2) / 2
    If TEMP7 = 12 Then HouseHarmony(12) = ((PlanetHarmony(5) + PlanetHarmony(8)) / 2) / 2
'do ruler of intercepted sign in the house
    If TEMP8 - TEMP7 >= 2 Then
CHP3X:     TEMP7 = TEMP7 + 1: If TEMP7 = TEMP8 Then GoTo CHP4X
       T7 = TEMP7: If T7 > 12 Then T7 = T7 - 12
       If T7 = 1 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(4) / 4
       If T7 = 2 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(3) / 4
       If T7 = 3 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(2) / 4
       If T7 = 4 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(10) / 4
       If T7 = 5 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(1) / 4
       If T7 = 6 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(2) / 4
       If T7 = 7 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(3) / 4
       If T7 = 8 Then HouseHarmony(12) = HouseHarmony(12) + ((PlanetHarmony(4) + PlanetHarmony(9)) / 2) / 4
       If T7 = 9 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(5) / 4
       If T7 = 10 Then HouseHarmony(12) = HouseHarmony(12) + PlanetHarmony(6) / 4
       If T7 = 11 Then HouseHarmony(12) = HouseHarmony(12) + ((PlanetHarmony(6) + PlanetHarmony(7)) / 2) / 4
       If T7 = 12 Then HouseHarmony(12) = HouseHarmony(12) + ((PlanetHarmony(5) + PlanetHarmony(8)) / 2) / 4
       GoTo CHP3X
    End If
CHP4X:
'find unoccupied house discord for houses 1 - 11
    For x = 1 To 11: TEMP7 = Int(HC3(x) / 30) + 1
    TEMP8 = Int(HC3(x + 1) / 30) + 1
    If TEMP8 - TEMP7 < 0 Then TEMP8 = TEMP8 + 12
'do ruler of sign on house cusp
    If TEMP7 = 1 Then HouseDiscord(x) = PlanetDiscord(4) / 2
    If TEMP7 = 2 Then HouseDiscord(x) = PlanetDiscord(3) / 2
    If TEMP7 = 3 Then HouseDiscord(x) = PlanetDiscord(2) / 2
    If TEMP7 = 4 Then HouseDiscord(x) = PlanetDiscord(10) / 2
    If TEMP7 = 5 Then HouseDiscord(x) = PlanetDiscord(1) / 2
    If TEMP7 = 6 Then HouseDiscord(x) = PlanetDiscord(2) / 2
    If TEMP7 = 7 Then HouseDiscord(x) = PlanetDiscord(3) / 2
    If TEMP7 = 8 Then HouseDiscord(x) = ((PlanetDiscord(4) + PlanetDiscord(9)) / 2) / 2
    If TEMP7 = 9 Then HouseDiscord(x) = PlanetDiscord(5) / 2
    If TEMP7 = 10 Then HouseDiscord(x) = PlanetDiscord(6) / 2
    If TEMP7 = 11 Then HouseDiscord(x) = ((PlanetDiscord(6) + PlanetDiscord(7)) / 2) / 2
    If TEMP7 = 12 Then HouseDiscord(x) = ((PlanetDiscord(5) + PlanetDiscord(8)) / 2) / 2
'do ruler of intercepted sign in the house
    If TEMP8 - TEMP7 >= 2 Then
CHP5:      TEMP7 = TEMP7 + 1: If TEMP7 = TEMP8 Then GoTo CHP6
       T7 = TEMP7: If T7 > 12 Then T7 = T7 - 12
       If T7 = 1 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(4) / 4
       If T7 = 2 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(3) / 4
       If T7 = 3 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(2) / 4
       If T7 = 4 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(10) / 4
       If T7 = 5 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(1) / 4
       If T7 = 6 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(2) / 4
       If T7 = 7 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(3) / 4
       If T7 = 8 Then HouseDiscord(x) = HouseDiscord(x) + ((PlanetDiscord(4) + PlanetDiscord(9)) / 2) / 4
       If T7 = 9 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(5) / 4
       If T7 = 10 Then HouseDiscord(x) = HouseDiscord(x) + PlanetDiscord(6) / 4
       If T7 = 11 Then HouseDiscord(x) = HouseDiscord(x) + ((PlanetDiscord(6) + PlanetDiscord(7)) / 2) / 4
       If T7 = 12 Then HouseDiscord(x) = HouseDiscord(x) + ((PlanetDiscord(5) + PlanetDiscord(8)) / 2) / 4
       GoTo CHP5
    End If
CHP6:   Next x
'find unoccupied house discord for house 12
    TEMP7 = Int(HC3(12) / 30) + 1: TEMP8 = Int(HC3(1) / 30) + 1
    If TEMP8 - TEMP7 < 0 Then TEMP8 = TEMP8 + 12
'do ruler of sign on house cusp
    If TEMP7 = 1 Then HouseDiscord(12) = PlanetDiscord(4) / 2
    If TEMP7 = 2 Then HouseDiscord(12) = PlanetDiscord(3) / 2
    If TEMP7 = 3 Then HouseDiscord(12) = PlanetDiscord(2) / 2
    If TEMP7 = 4 Then HouseDiscord(12) = PlanetDiscord(10) / 2
    If TEMP7 = 5 Then HouseDiscord(12) = PlanetDiscord(1) / 2
    If TEMP7 = 6 Then HouseDiscord(12) = PlanetDiscord(2) / 2
    If TEMP7 = 7 Then HouseDiscord(12) = PlanetDiscord(3) / 2
    If TEMP7 = 8 Then HouseDiscord(12) = ((PlanetDiscord(4) + PlanetDiscord(9)) / 2) / 2
    If TEMP7 = 9 Then HouseDiscord(12) = PlanetDiscord(5) / 2
    If TEMP7 = 10 Then HouseDiscord(12) = PlanetDiscord(6) / 2
    If TEMP7 = 11 Then HouseDiscord(12) = ((PlanetDiscord(6) + PlanetDiscord(7)) / 2) / 2
    If TEMP7 = 12 Then HouseDiscord(12) = ((PlanetDiscord(5) + PlanetDiscord(8)) / 2) / 2
'do ruler of intercepted sign in the house
    If TEMP8 - TEMP7 >= 2 Then
CHP5X:     TEMP7 = TEMP7 + 1: If TEMP7 = TEMP8 Then GoTo CHP6X
       T7 = TEMP7: If T7 > 12 Then T7 = T7 - 12
       If T7 = 1 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(4) / 4
       If T7 = 2 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(3) / 4
       If T7 = 3 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(2) / 4
       If T7 = 4 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(10) / 4
       If T7 = 5 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(1) / 4
       If T7 = 6 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(2) / 4
       If T7 = 7 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(3) / 4
       If T7 = 8 Then HouseDiscord(12) = HouseDiscord(12) + ((PlanetDiscord(4) + PlanetDiscord(9)) / 2) / 4
       If T7 = 9 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(5) / 4
       If T7 = 10 Then HouseDiscord(12) = HouseDiscord(12) + PlanetDiscord(6) / 4
       If T7 = 11 Then HouseDiscord(12) = HouseDiscord(12) + ((PlanetDiscord(6) + PlanetDiscord(7)) / 2) / 4
       If T7 = 12 Then HouseDiscord(12) = HouseDiscord(12) + ((PlanetDiscord(5) + PlanetDiscord(8)) / 2) / 4
       GoTo CHP5X
    End If
CHP6X:  For x = 1 To 10: TEMP7 = Val(Mid$(NatalPH$, x * 10 - 9, 2))
    HousePower(TEMP7) = HousePower(TEMP7) + PlanetTotalPower(x)
    HouseHarmony(TEMP7) = HouseHarmony(TEMP7) + PlanetHarmony(x)
    HouseDiscord(TEMP7) = HouseDiscord(TEMP7) + PlanetDiscord(x): Next x
    HousePower(1) = HousePower(1) + PlanetTotalPower(14)
    HouseHarmony(1) = HouseHarmony(1) + PlanetHarmony(14)
    HouseDiscord(1) = HouseDiscord(1) + PlanetDiscord(14)
    HousePower(10) = HousePower(10) + PlanetTotalPower(15)
    HouseHarmony(10) = HouseHarmony(10) + PlanetHarmony(15)
    HouseDiscord(10) = HouseDiscord(10) + PlanetDiscord(15)
End Sub