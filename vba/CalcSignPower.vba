Sub CalcSignPower()
'find unoccupied sign power
    If InterceptedSign%(1) = 0 Then SignPower(1) = PlanetTotalPower(4) / 4
    If InterceptedSign%(1) = 1 Then SignPower(1) = PlanetTotalPower(4) / 2
    If InterceptedSign%(1) = 2 Then SignPower(1) = PlanetTotalPower(4)
    If InterceptedSign%(1) = 3 Then SignPower(1) = PlanetTotalPower(4) * 2
    If InterceptedSign%(2) = 0 Then SignPower(2) = PlanetTotalPower(3) / 4
    If InterceptedSign%(2) = 1 Then SignPower(2) = PlanetTotalPower(3) / 2
    If InterceptedSign%(2) = 2 Then SignPower(2) = PlanetTotalPower(3)
    If InterceptedSign%(2) = 3 Then SignPower(2) = PlanetTotalPower(3) * 2
    If InterceptedSign%(3) = 0 Then SignPower(3) = PlanetTotalPower(2) / 4
    If InterceptedSign%(3) = 1 Then SignPower(3) = PlanetTotalPower(2) / 2
    If InterceptedSign%(3) = 2 Then SignPower(3) = PlanetTotalPower(2)
    If InterceptedSign%(3) = 3 Then SignPower(3) = PlanetTotalPower(2) * 2
    If InterceptedSign%(4) = 0 Then SignPower(4) = PlanetTotalPower(10) / 4
    If InterceptedSign%(4) = 1 Then SignPower(4) = PlanetTotalPower(10) / 2
    If InterceptedSign%(4) = 2 Then SignPower(4) = PlanetTotalPower(10)
    If InterceptedSign%(4) = 3 Then SignPower(4) = PlanetTotalPower(10) * 2
    If InterceptedSign%(5) = 0 Then SignPower(5) = PlanetTotalPower(1) / 4
    If InterceptedSign%(5) = 1 Then SignPower(5) = PlanetTotalPower(1) / 2
    If InterceptedSign%(5) = 2 Then SignPower(5) = PlanetTotalPower(1)
    If InterceptedSign%(5) = 3 Then SignPower(5) = PlanetTotalPower(1) * 2
    If InterceptedSign%(6) = 0 Then SignPower(6) = PlanetTotalPower(2) / 4
    If InterceptedSign%(6) = 1 Then SignPower(6) = PlanetTotalPower(2) / 2
    If InterceptedSign%(6) = 2 Then SignPower(6) = PlanetTotalPower(2)
    If InterceptedSign%(6) = 3 Then SignPower(6) = PlanetTotalPower(2) * 2
    If InterceptedSign%(7) = 0 Then SignPower(7) = PlanetTotalPower(3) / 4
    If InterceptedSign%(7) = 1 Then SignPower(7) = PlanetTotalPower(3) / 2
    If InterceptedSign%(7) = 2 Then SignPower(7) = PlanetTotalPower(3)
    If InterceptedSign%(7) = 3 Then SignPower(7) = PlanetTotalPower(3) * 2
    If InterceptedSign%(8) = 0 Then SignPower(8) = ((PlanetTotalPower(4) + PlanetTotalPower(9)) / 2) / 4
    If InterceptedSign%(8) = 1 Then SignPower(8) = ((PlanetTotalPower(4) + PlanetTotalPower(9)) / 2) / 2
    If InterceptedSign%(8) = 2 Then SignPower(8) = ((PlanetTotalPower(4) + PlanetTotalPower(9)) / 2)
    If InterceptedSign%(8) = 3 Then SignPower(8) = ((PlanetTotalPower(4) + PlanetTotalPower(9)) / 2) * 2
    If InterceptedSign%(9) = 0 Then SignPower(9) = PlanetTotalPower(5) / 4
    If InterceptedSign%(9) = 1 Then SignPower(9) = PlanetTotalPower(5) / 2
    If InterceptedSign%(9) = 2 Then SignPower(9) = PlanetTotalPower(5)
    If InterceptedSign%(9) = 3 Then SignPower(9) = PlanetTotalPower(5) * 2
    If InterceptedSign%(10) = 0 Then SignPower(10) = PlanetTotalPower(6) / 4
    If InterceptedSign%(10) = 1 Then SignPower(10) = PlanetTotalPower(6) / 2
    If InterceptedSign%(10) = 2 Then SignPower(10) = PlanetTotalPower(6)
    If InterceptedSign%(10) = 3 Then SignPower(10) = PlanetTotalPower(6) * 2
    If InterceptedSign%(11) = 0 Then SignPower(11) = ((PlanetTotalPower(6) + PlanetTotalPower(7)) / 2) / 4
    If InterceptedSign%(11) = 1 Then SignPower(11) = ((PlanetTotalPower(6) + PlanetTotalPower(7)) / 2) / 2
    If InterceptedSign%(11) = 2 Then SignPower(11) = ((PlanetTotalPower(6) + PlanetTotalPower(7)) / 2)
    If InterceptedSign%(11) = 3 Then SignPower(11) = ((PlanetTotalPower(6) + PlanetTotalPower(7)) / 2) * 2
    If InterceptedSign%(12) = 0 Then SignPower(12) = ((PlanetTotalPower(5) + PlanetTotalPower(8)) / 2) / 4
    If InterceptedSign%(12) = 1 Then SignPower(12) = ((PlanetTotalPower(5) + PlanetTotalPower(8)) / 2) / 2
    If InterceptedSign%(12) = 2 Then SignPower(12) = ((PlanetTotalPower(5) + PlanetTotalPower(8)) / 2)
    If InterceptedSign%(12) = 3 Then SignPower(12) = ((PlanetTotalPower(5) + PlanetTotalPower(8)) / 2) * 2
'find unoccupied sign harmony
    If InterceptedSign%(1) = 0 Then SignHarmony(1) = PlanetHarmony(4) / 4
    If InterceptedSign%(1) = 1 Then SignHarmony(1) = PlanetHarmony(4) / 2
    If InterceptedSign%(1) = 2 Then SignHarmony(1) = PlanetHarmony(4)
    If InterceptedSign%(1) = 3 Then SignHarmony(1) = PlanetHarmony(4) * 2
    If InterceptedSign%(2) = 0 Then SignHarmony(2) = PlanetHarmony(3) / 4
    If InterceptedSign%(2) = 1 Then SignHarmony(2) = PlanetHarmony(3) / 2
    If InterceptedSign%(2) = 2 Then SignHarmony(2) = PlanetHarmony(3)
    If InterceptedSign%(2) = 3 Then SignHarmony(2) = PlanetHarmony(3) * 2
    If InterceptedSign%(3) = 0 Then SignHarmony(3) = PlanetHarmony(2) / 4
    If InterceptedSign%(3) = 1 Then SignHarmony(3) = PlanetHarmony(2) / 2
    If InterceptedSign%(3) = 2 Then SignHarmony(3) = PlanetHarmony(2)
    If InterceptedSign%(3) = 3 Then SignHarmony(3) = PlanetHarmony(2) * 2
    If InterceptedSign%(4) = 0 Then SignHarmony(4) = PlanetHarmony(10) / 4
    If InterceptedSign%(4) = 1 Then SignHarmony(4) = PlanetHarmony(10) / 2
    If InterceptedSign%(4) = 2 Then SignHarmony(4) = PlanetHarmony(10)
    If InterceptedSign%(4) = 3 Then SignHarmony(4) = PlanetHarmony(10) * 2
    If InterceptedSign%(5) = 0 Then SignHarmony(5) = PlanetHarmony(1) / 4
    If InterceptedSign%(5) = 1 Then SignHarmony(5) = PlanetHarmony(1) / 2
    If InterceptedSign%(5) = 2 Then SignHarmony(5) = PlanetHarmony(1)
    If InterceptedSign%(5) = 3 Then SignHarmony(5) = PlanetHarmony(1) * 2
    If InterceptedSign%(6) = 0 Then SignHarmony(6) = PlanetHarmony(2) / 4
    If InterceptedSign%(6) = 1 Then SignHarmony(6) = PlanetHarmony(2) / 2
    If InterceptedSign%(6) = 2 Then SignHarmony(6) = PlanetHarmony(2)
    If InterceptedSign%(6) = 3 Then SignHarmony(6) = PlanetHarmony(2) * 2
    If InterceptedSign%(7) = 0 Then SignHarmony(7) = PlanetHarmony(3) / 4
    If InterceptedSign%(7) = 1 Then SignHarmony(7) = PlanetHarmony(3) / 2
    If InterceptedSign%(7) = 2 Then SignHarmony(7) = PlanetHarmony(3)
    If InterceptedSign%(7) = 3 Then SignHarmony(7) = PlanetHarmony(3) * 2
    If InterceptedSign%(8) = 0 Then SignHarmony(8) = ((PlanetHarmony(4) + PlanetHarmony(9)) / 2) / 4
    If InterceptedSign%(8) = 1 Then SignHarmony(8) = ((PlanetHarmony(4) + PlanetHarmony(9)) / 2) / 2
    If InterceptedSign%(8) = 2 Then SignHarmony(8) = ((PlanetHarmony(4) + PlanetHarmony(9)) / 2)
    If InterceptedSign%(8) = 3 Then SignHarmony(8) = ((PlanetHarmony(4) + PlanetHarmony(9)) / 2) * 2
    If InterceptedSign%(9) = 0 Then SignHarmony(9) = PlanetHarmony(5) / 4
    If InterceptedSign%(9) = 1 Then SignHarmony(9) = PlanetHarmony(5) / 2
    If InterceptedSign%(9) = 2 Then SignHarmony(9) = PlanetHarmony(5)
    If InterceptedSign%(9) = 3 Then SignHarmony(9) = PlanetHarmony(5) * 2
    If InterceptedSign%(10) = 0 Then SignHarmony(10) = PlanetHarmony(6) / 4
    If InterceptedSign%(10) = 1 Then SignHarmony(10) = PlanetHarmony(6) / 2
    If InterceptedSign%(10) = 2 Then SignHarmony(10) = PlanetHarmony(6)
    If InterceptedSign%(10) = 3 Then SignHarmony(10) = PlanetHarmony(6) * 2
    If InterceptedSign%(11) = 0 Then SignHarmony(11) = ((PlanetHarmony(6) + PlanetHarmony(7)) / 2) / 4
    If InterceptedSign%(11) = 1 Then SignHarmony(11) = ((PlanetHarmony(6) + PlanetHarmony(7)) / 2) / 2
    If InterceptedSign%(11) = 2 Then SignHarmony(11) = ((PlanetHarmony(6) + PlanetHarmony(7)) / 2)
    If InterceptedSign%(11) = 3 Then SignHarmony(11) = ((PlanetHarmony(6) + PlanetHarmony(7)) / 2) * 2
    If InterceptedSign%(12) = 0 Then SignHarmony(12) = ((PlanetHarmony(5) + PlanetHarmony(8)) / 2) / 4
    If InterceptedSign%(12) = 1 Then SignHarmony(12) = ((PlanetHarmony(5) + PlanetHarmony(8)) / 2) / 2
    If InterceptedSign%(12) = 2 Then SignHarmony(12) = ((PlanetHarmony(5) + PlanetHarmony(8)) / 2)
    If InterceptedSign%(12) = 3 Then SignHarmony(12) = ((PlanetHarmony(5) + PlanetHarmony(8)) / 2) * 2
'find unoccupied sign discord
    If InterceptedSign%(1) = 0 Then SignDiscord(1) = PlanetDiscord(4) / 4
    If InterceptedSign%(1) = 1 Then SignDiscord(1) = PlanetDiscord(4) / 2
    If InterceptedSign%(1) = 2 Then SignDiscord(1) = PlanetDiscord(4)
    If InterceptedSign%(1) = 3 Then SignDiscord(1) = PlanetDiscord(4) * 2
    If InterceptedSign%(2) = 0 Then SignDiscord(2) = PlanetDiscord(3) / 4
    If InterceptedSign%(2) = 1 Then SignDiscord(2) = PlanetDiscord(3) / 2
    If InterceptedSign%(2) = 2 Then SignDiscord(2) = PlanetDiscord(3)
    If InterceptedSign%(2) = 3 Then SignDiscord(2) = PlanetDiscord(3) * 2
    If InterceptedSign%(3) = 0 Then SignDiscord(3) = PlanetDiscord(2) / 4
    If InterceptedSign%(3) = 1 Then SignDiscord(3) = PlanetDiscord(2) / 2
    If InterceptedSign%(3) = 2 Then SignDiscord(3) = PlanetDiscord(2)
    If InterceptedSign%(3) = 3 Then SignDiscord(3) = PlanetDiscord(2) * 2
    If InterceptedSign%(4) = 0 Then SignDiscord(4) = PlanetDiscord(10) / 4
    If InterceptedSign%(4) = 1 Then SignDiscord(4) = PlanetDiscord(10) / 2
    If InterceptedSign%(4) = 2 Then SignDiscord(4) = PlanetDiscord(10)
    If InterceptedSign%(4) = 3 Then SignDiscord(4) = PlanetDiscord(10) * 2
    If InterceptedSign%(5) = 0 Then SignDiscord(5) = PlanetDiscord(1) / 4
    If InterceptedSign%(5) = 1 Then SignDiscord(5) = PlanetDiscord(1) / 2
    If InterceptedSign%(5) = 2 Then SignDiscord(5) = PlanetDiscord(1)
    If InterceptedSign%(5) = 3 Then SignDiscord(5) = PlanetDiscord(1) * 2
    If InterceptedSign%(6) = 0 Then SignDiscord(6) = PlanetDiscord(2) / 4
    If InterceptedSign%(6) = 1 Then SignDiscord(6) = PlanetDiscord(2) / 2
    If InterceptedSign%(6) = 2 Then SignDiscord(6) = PlanetDiscord(2)
    If InterceptedSign%(6) = 3 Then SignDiscord(6) = PlanetDiscord(2) * 2
    If InterceptedSign%(7) = 0 Then SignDiscord(7) = PlanetDiscord(3) / 4
    If InterceptedSign%(7) = 1 Then SignDiscord(7) = PlanetDiscord(3) / 2
    If InterceptedSign%(7) = 2 Then SignDiscord(7) = PlanetDiscord(3)
    If InterceptedSign%(7) = 3 Then SignDiscord(7) = PlanetDiscord(3) * 2
    If InterceptedSign%(8) = 0 Then SignDiscord(8) = ((PlanetDiscord(4) + PlanetDiscord(9)) / 2) / 4
    If InterceptedSign%(8) = 1 Then SignDiscord(8) = ((PlanetDiscord(4) + PlanetDiscord(9)) / 2) / 2
    If InterceptedSign%(8) = 2 Then SignDiscord(8) = ((PlanetDiscord(4) + PlanetDiscord(9)) / 2)
    If InterceptedSign%(8) = 3 Then SignDiscord(8) = ((PlanetDiscord(4) + PlanetDiscord(9)) / 2) * 2
    If InterceptedSign%(9) = 0 Then SignDiscord(9) = PlanetDiscord(5) / 4
    If InterceptedSign%(9) = 1 Then SignDiscord(9) = PlanetDiscord(5) / 2
    If InterceptedSign%(9) = 2 Then SignDiscord(9) = PlanetDiscord(5)
    If InterceptedSign%(9) = 3 Then SignDiscord(9) = PlanetDiscord(5) * 2
    If InterceptedSign%(10) = 0 Then SignDiscord(10) = PlanetDiscord(6) / 4
    If InterceptedSign%(10) = 1 Then SignDiscord(10) = PlanetDiscord(6) / 2
    If InterceptedSign%(10) = 2 Then SignDiscord(10) = PlanetDiscord(6)
    If InterceptedSign%(10) = 3 Then SignDiscord(10) = PlanetDiscord(6) * 2
    If InterceptedSign%(11) = 0 Then SignDiscord(11) = ((PlanetDiscord(6) + PlanetDiscord(7)) / 2) / 4
    If InterceptedSign%(11) = 1 Then SignDiscord(11) = ((PlanetDiscord(6) + PlanetDiscord(7)) / 2) / 2
    If InterceptedSign%(11) = 2 Then SignDiscord(11) = ((PlanetDiscord(6) + PlanetDiscord(7)) / 2)
    If InterceptedSign%(11) = 3 Then SignDiscord(11) = ((PlanetDiscord(6) + PlanetDiscord(7)) / 2) * 2
    If InterceptedSign%(12) = 0 Then SignDiscord(12) = ((PlanetDiscord(5) + PlanetDiscord(8)) / 2) / 4
    If InterceptedSign%(12) = 1 Then SignDiscord(12) = ((PlanetDiscord(5) + PlanetDiscord(8)) / 2) / 2
    If InterceptedSign%(12) = 2 Then SignDiscord(12) = ((PlanetDiscord(5) + PlanetDiscord(8)) / 2)
    If InterceptedSign%(12) = 3 Then SignDiscord(12) = ((PlanetDiscord(5) + PlanetDiscord(8)) / 2) * 2
    For x = 1 To 15: If x = 11 Or x = 12 Or x = 13 Then GoTo CSP1
    TEMP7 = Int(Natal(x) / 30) + 1: SignPower(TEMP7) = SignPower(TEMP7) + PlanetTotalPower(x)
    SignHarmony(TEMP7) = SignHarmony(TEMP7) + PlanetHarmony(x)
    SignDiscord(TEMP7) = SignDiscord(TEMP7) + PlanetDiscord(x)
CSP1:   Next x
End Sub