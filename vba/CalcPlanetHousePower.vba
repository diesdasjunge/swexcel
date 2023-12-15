Sub CalcPlanetHousePower()
    Dim x, y

    For x = 1 To 10: y = Val(Mid$(NatalPH$, x * 10 - 9, 2))
    PlanetHousePower(x) = XHousePowerVariation!(y) * DistanceFromCusp(x) / DegreesInEachHouse(y)
    PlanetHousePower(x) = XHousePower!(y) - PlanetHousePower(x)
    Next x
End Sub