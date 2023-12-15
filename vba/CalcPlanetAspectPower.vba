Sub CalcPlanetAspectPower()
    For y = 1 To 15: PlanetAspectPower(y) = 0: PlanetHarmony(y) = 0
    PlanetDiscord(y) = 0: Next y
    PlanetAspectPower(14) = 15: PlanetAspectPower(15) = 15
    For y = 1 To 15: If y = 11 Or y = 12 Or y = 13 Then GoTo PTT3
    For x = 1 To 15: If y = x Then GoTo PTT2
    If x = 11 Or x = 12 Or x = 13 Then GoTo PTT2
    YH1 = Val(Mid$(NatalPH$, y * 10 - 9, 2)): XH1 = Val(Mid$(NatalPH$, x * 10 - 9, 2))
    If y >= 14 Then YH1 = 10
    If x >= 14 Then XH1 = 10
'find ORB for planet Y, dependent upon what house Y is in AND whether it is a Luminary or a planet
    If y >= 14 Then
       ORB30Y = 3: ORB45Y = 5: ORB60Y = 7: ORB90Y = 10: ORB180Y = 12
       POW30Y = 3: POW45Y = 5: POW60Y = 7: POW90Y = 10: POW180Y = 12: GoTo CAP2
    End If
'using ONLY Sun and Moon as luminaries
       If (y <> 1 And y <> 10) And (YH1 = 3 Or YH1 = 6 Or YH1 = 9 Or YH1 = 12) Then
          ORB30Y = 1: ORB45Y = 3: ORB60Y = 5: ORB90Y = 6: ORB180Y = 8
          POW30Y = 1: POW45Y = 3: POW60Y = 5: POW90Y = 6: POW180Y = 8
          If y = 2 Then POW30Y = POW30Y + 1: POW45Y = POW45Y + 1: POW60Y = POW60Y + 1: POW90Y = POW90Y + 2: POW180Y = POW180Y + 3
       End If
       If (y = 1 Or y = 10) And (YH1 = 3 Or YH1 = 6 Or YH1 = 9 Or YH1 = 12) Then
          ORB30Y = 2: ORB45Y = 4: ORB60Y = 6: ORB90Y = 8: ORB180Y = 11
          POW30Y = 2: POW45Y = 4: POW60Y = 6: POW90Y = 8: POW180Y = 11
       End If
       If (y <> 1 And y <> 10) And (YH1 = 2 Or YH1 = 5 Or YH1 = 8 Or YH1 = 11) Then
          ORB30Y = 2: ORB45Y = 4: ORB60Y = 6: ORB90Y = 8: ORB180Y = 10
          POW30Y = 2: POW45Y = 4: POW60Y = 6: POW90Y = 8: POW180Y = 10
          If y = 2 Then POW30Y = POW30Y + 1: POW45Y = POW45Y + 1: POW60Y = POW60Y + 1: POW90Y = POW90Y + 2: POW180Y = POW180Y + 3
       End If
       If (y = 1 Or y = 10) And (YH1 = 2 Or YH1 = 5 Or YH1 = 8 Or YH1 = 11) Then
          ORB30Y = 3: ORB45Y = 5: ORB60Y = 7: ORB90Y = 10: ORB180Y = 13
          POW30Y = 3: POW45Y = 5: POW60Y = 7: POW90Y = 10: POW180Y = 13
       End If
       If (y <> 1 And y <> 10) And (YH1 = 1 Or YH1 = 4 Or YH1 = 7 Or YH1 = 10) Then
          ORB30Y = 3: ORB45Y = 5: ORB60Y = 7: ORB90Y = 10: ORB180Y = 12
          POW30Y = 3: POW45Y = 5: POW60Y = 7: POW90Y = 10: POW180Y = 12
          If y = 2 Then POW30Y = POW30Y + 1: POW45Y = POW45Y + 1: POW60Y = POW60Y + 1: POW90Y = POW90Y + 2: POW180Y = POW180Y + 3
       End If
       If (y = 1 Or y = 10) And (YH1 = 1 Or YH1 = 4 Or YH1 = 7 Or YH1 = 10) Then
          ORB30Y = 4: ORB45Y = 6: ORB60Y = 8: ORB90Y = 12: ORB180Y = 15
          POW30Y = 4: POW45Y = 6: POW60Y = 8: POW90Y = 12: POW180Y = 15
       End If
'
'find ORB for planet X, dependent upon what house X is in AND whether it is a Luminary or a planet
CAP2:   If x >= 14 Then
       ORB30X = 3: ORB45X = 5: ORB60X = 7: ORB90X = 10: ORB180X = 12
       POW30X = 3: POW45X = 5: POW60X = 7: POW90X = 10: POW180X = 12: GoTo CAP1
    End If
'using ONLY Sun and Moon as luminaries
       If (x <> 1 And x <> 10) And (XH1 = 3 Or XH1 = 6 Or XH1 = 9 Or XH1 = 12) Then
          ORB30X = 1: ORB45X = 3: ORB60X = 5: ORB90X = 6: ORB180X = 8
          POW30X = 1: POW45X = 3: POW60X = 5: POW90X = 6: POW180X = 8
          If x = 2 Then POW30X = POW30X + 1: POW45X = POW45X + 1: POW60X = POW60X + 1: POW90X = POW90X + 2: POW180X = POW180X + 3
       End If
       If (x = 1 Or x = 10) And (XH1 = 3 Or XH1 = 6 Or XH1 = 9 Or XH1 = 12) Then
          ORB30X = 2: ORB45X = 4: ORB60X = 6: ORB90X = 8: ORB180X = 11
          POW30X = 2: POW45X = 4: POW60X = 6: POW90X = 8: POW180X = 11
       End If
       If (x <> 1 And x <> 10) And (XH1 = 2 Or XH1 = 5 Or XH1 = 8 Or XH1 = 11) Then
          ORB30X = 2: ORB45X = 4: ORB60X = 6: ORB90X = 8: ORB180X = 10
          POW30X = 2: POW45X = 4: POW60X = 6: POW90X = 8: POW180X = 10
          If x = 2 Then POW30X = POW30X + 1: POW45X = POW45X + 1: POW60X = POW60X + 1: POW90X = POW90X + 2: POW180X = POW180X + 3
       End If
       If (x = 1 Or x = 10) And (XH1 = 2 Or XH1 = 5 Or XH1 = 8 Or XH1 = 11) Then
          ORB30X = 3: ORB45X = 5: ORB60X = 7: ORB90X = 10: ORB180X = 13
          POW30X = 3: POW45X = 5: POW60X = 7: POW90X = 10: POW180X = 13
       End If
       If (x <> 1 And x <> 10) And (XH1 = 1 Or XH1 = 4 Or XH1 = 7 Or XH1 = 10) Then
          ORB30X = 3: ORB45X = 5: ORB60X = 7: ORB90X = 10: ORB180X = 12
          POW30X = 3: POW45X = 5: POW60X = 7: POW90X = 10: POW180X = 12
          If x = 2 Then POW30X = POW30X + 1: POW45X = POW45X + 1: POW60X = POW60X + 1: POW90X = POW90X + 2: POW180X = POW180X + 3
       End If
       If (x = 1 Or x = 10) And (XH1 = 1 Or XH1 = 4 Or XH1 = 7 Or XH1 = 10) Then
          ORB30X = 4: ORB45X = 6: ORB60X = 8: ORB90X = 12: ORB180X = 15
          POW30X = 4: POW45X = 6: POW60X = 8: POW90X = 12: POW180X = 15
       End If
'
CAP1:   ORB30 = ORB30Y: If ORB30X >= ORB30Y Then ORB30 = ORB30X
    ORB45 = ORB45Y: If ORB45X >= ORB45Y Then ORB45 = ORB45X
    ORB60 = ORB60Y: If ORB60X >= ORB60Y Then ORB60 = ORB60X
    ORB90 = ORB90Y: If ORB90X >= ORB90Y Then ORB90 = ORB90X
    ORB180 = ORB180Y: If ORB180X >= ORB180Y Then ORB180 = ORB180X
    POW30 = POW30Y: If POW30X >= POW30Y Then POW30 = POW30X
    POW45 = POW45Y: If POW45X >= POW45Y Then POW45 = POW45X
    POW60 = POW60Y: If POW60X >= POW60Y Then POW60 = POW60X
    POW90 = POW90Y: If POW90X >= POW90Y Then POW90 = POW90X
    POW180 = POW180Y: If POW180X >= POW180Y Then POW180 = POW180X
    DA = Abs(Natal(y) - Natal(x)): If DA > 180 Then DA = 360 - DA
    Q = 1: K = DA
    If K <= ORB180 Then Q = 2: ORBXX = POW180: DAXX = DA: GoTo ACX3
    If K <= (30 + ORB30) And K >= (30 - ORB30) Then Q = 8: ORBXX = POW30: DAXX = DA - 30: GoTo ACX3
    If K <= (45 + ORB45) And K >= (45 - ORB45) Then Q = 9: ORBXX = POW45: DAXX = DA - 45: GoTo ACX3
    If K <= (60 + ORB60) And K >= (60 - ORB60) Then Q = 3: ORBXX = POW60: DAXX = DA - 60: GoTo ACX3
    If K <= (90 + ORB90) And K >= (90 - ORB90) Then Q = 4: ORBXX = POW90: DAXX = DA - 90: GoTo ACX3
'DA is added here to separate the overlap in the two aspects from 129 - 132 degrees for luminaries
    If DA <= 130 And K <= (120 + ORB90) And K >= (120 - ORB90) Then Q = 5: ORBXX = POW90: DAXX = DA - 120: GoTo ACX3
    If DA > 130 And K <= (135 + ORB45) And K >= (135 - ORB45) Then Q = 10: ORBXX = POW45: DAXX = DA - 135: GoTo ACX3
    If K <= (150 + ORB30) And K >= (150 - ORB30) Then Q = 11: ORBXX = POW30: DAXX = DA - 150: GoTo ACX3
    If K >= (180 - ORB180) Then Q = 6: ORBXX = POW180: DAXX = DA - 180
ACX3:   If Q = 1 Or Q > 11 Then GoTo PTT4
    PlanetAspectPower(y) = PlanetAspectPower(y) + ORBXX - Abs(DAXX)
    If y < x Then Cosmo!(y, x) = Cosmo!(y, x) + ORBXX - Abs(DAXX)
'get planetary harmony and discord
    If Q = 5 Or Q = 3 Or Q = 8 Then
       HarmonyX = ORBXX - Abs(DAXX)
       PlanetHarmony(y) = PlanetHarmony(y) + HarmonyX
       If y > x Then Cosmo!(y, x) = Cosmo!(y, x) + HarmonyX
    End If
    If Q = 4 Or Q = 9 Or Q = 10 Or Q = 6 Then
       DiscordX = ORBXX - Abs(DAXX)
       PlanetDiscord(y) = PlanetDiscord(y) + DiscordX
       If y > x Then Cosmo!(y, x) = Cosmo!(y, x) - DiscordX
    End If
    If x = 5 Or y = 5 Then
       HarmonyX = ((ORBXX - Abs(DAXX)) / 2)
       PlanetHarmony(y) = PlanetHarmony(y) + HarmonyX
       If y > x Then Cosmo!(y, x) = Cosmo!(y, x) + HarmonyX
    End If
    If x = 3 Or y = 3 Then
       HarmonyX = ((ORBXX - Abs(DAXX)) / 4)
       PlanetHarmony(y) = PlanetHarmony(y) + HarmonyX
       If y > x Then Cosmo!(y, x) = Cosmo!(y, x) + HarmonyX
    End If
    If x = 6 Or y = 6 Then
       DiscordX = ((ORBXX - Abs(DAXX)) / 2)
       PlanetDiscord(y) = PlanetDiscord(y) + DiscordX
       If y > x Then Cosmo!(y, x) = Cosmo!(y, x) - DiscordX
    End If
    If x = 4 Or y = 4 Then
       DiscordX = ((ORBXX - Abs(DAXX)) / 4)
       PlanetDiscord(y) = PlanetDiscord(y) + DiscordX
       If y > x Then Cosmo!(y, x) = Cosmo!(y, x) - DiscordX
    End If

PTT4:   DiffDecl = Abs(Abs(NatalDecl(y)) - Abs(NatalDecl(x)))
    If DiffDecl < 1 Then
       If (y <> 1 And y <> 2 And y <> 10) And (YH1 = 3 Or YH1 = 6 Or YH1 = 9 Or YH1 = 12) Then
          ORB180Y = 8
       End If
       If (y = 1 Or y = 2 Or y = 10) And (YH1 = 3 Or YH1 = 6 Or YH1 = 9 Or YH1 = 12) Then
          ORB180Y = 11
       End If
       If (y <> 1 And y <> 2 And y <> 10) And (YH1 = 2 Or YH1 = 5 Or YH1 = 8 Or YH1 = 11) Then
          ORB180Y = 10
       End If
       If (y = 1 Or y = 2 Or y = 10) And (YH1 = 2 Or YH1 = 5 Or YH1 = 8 Or YH1 = 11) Then
          ORB180Y = 13
       End If
       If (y <> 1 And y <> 2 And y <> 10) And (YH1 = 1 Or YH1 = 4 Or YH1 = 7 Or YH1 = 10) Then
          ORB180Y = 12
       End If
       If (y = 1 Or y = 2 Or y = 10) And (YH1 = 1 Or YH1 = 4 Or YH1 = 7 Or YH1 = 10) Then
          ORB180Y = 15
       End If
'
       If (x <> 1 And x <> 2 And x <> 10) And (XH1 = 3 Or XH1 = 6 Or XH1 = 9 Or XH1 = 12) Then
          ORB180X = 8
       End If
       If (x = 1 Or x = 2 Or x = 10) And (XH1 = 3 Or XH1 = 6 Or XH1 = 9 Or XH1 = 12) Then
          ORB180X = 11
       End If
       If (x <> 1 And x <> 2 And x <> 10) And (XH1 = 2 Or XH1 = 5 Or XH1 = 8 Or XH1 = 11) Then
          ORB180X = 10
       End If
       If (x = 1 Or x = 2 Or x = 10) And (XH1 = 2 Or XH1 = 5 Or XH1 = 8 Or XH1 = 11) Then
          ORB180X = 13
       End If
       If (x <> 1 And x <> 2 And x <> 10) And (XH1 = 1 Or XH1 = 4 Or XH1 = 7 Or XH1 = 10) Then
          ORB180X = 12
       End If
       If (x = 1 Or x = 2 Or x = 10) And (XH1 = 1 Or XH1 = 4 Or XH1 = 7 Or XH1 = 10) Then
          ORB180X = 15
       End If
'
       ORB180 = ORB180Y: If ORB180X >= ORB180Y Then ORB180 = ORB180X
       DeclPower = ORB180 * (1 - DiffDecl)
       PlanetAspectPower(y) = PlanetAspectPower(y) + DeclPower
       If y < x Then Cosmo!(y, x) = Cosmo!(y, x) + DeclPower
       If x = 5 Or y = 5 Then
          HarmonyX = DeclPower / 2
          PlanetHarmony(y) = PlanetHarmony(y) + HarmonyX
          If y > x Then Cosmo!(y, x) = Cosmo!(y, x) + HarmonyX
       End If
       If x = 3 Or y = 3 Then
          HarmonyX = DeclPower / 4
          PlanetHarmony(y) = PlanetHarmony(y) + HarmonyX
          If y > x Then Cosmo!(y, x) = Cosmo!(y, x) + HarmonyX
       End If
       If x = 6 Or y = 6 Then
          DiscordX = DeclPower / 2
          PlanetDiscord(y) = PlanetDiscord(y) + DiscordX
          If y > x Then Cosmo!(y, x) = Cosmo!(y, x) - DiscordX
       End If
       If x = 4 Or y = 4 Then
          DiscordX = DeclPower / 4
          PlanetDiscord(y) = PlanetDiscord(y) + DiscordX
          If y > x Then Cosmo!(y, x) = Cosmo!(y, x) - DiscordX
       End If
    End If
PTT2:   Next x
PTT3:   Next y
End Sub