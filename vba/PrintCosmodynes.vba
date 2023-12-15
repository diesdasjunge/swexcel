Sub PrintCosmodynes()
    Dim cal As Byte, h As Double, iflag As Long, ret_flag As Long, serr$, x(6) As Double, xx As Single, y%, z%
    Dim ascmc(10) As Double, cusp(13) As Double, xecl(6) As Double
    
    crlf = Chr$(13) & Chr$(10)
    LineCounter% = 0
    
    TextLineCounter% = 0            'clear buffer for form
    ReDim TextOnForm$(0 To 1)
    
    PX1$ = "SUN       MERCURY   VENUS     MARS      JUPITER   SATURN    URANUS    NEPTUNE   "
    PX1$ = PX1$ & "PLUTO     MOON      N. NODE   S. NODE   P. OF F.  ASC       MC        "

    SX1$ = "ARITAUGEMCANLEOVIRLIBSCOSAGCPRAQUPIS"
    
    'swe_set_ephe_path (Application.Path)
   
  

    imonth% = Sheet1.Range("a2")
    iday% = Sheet1.Range("b2")
    iyear% = Sheet1.Range("c2")
    
    ihour% = Sheet1.Range("d2") - Sheet1.Range("g2")
    imin% = Sheet1.Range("e2")
    
'get hours in pure decimal form
    h = ihour% + imin% / 60# + Sheet1.Range("f2") / 3600
    
    lon = Sheet1.Range("h2")
    lat = Sheet1.Range("i2")

'the next two functions do the same job, converting a calendar date into a Julian day number
'swe_date_conversion() checks for legal dates while swe_julday() handles even illegal things like 45 Januar etc.
    cal = 103  ' g for gregorian calendar

    tjd_ut = swe_julday(iyear%, imonth%, iday%, h, 1)
    retval = swe_date_conversion(iyear, imonth, iday, h, cal, tjd_ut)

    
    'calculate obliquity of ecliptic
    retval = swe_calc_ut(tjd_ut, -1, 0, xecl(0), serr$)
    epsilon = xecl(0)       'in degrees


    For z% = 0 To 9
        If z% = 0 Then
            y% = 1
        ElseIf z% = 1 Then
            y% = 10
        Else
            y% = z%
        End If
        
        'first, get planet longitudes
        iflag = 2 + 256
        ret_flag = swe_calc(tjd_ut + swe_deltat(tjd_ut), z%, iflag, x(0), serr$)
        Natal(y%) = x(0)
        Planets(y%) = x(0)
    
        'first, get planet declinations
        iflag = 2 + 256 + 2048
        ret_flag = swe_calc(tjd_ut + swe_deltat(tjd_ut), z%, iflag, x(0), serr$)
        NatalDecl(y%) = x(1)
    Next z%
    
    'get house cusps (Placidus)
    t$ = "P"

    ret_flag = swe_houses(tjd_ut, lat, lon, Asc(t$), cusp(0), ascmc(0))         'tropical

    For y% = 1 To 12
        HC3(y%) = cusp(y%)
    Next y%

    For y% = 1 To 9
        HC(y% + 3) = HC3(y%)
    Next y%

    For y% = 10 To 12
        HC(y% - 9) = HC3(y%)
    Next y%


    Natal(14) = HC(4)
    Natal(15) = HC(1)
    
        
    For y% = 1 To 13
        NatalHouse(y%) = HC(y%)
    Next y%
    
    
    'find house positions for each planet
    HCX1$ = "10TH11TH12TH 1ST 2ND 3RD 4TH 5TH 6TH 7TH 8TH 9TH    "
    PH$ = Space(150)
    
    Call PLANETSINHOUSES
    
    NatalPH$ = PH$
    
    
    Call AngleDeclinations      'get declination of Ascendant and Midheaven
    
    
    'set up to calculate cosmodynes
    PN$(1) = "SUN": PN$(2) = "MERCURY": PN$(3) = "VENUS"
    PN$(4) = "MARS": PN$(5) = "JUPITER": PN$(6) = "SATURN"
    PN$(7) = "URANUS": PN$(8) = "NEPTUNE": PN$(9) = "PLUTO"
    PN$(10) = "MOON": PN$(11) = "N. NODE": PN$(12) = "S. NODE":
    PN$(13) = "PART F.": PN$(14) = "ASC": PN$(15) = "MC "
    
    
    ReDim SORT(15), SORTPOS(15)
    ReDim XHousePower!(12), XHousePowerVariation!(12)

    XHousePower!(1) = 15: XHousePower!(2) = 8.5: XHousePower!(3) = 8
    XHousePower!(4) = 14: XHousePower!(5) = 7.5: XHousePower!(6) = 7
    XHousePower!(7) = 14.5: XHousePower!(8) = 10.9: XHousePower!(9) = 10
    XHousePower!(10) = 15: XHousePower!(11) = 11.9: XHousePower!(12) = 9.3
    
    XHousePowerVariation!(1) = 2: XHousePowerVariation!(2) = 0.5
    XHousePowerVariation!(3) = 0.5: XHousePowerVariation!(4) = 2
    XHousePowerVariation!(5) = 0.5: XHousePowerVariation!(6) = 0.5
    XHousePowerVariation!(7) = 2: XHousePowerVariation!(8) = 0.9
    XHousePowerVariation!(9) = 0.7: XHousePowerVariation!(10) = 2
    XHousePowerVariation!(11) = 1: XHousePowerVariation!(12) = 0.7
    
    
    ReDim DegreesInEachHouse(12)
    ReDim DistanceFromCusp(10), PlanetHousePower(15), PlanetAspectPower(15)
    ReDim PlanetHarmony(15), PlanetDiscord(15), PlanetTotalPower(15)
    ReDim InterceptedSign%(12), InterceptedHouse%(12), Cosmo!(15, 15)
    ReDim SignPower(12), SignHarmony(12), SignDiscord(12)
    ReDim HousePower(12), HouseHarmony(12), HouseDiscord(12)
    
    For xx = 10 To 12: HC3(xx) = NatalHouse(xx - 9): Next xx
    For xx = 1 To 9: HC3(xx) = NatalHouse(xx + 3): Next xx

    For xx = 1 To 10
        y = Val(Mid$(NatalPH$, xx * 10 - 9, 2))
        
        DistanceFromCusp(xx) = Natal(xx) - HC3(y)
        If DistanceFromCusp(xx) < 0 Then DistanceFromCusp(xx) = DistanceFromCusp(xx) + 360
    Next xx

    Call DegInEachHouse
    Call CalcPlanetHousePower
    Call CalcPlanetAspectPower
    
    For xx = 1 To 15
        PlanetTotalPower(xx) = PlanetHousePower(xx) + PlanetAspectPower(xx)
    Next xx
    
    Call Reception
    Call Dignities
    Call FigureSignInterceptions
    Call FigureHouseInterceptions
    Call CalcSignPower
    Call CalcHousePower
    
'sort the planets according to highest POWER ranking.
    TT1 = 0: TT2 = 0: TT3 = 0: TT4 = 0: TT5 = 0: TT6 = 0
    For xx = 1 To 15
        SORT(xx) = PlanetTotalPower(xx)
        SORTPOS(xx) = xx
    Next xx
    
    For xx = 1 To 14
        For y = xx + 1 To 15
            If SORT(y) > SORT(xx) Then
                TEMP = SORT(xx): TEMP1 = SORTPOS(xx)
                SORT(xx) = SORT(y): SORTPOS(xx) = SORTPOS(y)
                SORT(y) = TEMP: SORTPOS(y) = TEMP1
            End If
        Next y
    Next xx
    
    Call LineFeed
    
    Buffer = Space(2) & "Planet" & Space(13) & "Power    %" & Space(9) & "Har/Disc"
    Call PrintBuffer
'
    For xx = 1 To 12: TT1 = TT1 + PlanetTotalPower(SORTPOS(xx))
    TT2 = TT2 + PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx))
    TT3 = TT3 + PlanetHarmony(SORTPOS(xx)): TT4 = TT4 + PlanetDiscord(SORTPOS(xx))
    If PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx)) >= 0 Then TT5 = TT5 + (PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx)))
    If PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx)) < 0 Then TT6 = TT6 - (PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx)))
    Next xx
'
    For xx = 1 To 12
        Buffer = Space(2) & Mid$(PX1$, SORTPOS(xx) * 10 - 9, 10) & Space(7)
        Buffer = Buffer & Align4(Format$(PlanetTotalPower(SORTPOS(xx)), "###0.00")) & Space(2)
        Buffer = Buffer & Align2(Format$(100 * PlanetTotalPower(SORTPOS(xx)) / TT1, "#0.0")) & Space(7)

        Buffer = Buffer & Align4(Format$(PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx)), "###0.00"))
        Call PrintBuffer
    Next xx

    Call LineFeed: Buffer = Space(2) & "TOTALS" & Space(11)
    Buffer = Buffer & Align4(Format$(TT1, "###0.00")) & Space(13) & Align4(Format$(TT2, "###0.00"))
    Call PrintBuffer

'sort the planets according to highest OVERALL HARMONY/DISCORD ranking.
    For xx = 1 To 15: SORT(xx) = PlanetHarmony(xx) - PlanetDiscord(xx): SORTPOS(xx) = xx: Next xx
    For xx = 1 To 14: For y = xx + 1 To 15
    If SORT(y) > SORT(xx) Then
       TEMP = SORT(xx): TEMP1 = SORTPOS(xx)
       SORT(xx) = SORT(y): SORTPOS(xx) = SORTPOS(y)
       SORT(y) = TEMP: SORTPOS(y) = TEMP1
    End If
    Next y: Next xx

    Buffer = Space(51) & "%": Call PrintBuffer
    For xx = 1 To 15
        If SORTPOS(xx) = 11 Or SORTPOS(xx) = 12 Or SORTPOS(xx) = 13 Then GoTo PCX1
        Buffer = Space(2) & Mid$(PX1$, SORTPOS(xx) * 10 - 9, 10) & Space(7)
        Buffer = Buffer & Align4(Format$(PlanetTotalPower(SORTPOS(xx)), "###0.00")) & Space(13)
        Buffer = Buffer & Align4(Format$(PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx)), "###0.00")) & Space(3)

        If PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx)) >= 0 Then
        Buffer = Buffer & Align2(Format$(100 * (PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx))) / TT5, "#0.0")) & Space(5)
        End If
        If PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx)) < 0 Then
        Buffer = Buffer & Align2(Format$(Abs(100 * (PlanetHarmony(SORTPOS(xx)) - PlanetDiscord(SORTPOS(xx))) / TT6), "#0.0")) & Space(5)
        End If
        Call PrintBuffer
PCX1:   Next xx

'sort the signs according to highest POWER ranking.
    TT1 = 0: TT2 = 0: TT3 = 0: TT4 = 0: TT5 = 0: TT6 = 0
    For xx = 1 To 12: SORT(xx) = SignPower(xx): SORTPOS(xx) = xx: Next xx
    For xx = 1 To 11: For y = xx + 1 To 12
    If SORT(y) > SORT(xx) Then
       TEMP = SORT(xx): TEMP1 = SORTPOS(xx)
       SORT(xx) = SORT(y): SORTPOS(xx) = SORTPOS(y)
       SORT(y) = TEMP: SORTPOS(y) = TEMP1
    End If
    Next y: Next xx
    Call LineFeed
    
    Buffer = Space(2) & "Sign" & Space(15) & "Power    %" & Space(9) & "Har/Disc"
    Call PrintBuffer

    For xx = 1 To 12: TT1 = TT1 + SignPower(SORTPOS(xx))
    TT2 = TT2 + SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx))
    TT3 = TT3 + SignHarmony(SORTPOS(xx)): TT4 = TT4 + SignDiscord(SORTPOS(xx))
    If SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx)) >= 0 Then TT5 = TT5 + (SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx)))
    If SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx)) < 0 Then TT6 = TT6 - (SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx)))
    Next xx

    For xx = 1 To 12
    Buffer = Space(2) & Mid$(SX1$, SORTPOS(xx) * 3 - 2, 3) & Space(14)
    Buffer = Buffer & Align4(Format$(SignPower(SORTPOS(xx)), "###0.00")) & Space(2)
    Buffer = Buffer & Align2(Format$(100 * SignPower(SORTPOS(xx)) / TT1, "#0.0")) & Space(7)

    Buffer = Buffer & Align4(Format$(SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx)), "###0.00"))
    Call PrintBuffer
    Next xx

    Call LineFeed: Buffer = Space(2) & "TOTALS" & Space(11)
    Buffer = Buffer & Align4(Format$(TT1, "###0.00")) & Space(13) & Align4(Format$(TT2, "###0.00"))
    Call PrintBuffer

'sort the signs according to highest OVERALL HARMONY/DISCORD ranking.
    For xx = 1 To 12: SORT(xx) = SignHarmony(xx) - SignDiscord(xx): SORTPOS(xx) = xx: Next xx
    For xx = 1 To 11: For y = xx + 1 To 12
    If SORT(y) > SORT(xx) Then
       TEMP = SORT(xx): TEMP1 = SORTPOS(xx)
       SORT(xx) = SORT(y): SORTPOS(xx) = SORTPOS(y)
       SORT(y) = TEMP: SORTPOS(y) = TEMP1
    End If
    Next y: Next xx

    Buffer = Space(51) & "%": Call PrintBuffer
    For xx = 1 To 12
        Buffer = Space(2) & Mid$(SX1$, SORTPOS(xx) * 3 - 2, 3) & Space(14)
        Buffer = Buffer & Align4(Format$(SignPower(SORTPOS(xx)), "###0.00")) & Space(13)
        Buffer = Buffer & Align4(Format$(SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx)), "###0.00")) & Space(3)

        If SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx)) >= 0 Then
        Buffer = Buffer & Align2(Format$(100 * (SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx))) / TT5, "#0.0")) & Space(5)
        End If
        If SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx)) < 0 Then
        Buffer = Buffer & Align2(Format$(Abs(100 * (SignHarmony(SORTPOS(xx)) - SignDiscord(SORTPOS(xx))) / TT6), "#0.0")) & Space(5)
        End If
        Call PrintBuffer
    Next xx

'sort the houses according to highest POWER ranking.
    TT1 = 0: TT2 = 0: TT3 = 0: TT4 = 0: TT5 = 0: TT6 = 0
    For xx = 1 To 12: SORT(xx) = HousePower(xx): SORTPOS(xx) = xx: Next xx
    For xx = 1 To 11: For y = xx + 1 To 12
    If SORT(y) > SORT(xx) Then
       TEMP = SORT(xx): TEMP1 = SORTPOS(xx)
       SORT(xx) = SORT(y): SORTPOS(xx) = SORTPOS(y)
       SORT(y) = TEMP: SORTPOS(y) = TEMP1
    End If
    Next y: Next xx: Call LineFeed

    Buffer = Space(2) & "House" & Space(14) & "Power    %" & Space(9) & "Har/Disc"
    Call PrintBuffer

    For xx = 1 To 12: TT1 = TT1 + HousePower(SORTPOS(xx))
    TT2 = TT2 + HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx))
    TT3 = TT3 + HouseHarmony(SORTPOS(xx)): TT4 = TT4 + HouseDiscord(SORTPOS(xx))
    If HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx)) >= 0 Then TT5 = TT5 + (HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx)))
    If HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx)) < 0 Then TT6 = TT6 - (HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx)))
    Next xx

    For xx = 1 To 12
    Buffer = Space(3)
    If SORTPOS(xx) < 10 Then
        Buffer = Buffer & Space(1)
    End If
    Buffer = Buffer & Str$(SORTPOS(xx)) & Space(13)
    Buffer = Buffer & Align4(Format$(HousePower(SORTPOS(xx)), "###0.00")) & Space(2)
    Buffer = Buffer & Align2(Format$(100 * HousePower(SORTPOS(xx)) / TT1, "#0.0")) & Space(7)

    Buffer = Buffer & Align4(Format$(HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx)), "###0.00")) & Space(12)
    Call PrintBuffer
    Next xx

    Call LineFeed: Buffer = Space(2) & "TOTALS" & Space(11)
    Buffer = Buffer & Align4(Format$(TT1, "###0.00")) & Space(13) & Align4(Format$(TT2, "###0.00"))
    Call PrintBuffer

'sort the houses according to highest OVERALL HARMONY/DISCORD ranking.
    For xx = 1 To 12: SORT(xx) = HouseHarmony(xx) - HouseDiscord(xx): SORTPOS(xx) = xx: Next xx
    For xx = 1 To 11: For y = xx + 1 To 12
    If SORT(y) > SORT(xx) Then
       TEMP = SORT(xx): TEMP1 = SORTPOS(xx)
       SORT(xx) = SORT(y): SORTPOS(xx) = SORTPOS(y)
       SORT(y) = TEMP: SORTPOS(y) = TEMP1
    End If
    Next y: Next xx

    Buffer = Space(51) & "%": Call PrintBuffer
    For xx = 1 To 12
        Buffer = Space(3) & Align2(Str$(SORTPOS(xx))) & Space(13)
        Buffer = Buffer & Align4(Format$(HousePower(SORTPOS(xx)), "###0.00")) & Space(13)
        Buffer = Buffer & Align4(Format$(HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx)), "###0.00")) & Space(3)

        If HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx)) >= 0 Then
        Buffer = Buffer & Align2(Format$(100 * (HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx))) / TT5, "#0.0")) & Space(5)
        End If
        If HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx)) < 0 Then
        Buffer = Buffer & Align2(Format$(Abs(100 * (HouseHarmony(SORTPOS(xx)) - HouseDiscord(SORTPOS(xx))) / TT6), "#0.0")) & Space(5)
        End If
        Call PrintBuffer
    Next xx: Call LineFeed: Call LineFeed

'List the various other data that can be calculated
    Buffer = Space(2) & "Signs" & Space(14) & "Power" & Space(14) & "Har/Disc"
    Call PrintBuffer
    
    Buffer = Space(2) & "Fire" & Space(13)
    Buffer = Buffer & Align4(Format$(SignPower(1) + SignPower(5) + SignPower(9), "###0.00")) & Space(13)
    Buffer = Buffer & Align4(Format$(SignHarmony(1) + SignHarmony(5) + SignHarmony(9) - (SignDiscord(1) + SignDiscord(5) + SignDiscord(9)), "###0.00"))
    Call PrintBuffer
    
    Buffer = Space(2) & "Earth" & Space(12)
    Buffer = Buffer & Align4(Format$(SignPower(2) + SignPower(6) + SignPower(10), "###0.00")) & Space(13)
    Buffer = Buffer & Align4(Format$(SignHarmony(2) + SignHarmony(6) + SignHarmony(10) - (SignDiscord(2) + SignDiscord(6) + SignDiscord(10)), "###0.00"))
    Call PrintBuffer

    Buffer = Space(2) & "Air" & Space(14)
    Buffer = Buffer & Align4(Format$(SignPower(3) + SignPower(7) + SignPower(11), "###0.00")) & Space(13)
    Buffer = Buffer & Align4(Format$(SignHarmony(3) + SignHarmony(7) + SignHarmony(11) - (SignDiscord(3) + SignDiscord(7) + SignDiscord(11)), "###0.00"))
    Call PrintBuffer

    Buffer = Space(2) & "Water" & Space(12)
    Buffer = Buffer & Align4(Format$(SignPower(4) + SignPower(8) + SignPower(12), "###0.00")) & Space(13)
    Buffer = Buffer & Align4(Format$(SignHarmony(4) + SignHarmony(8) + SignHarmony(12) - (SignDiscord(4) + SignDiscord(8) + SignDiscord(12)), "###0.00"))
    Call PrintBuffer

    Call LineFeed
    Buffer = Space(2) & "Cardinal" & Space(9)
    Buffer = Buffer & Align4(Format$(SignPower(1) + SignPower(4) + SignPower(7) + SignPower(10), "###0.00")) & Space(13)
    Buffer = Buffer & Align4(Format$(SignHarmony(1) + SignHarmony(4) + SignHarmony(7) + SignHarmony(10) - (SignDiscord(1) + SignDiscord(4) + SignDiscord(7) + SignDiscord(10)), "###0.00"))
    Call PrintBuffer

    Buffer = Space(2) & "Fixed" & Space(12)
    Buffer = Buffer & Align4(Format$(SignPower(2) + SignPower(5) + SignPower(8) + SignPower(11), "###0.00")) & Space(13)
    Buffer = Buffer & Align4(Format$(SignHarmony(2) + SignHarmony(5) + SignHarmony(8) + SignHarmony(11) - (SignDiscord(2) + SignDiscord(5) + SignDiscord(8) + SignDiscord(11)), "###0.00"))
    Call PrintBuffer

    Buffer = Space(2) & "Common" & Space(11)
    Buffer = Buffer & Align4(Format$(SignPower(3) + SignPower(6) + SignPower(9) + SignPower(12), "###0.00")) & Space(13)
    Buffer = Buffer & Align4(Format$(SignHarmony(3) + SignHarmony(6) + SignHarmony(9) + SignHarmony(12) - (SignDiscord(3) + SignDiscord(6) + SignDiscord(9) + SignDiscord(12)), "###0.00"))
    Call PrintBuffer

    SouthPower = HousePower(7) + HousePower(8) + HousePower(9) + HousePower(10) + HousePower(11) + HousePower(12)
    SouthHarmony = HouseHarmony(7) + HouseHarmony(8) + HouseHarmony(9) + HouseHarmony(10) + HouseHarmony(11) + HouseHarmony(12)
    SouthDiscord = HouseDiscord(7) + HouseDiscord(8) + HouseDiscord(9) + HouseDiscord(10) + HouseDiscord(11) + HouseDiscord(12)
    NorthPower = HousePower(1) + HousePower(2) + HousePower(3) + HousePower(4) + HousePower(5) + HousePower(6)
    NorthHarmony = HouseHarmony(1) + HouseHarmony(2) + HouseHarmony(3) + HouseHarmony(4) + HouseHarmony(5) + HouseHarmony(6)
    NorthDiscord = HouseDiscord(1) + HouseDiscord(2) + HouseDiscord(3) + HouseDiscord(4) + HouseDiscord(5) + HouseDiscord(6)
    EastPower = HousePower(10) + HousePower(11) + HousePower(12) + HousePower(1) + HousePower(2) + HousePower(3)
    EastHarmony = HouseHarmony(10) + HouseHarmony(11) + HouseHarmony(12) + HouseHarmony(1) + HouseHarmony(2) + HouseHarmony(3)
    EastDiscord = HouseDiscord(10) + HouseDiscord(11) + HouseDiscord(12) + HouseDiscord(1) + HouseDiscord(2) + HouseDiscord(3)
    WestPower = HousePower(4) + HousePower(5) + HousePower(6) + HousePower(7) + HousePower(8) + HousePower(9)
    WestHarmony = HouseHarmony(4) + HouseHarmony(5) + HouseHarmony(6) + HouseHarmony(7) + HouseHarmony(8) + HouseHarmony(9)
    WestDiscord = HouseDiscord(4) + HouseDiscord(5) + HouseDiscord(6) + HouseDiscord(7) + HouseDiscord(8) + HouseDiscord(9)
    
    Call LineFeed: Buffer = Space(2) & "South/North" & Space(6)
    Buffer = Buffer & Align4(Format$(SouthPower, "###0.00") & " / " & Format$(NorthPower, "###0.00")) & Space(4)
    Buffer = Buffer & Align4(Format$(SouthHarmony - SouthDiscord, "###0.00") & " / " & Format$(NorthHarmony - NorthDiscord, "###0.00"))
    Call PrintBuffer

    Buffer = Space(2) & "East/West" & Space(8)
    Buffer = Buffer & Align4(Format$(EastPower, "###0.00") & " / " & Format$(WestPower, "###0.00")) & Space(4)
    Buffer = Buffer & Align4(Format$(EastHarmony - EastDiscord, "###0.00") & " / " & Format$(WestHarmony - WestDiscord, "###0.00"))
    
    Call PrintBuffer
    
    Call PrintPowerHarmonyTable
    
    Call AdjustForm
End Sub