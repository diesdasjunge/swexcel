Public Function Get_natal_data(m%, d%, yr%, hr%, n%, ss%, tz!, lat As Single, lon As Single) As Double
    Dim cal As Byte, h As Double, iflag As Long, ret_flag As Long, serr$, x(6) As Double, y%, z%
    Dim ascmc(10) As Double, cusp(13) As Double, xecl(6) As Double
    
    'swe_set_ephe_path (Application.path)

    imonth% = m%
    iday% = d%
    iyear% = yr%
    
    ihour% = hr% - tz!
    imin% = n%
    
'get hours in pure decimal form
    h = ihour% + imin% / 60# + ss% / 3600

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
    
    Call PrintCosmodynes
    
    
    Get_natal_data = PlanetTotalPower(1)
    
    
    
'    Cells(8, 1) = "Hello World"
    
    
'    Dim MyXL As Object
'    Dim mySheet As Excel.Worksheet
    
'    Set MyXL = GetObject(, "Excel.Application")
    
'    Set mySheet = MyXL.Application.ActiveSheet
    
'    Range("B5") = "A x B"
    
'    With mySheet.Cells(3, 4)
        '.Text = "43.338"
'    End With
    
'    Sheets("Sheet1").Cells(3, 4).Value = 165.34
    
'    Get_natal_data = PlanetTotalPower(1)

    
    'Application.CalculateFull           'forces full calculations
End Function