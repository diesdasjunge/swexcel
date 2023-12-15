Public Function Get_Ascendant_decl(m%, d%, yr%, hr%, n%, tz!, lat As Single, lon As Single) As Double
    Dim cal As Byte, h As Double, xx As Double, y As Double, z As Double
    Dim ascmc(10) As Double, cusp(13) As Double, x(6) As Double, xpin(2) As Double, xecl(6) As Double

    'swe_set_ephe_path (Application.path)

    imonth% = m%
    iday% = d%
    iyear% = yr%
    
    ihour% = hr% - tz!
    imin% = n%
    
'get hours in pure decimal form
    h = ihour% + imin% / 60#

'the next two functions do the same job, converting a calendar date into a Julian day number
'swe_date_conversion() checks for legal dates while swe_julday() handles even illegal things like 45 Januar etc.
    cal = 103  ' g for gregorian calendar

    tjd_ut = swe_julday(iyear%, imonth%, iday%, h, 1)
    retval = swe_date_conversion(iyear, imonth, iday, h, cal, tjd_ut)
    tjd_et = tjd_ut
    tjd_ut = tjd_et

    'calculate obliquity of ecliptic
    retval = swe_calc_ut(tjd_ut, -1, 0, xecl(0), serr$)
    epsilon = xecl(0)       'in degrees

'set Placidus houses, then get house cusps
    t$ = "P"
    ret_flag = swe_houses(tjd_ut, lat, lon, Asc(t$), cusp(0), ascmc(0))
    
    OB = -Radians(epsilon)
    
    xx = COSINE(ascmc(0))
    y = Cos(-OB) * Sine(ascmc(0))
    z = Sin(-OB) * Sine(ascmc(0))
    
    Get_Ascendant_decl = Cart2Sph(xx, y, z)
End Function