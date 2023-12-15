Public Function Get_MC(m%, d%, yr%, hr%, n%, tz!, lat As Single, lon As Single) As Double
    Dim cal As Byte, h As Double, y%, z%
    Dim ascmc(10) As Double, cusp(13) As Double, x(6) As Double, xpin(2) As Double

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

'set Placidus houses, then get house cusps
    t$ = "P"
    ret_flag = swe_houses(tjd_ut, lat, lon, Asc(t$), cusp(0), ascmc(0))
    
    Get_MC = ascmc(1)
End Function