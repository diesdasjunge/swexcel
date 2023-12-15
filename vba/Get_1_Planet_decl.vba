Public Function Get_1_Planet_decl(p_num%, m%, d%, yr%, hr%, n%, ss%, tz!) As Double
    Dim cal As Byte, h As Double, iflag As Long, ret_flag As Long, serr$, x(6) As Double, y%, z%

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

    iflag = 2 + 256 + 2048
    ret_flag = swe_calc(tjd_ut + swe_deltat(tjd_ut), p_num%, iflag, x(0), serr$)
    
    Get_1_Planet_decl = x(1)
End Function