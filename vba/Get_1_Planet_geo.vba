Function Get_1_Planet_geo(givenJD As Double, p_num%)
    Dim ret_flag As Long, serr$, x(6) As Double
    Dim et As Double, ut As Double
    Dim iflag As Long, ephem_num&

    ut = givenJD

    et = ut + swe_deltat(givenJD)

    serr$ = String(255, 0)

    iflag = 2 + 256

    ephem_num& = p_num%
    
    ret_flag = swe_calc(et, ephem_num&, iflag, x(0), serr$)

    current_speed = x(3)
    Get_1_Planet_geo = x(0)
End Function