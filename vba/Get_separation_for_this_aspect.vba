Public Function Get_separation_for_this_aspect()
    Dim DA As Double, DAYY As Double, d As Double

    'get distance from exact aspect for both planets on JD = later_jd
    DAYY = p2_pos_later - p1_pos_later
    DA = Abs(p2_pos_later - p1_pos_later)
    If DA > 180 Then DA = 360 - DA
    sep_t1 = DA
    If DAYY <= -180 Or (DAYY >= 0 And DAYY < 180) Then sep_t1 = -sep_t1

    'get distance from exact aspect for both planets on JD = earlier_jd
    DAYY = p2_pos - p1_pos
    DA = Abs(p2_pos - p1_pos)
    If DA > 180 Then DA = 360 - DA
    sep_t0 = DA
    If DAYY <= -180 Or (DAYY >= 0 And DAYY < 180) Then sep_t0 = -sep_t0
    
    If sep_t1 - sep_t0 = 0 Then
        d = 1000
    Else
        d = sep_t1 / (sep_t1 - sep_t0)     'since our interval is always 1 day
    End If

    Get_separation_for_this_aspect = d
End Function