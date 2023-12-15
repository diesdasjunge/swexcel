Public Sub Get_Initial_Planet_Positions(ByVal earlier_jd As Double, ByVal later_jd As Double, p_num%, end_pos)
'get positions of both planets on JD = later_jd and JD = earlier_jd
    p1_pos_later = Get_1_Planet_geo(later_jd, p_num%)
    p1_pos = Get_1_Planet_geo(earlier_jd, p_num%)

    p2_pos_later = end_pos
    p2_pos = end_pos
End Sub