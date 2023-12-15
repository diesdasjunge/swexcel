Sub SecantMethod(ByVal earlier_jd As Double, ByVal later_jd As Double, ByVal e As Double, ByVal m As Integer, p_num%, end_pos)
    Dim DA As Double, DAYY As Double, n As Integer, d As Double
    Dim y1 As Double, y2 As Double, y3 As Double, y4 As Double
    Dim dist1 As Double, dist2 As Double

    For n = 1 To m
        'get positions of both planets on JD = later_jd and JD = earlier_jd
        y1 = Get_1_Planet_geo(later_jd, p_num%)
        p1_current_speed = current_speed
        y3 = Get_1_Planet_geo(earlier_jd, p_num%)

        y2 = end_pos
        y4 = end_pos

        'get distance from exact aspect for both planets on JD = later_jd
        DAYY = y2 - y1
        DA = Abs(y2 - y1)
        If DA > 180 Then DA = 360 - DA
        dist1 = DA
        If DAYY <= -180 Or (DAYY >= 0 And DAYY < 180) Then dist1 = -dist1

        'get distance from exact aspect for both planets on JD = earlier_jd
        DAYY = y4 - y3
        DA = Abs(y4 - y3)
        If DA > 180 Then DA = 360 - DA
        dist2 = DA
        If DAYY <= -180 Or (DAYY >= 0 And DAYY < 180) Then dist2 = -dist2
        
        If dist1 - dist2 = 0 Then
            later_jd = (later_jd + earlier_jd) / 2
            d = 0
        Else
            d = ((later_jd - earlier_jd) / (dist1 - dist2)) * dist1
        End If

        If Abs(dist1 - dist2) > 20 And n >= 2 Then
            'keep from looping needlessly AND
            'protect against case where dist1 = -dist2, which gives false aspect
            'example 21 March 2006 - Moon 120 Mars - there is no trine, but an opposition
            later_jd = 0
            Exit For
        End If

        If Abs(d) < e Then
            Exit For
        End If

        earlier_jd = later_jd

        If Abs(d) >= 1.001 Then
            'out of range - there is no aspect in this time frame (1 day)
            later_jd = 0
            Exit For
        Else
            later_jd = later_jd - d
        End If
    Next n

    If n > m Then
        Result_JD = 0
    Else
        Result_JD = later_jd
    End If
End Sub