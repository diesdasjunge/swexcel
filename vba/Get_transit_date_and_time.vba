Public Function Get_transit_date_and_time(sm%, sd%, sy%, em%, ed%, ey%, p_num%, start_pos As Single, angle As Single) As String
    Dim end_pos As Single, separation As Double, starting_JD As Double, ending_JD As Double, x As Double
    
    StartJD = GetJulianDay(sm%, sd%, sy%)
    EndJD = GetJulianDay(em%, ed%, ey%)
    
    end_pos = start_pos + angle
    If end_pos >= 360 Then end_pos = end_pos - 360

    For x = StartJD To EndJD
        Call Get_Initial_Planet_Positions(x, x + 1, p_num%, end_pos)
        
        separation = Get_separation_for_this_aspect         'aspect is always 0 - always looking for conjunction

        If Abs(separation) < 1.001 Then
            Call SecantMethod(x, x + 1, 0.00007, 100, p_num%, end_pos)

            rnd_off = Result_JD
            If rnd_off >= x And rnd_off <= x + 1 Then
                Call ConvertJDtoDateandTime(Result_JD)
                
                Get_transit_date_and_time = JDDate$ & " " & JDTime$
                
                Exit Function
            End If
        End If
    Next x

    Get_transit_date_and_time = ""
End Function