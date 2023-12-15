Sub FigureHouseInterceptions()
'get how many signs are intercepted in each house
    For x = 1 To 12: InterceptedHouse%(x) = 0: Next x
    For x = 1 To 11: TEMP7 = Int(HC3(x) / 30) + 1
    TEMP8 = Int(HC3(x + 1) / 30) + 1
    If TEMP8 - TEMP7 < 0 Then TEMP8 = TEMP8 + 12
    InterceptedHouse%(x) = TEMP8 - TEMP7: Next x
End Sub