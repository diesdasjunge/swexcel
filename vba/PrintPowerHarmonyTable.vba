Sub PrintPowerHarmonyTable()
'print out planet's power and harmony in graph-type form
    Call LineFeed: Call LineFeed
    Buffer = Space(25) + "Power is listed in upper right section.": Call PrintBuffer
    Buffer = Space(25) + "Harmony is listed in lower left section.": Call PrintBuffer
    Call LineFeed
    Buffer = Space(10) + "Sun    Mer    Ven    Mar    Jup    Sat    Ura    Nep    Plu    Moo    Asc    MC"
    Call PrintBuffer
    For y = 1 To 15
    If y < 11 Or y > 13 Then
        Buffer = Space(2) + Left$(PN$(y), 3) + "  "
    Else
        GoTo ppht3
    End If
    For x = 1 To 15: If x = 11 Or x = 12 Or x = 13 Then GoTo ppht2
    If y = x Then Buffer = Buffer + "  **** ": GoTo ppht2
    If Val(Align4x(Format$(Cosmo!(y, x), "#0.00"))) = 0 Then
        Buffer = Buffer + "  ---- "
    Else
        Buffer = Buffer + Align4x(Format$(Cosmo!(y, x), "#0.00")) + " "
    End If
ppht2:  Next x: Call PrintBuffer
ppht3:  Next y
End Sub