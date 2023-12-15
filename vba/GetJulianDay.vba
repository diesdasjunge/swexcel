Function GetJulianDay(XMonth As Integer, XDay As Integer, XYear As Integer)
    Dim XIM As Double, XJD As Double

'find Julian Day for midnight of day in question
    XIM = 12 * (CDbl(XYear) + 4800) + XMonth - 3

    XJD = (2 * (XIM - Int(XIM / 12) * 12) + 7 + 365 * XIM) / 12
    XJD = Int(XJD) + XDay + Int(XIM / 48) - 32083

    If XJD > 2299171 Then
        XJD = (XJD + Int(XIM / 4800) - Int(XIM / 1200) + 38)
    End If

    GetJulianDay = XJD - 0.5
End Function