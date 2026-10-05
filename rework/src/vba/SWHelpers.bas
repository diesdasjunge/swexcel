Attribute VB_Name = "SWHelpers"
Option Explicit

' SPDX-License-Identifier: AGPL-3.0-or-later
' Local Gregorian civil fields -> TT/UT1, with a numeric UTC offset (no DST inference).
Public Function SW_UTC_JD(ByVal year As Long, ByVal month As Long, ByVal day As Long, ByVal hour As Long, ByVal minute As Long, ByVal second As Double, Optional ByVal offsetHours As Double = 0#, Optional ByVal detail As Boolean = False) As Variant
    Dim utc As Variant, valid As Variant
    On Error GoTo Failed
    valid = SW_JULDAY(year, month, day)
    If IsError(valid) Then SWRaise "Invalid Gregorian date."
    If hour < 0 Or hour > 23 Or minute < 0 Or minute > 59 Or second < 0# Or second >= 61# Or Abs(offsetHours) > 24# Then SWRaise "Invalid clock or UTC offset."
    utc = SW_SWE_UTC_TIME_ZONE(year, month, day, hour, minute, second, offsetHours)
    If Not IsArray(utc) Then SWRaise "UTC offset conversion failed."
    SW_UTC_JD = SW_SWE_UTC_TO_JD(utc(1, 1), utc(1, 2), utc(1, 3), utc(1, 4), utc(1, 5), utc(1, 6), 1, , detail)
    Exit Function
Failed:
    SW_UTC_JD = CVErr(xlErrValue)
End Function

' Twelve cusps, or 36 Gauquelin sectors. Detail includes angles and daily speeds.
Public Function SW_HOUSES(ByVal julianDayUT As Double, ByVal latitude As Double, ByVal longitude As Double, Optional ByVal system As String = "P", Optional ByVal flags As Long = 0, Optional ByVal options As Variant, Optional ByVal detail As Boolean = False) As Variant
    Dim values As Variant, output() As Variant, n As Long, k As Long, row As Long
    On Error GoTo Failed
    If IsMissing(options) Then options = Empty
    values = SW_SWE_HOUSES_EX2(julianDayUT, flags, latitude, longitude, system, options, True)
    If detail Then
        SW_HOUSES = values
        Exit Function
    End If
    If values(2, 2) <> "OK" Then SWRaise CStr(values(3, 2))
    n = SWApiHouseCount(Asc(system)): ReDim output(1 To 1, 1 To n)
    For row = 9 To UBound(values, 1)
        If Left$(CStr(values(row, 1)), 6) = "cusps[" Then
            k = k + 1: output(1, k) = values(row, 2)
        End If
    Next row
    If k <> n Then SWRaise "Incomplete house result."
    SW_HOUSES = output
    Exit Function
Failed:
    SW_HOUSES = CVErr(xlErrNA)
End Function

' Date series: JD UT1 plus the six native coordinates/speeds; flags select units.
' Any non-OK calculation fails the entire compact table instead of hiding a gap.
Public Function SW_POSITIONS(ByVal startJD As Double, ByVal stepDays As Double, ByVal count As Long, ByVal body As Long, Optional ByVal flags As Long = 258, Optional ByVal options As Variant) As Variant
    Dim output() As Variant, values As Variant, row As Long, col As Long, jd As Double
    On Error GoTo Failed
    If count < 1 Or count > 10000 Then SWRaise "Row count must be 1..10000."
    If IsMissing(options) Then options = Empty
    ReDim output(1 To count, 1 To 7)
    For row = 1 To count
        jd = startJD + (row - 1) * stepDays
        values = SW_SWE_CALC_UT(jd, body, flags, options)
        If Not IsArray(values) Then SWRaise "A date-series calculation failed; use SW_SWE_CALC_UT with detail for that date."
        output(row, 1) = jd
        For col = 1 To 6
            output(row, col + 1) = values(1, col)
        Next col
    Next row
    SW_POSITIONS = output
    Exit Function
Failed:
    SW_POSITIONS = CVErr(xlErrNA)
End Function
