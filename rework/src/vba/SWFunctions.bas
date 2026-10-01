Attribute VB_Name = "SWFunctions"
Option Explicit

' SPDX-License-Identifier: AGPL-3.0-or-later
' Worksheet functions return values only. They never write cells or show UI.
Public Function SW_VERSION() As Variant
    On Error GoTo Failed
    SW_VERSION = SWEngineVersion()
    Exit Function
Failed:
    SW_VERSION = CVErr(xlErrValue)
End Function

Public Function SW_RUNTIME_STATUS() As String
    SW_RUNTIME_STATUS = SWRuntimeStatus()
End Function

Public Function SW_ENGINE_PATH() As Variant
    On Error GoTo Failed
    SW_ENGINE_PATH = SWEnginePath()
    Exit Function
Failed:
    SW_ENGINE_PATH = CVErr(xlErrValue)
End Function

Public Function SW_DATA_PATH() As Variant
    On Error GoTo Failed
    SW_DATA_PATH = SWDataPath()
    Exit Function
Failed:
    SW_DATA_PATH = CVErr(xlErrValue)
End Function

Public Function SW_BODY_NAME(ByVal body As Long) As Variant
    Dim bytes(0 To 255) As Byte
    Dim acquired As Boolean
#If Mac Then
#ElseIf Win64 Then
    Dim pointer As LongPtr
#End If
    On Error GoTo Failed
    If body < 0 Then SWRaise "Body ID must be nonnegative."
    SWBeginCalculation
    acquired = True
    SWApplyCalculationOptions 0, 0#, 0#, 0#
#If Mac Then
    SWRaise "Windows is required."
#ElseIf Win64 Then
    pointer = Native_swe_get_planet_name(body, bytes(0))
    If pointer <> VarPtr(bytes(0)) Then SWRaise "The engine returned an unexpected body-name pointer."
    SW_BODY_NAME = SWBufferText(bytes)
#Else
    SWRaise "64-bit Windows Excel is required."
#End If
    SWEndCalculation
    Exit Function
Failed:
    If acquired Then SWEndCalculation
    SW_BODY_NAME = CVErr(xlErrValue)
End Function

' Civil calendar to Julian day in the same time scale as hour.
' Passing UTC clock fields here does not perform the UTC -> UT1/TT conversion.
Public Function SW_JULDAY(ByVal year As Long, ByVal month As Long, ByVal day As Long, Optional ByVal hour As Double = 0#, Optional ByVal calendar As Long = 1) As Variant
    Dim result As Double, checkYear As Long, checkMonth As Long, checkDay As Long, checkHour As Double
    On Error GoTo Failed
    If calendar <> 0 And calendar <> 1 Then SWRaise "Calendar must be 0 (Julian) or 1 (Gregorian)."
    If month < 1 Or month > 12 Or day < 1 Or day > 31 Or hour < 0# Or hour >= 24# Then SWRaise "Invalid civil date or decimal hour."
    SWEnsureEngine
    result = Native_swe_julday(year, month, day, hour, calendar)
    Native_swe_revjul result, calendar, checkYear, checkMonth, checkDay, checkHour
    If checkYear <> year Or checkMonth <> month Or checkDay <> day Then SWRaise "This day does not exist in the selected calendar."
    SW_JULDAY = result
    Exit Function
Failed:
    SW_JULDAY = CVErr(xlErrValue)
End Function

Public Function SW_DEGNORM(ByVal value As Double) As Variant
    On Error GoTo Failed
    SWEnsureEngine
    SW_DEGNORM = Native_swe_degnorm(value)
    Exit Function
Failed:
    SW_DEGNORM = CVErr(xlErrValue)
End Function

' Six outputs: longitude/latitude/distance and their daily speeds by default.
' Flags can select equatorial, radians, or XYZ coordinates, matching swe_calc_ut.
Public Function SW_CALC_UT(ByVal julianDayUT As Double, ByVal body As Long, Optional ByVal flags As Long = 258, Optional ByVal siderealMode As Long = 0, Optional ByVal longitudeEast As Variant, Optional ByVal latitudeNorth As Variant, Optional ByVal altitudeMetres As Double = 0#) As Variant
    Dim result As SWPositionResult, output(1 To 1, 1 To 6) As Variant, column As Long
    result = SWCalculateUT(julianDayUT, body, flags, siderealMode, longitudeEast, latitudeNorth, altitudeMetres)
    If result.Status <> "OK" Then
        SW_CALC_UT = CVErr(xlErrNA)
        Exit Function
    End If
    For column = 1 To 6
        output(1, column) = result.Values(column - 1)
    Next column
    SW_CALC_UT = output
End Function

Public Function SW_LONGITUDE(ByVal julianDayUT As Double, ByVal body As Long, Optional ByVal flags As Long = 258, Optional ByVal siderealMode As Long = 0, Optional ByVal longitudeEast As Variant, Optional ByVal latitudeNorth As Variant, Optional ByVal altitudeMetres As Double = 0#) As Variant
    Dim result As SWPositionResult
    If (flags And (SW_FLAG_EQUATORIAL Or SW_FLAG_XYZ Or SW_FLAG_RADIANS)) <> 0 Then
        SW_LONGITUDE = CVErr(xlErrValue)
        Exit Function
    End If
    result = SWCalculateUT(julianDayUT, body, flags, siderealMode, longitudeEast, latitudeNorth, altitudeMetres)
    If result.Status <> "OK" Then
        SW_LONGITUDE = CVErr(xlErrNA)
    Else
        SW_LONGITUDE = result.Values(0)
    End If
End Function

Public Function SW_POSITION_DETAIL(ByVal julianDayUT As Double, ByVal body As Long, Optional ByVal flags As Long = 258, Optional ByVal siderealMode As Long = 0, Optional ByVal longitudeEast As Variant, Optional ByVal latitudeNorth As Variant, Optional ByVal altitudeMetres As Double = 0#) As Variant
    Dim result As SWPositionResult, output(1 To 22, 1 To 3) As Variant
    Dim labels As Variant, units As Variant, values As Variant, row As Long, angleUnit As String
    result = SWCalculateUT(julianDayUT, body, flags, siderealMode, longitudeEast, latitudeNorth, altitudeMetres)
    angleUnit = "degree"
    If (flags And SW_FLAG_RADIANS) <> 0 Then angleUnit = "radian"
    labels = Array("Status", "Julian day UT1", "Body", "Input flags", "Actual flags", "Requested ephemeris", "Actual ephemeris", "Longitude", "Latitude", "Distance", "Longitude speed", "Latitude speed", "Distance speed", "Warning", "Engine path", "Engine version", "Data path", "Sidereal mode", "Observer longitude", "Observer latitude", "Observer altitude")
    units = Array("", "day", "Swiss ID", "bit mask", "bit mask", "", "", angleUnit, angleUnit, "AU", angleUnit & "/day", angleUnit & "/day", "AU/day", "", "", "", "", "Swiss ID", "degree east", "degree north", "metre")
    If (flags And SW_FLAG_EQUATORIAL) <> 0 Then
        labels(7) = "Right ascension": labels(8) = "Declination"
        labels(10) = "Right ascension speed": labels(11) = "Declination speed"
    End If
    If (flags And SW_FLAG_XYZ) <> 0 Then
        labels(7) = "X": labels(8) = "Y": labels(9) = "Z"
        labels(10) = "X speed": labels(11) = "Y speed": labels(12) = "Z speed"
        units(7) = "AU": units(8) = "AU": units(9) = "AU"
        units(10) = "AU/day": units(11) = "AU/day": units(12) = "AU/day"
    End If
    values = Array(result.Status, result.JulianDayUT, result.Body, result.InputFlags, result.ActualFlags, result.RequestedEphemeris, result.ActualEphemeris, result.Values(0), result.Values(1), result.Values(2), result.Values(3), result.Values(4), result.Values(5), result.Warning, result.EnginePath, result.EngineVersion, result.DataPath, result.SiderealMode, result.LongitudeEast, result.LatitudeNorth, result.AltitudeMetres)
    If result.Status = "ERROR" Then
        ' Unwritten native outputs are not valid numerical zeroes.
        For row = 7 To 12
            values(row) = CVErr(xlErrNA)
        Next row
        values(4) = CVErr(xlErrNA)
        values(6) = "Unavailable"
    End If
    output(1, 1) = "Field": output(1, 2) = "Value": output(1, 3) = "Unit"
    For row = 1 To 21
        output(row + 1, 1) = labels(row - 1)
        output(row + 1, 2) = values(row - 1)
        output(row + 1, 3) = units(row - 1)
    Next row
    SW_POSITION_DETAIL = output
End Function
