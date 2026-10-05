Attribute VB_Name = "SWCalculation"
Option Explicit
Option Private Module

' SPDX-License-Identifier: AGPL-3.0-or-later
' Result metadata belongs to this call, never to a global last-error slot.
Public Type SWPositionResult
    Values(0 To 5) As Double
    JulianDayUT As Double
    Body As Long
    InputFlags As Long
    ActualFlags As Long
    RequestedEphemeris As String
    ActualEphemeris As String
    Warning As String
    Status As String
    EnginePath As String
    EngineVersion As String
    DataPath As String
    SiderealMode As Long
    LongitudeEast As Double
    LatitudeNorth As Double
    AltitudeMetres As Double
End Type

Public Function SWEphemerisName(ByVal flags As Long) As String
    Select Case flags And 7
        Case SW_FLAG_JPL: SWEphemerisName = "JPL"
        Case SW_FLAG_SWISS: SWEphemerisName = "Swiss"
        Case SW_FLAG_MOSHIER: SWEphemerisName = "Moshier"
        Case 0: SWEphemerisName = "Swiss"
        Case Else: SWEphemerisName = "Invalid ephemeris flags"
    End Select
End Function

Public Function SWCalculateUT(ByVal julianDayUT As Double, ByVal body As Long, ByVal flags As Long, ByVal siderealMode As Long, ByVal longitudeEast As Variant, ByVal latitudeNorth As Variant, ByVal altitudeMetres As Double) As SWPositionResult
    Dim result As SWPositionResult, errorBytes(0 To 255) As Byte
    Dim acquired As Boolean, message As String, model As Long
    On Error GoTo Failed
    result.JulianDayUT = julianDayUT
    result.Body = body
    result.InputFlags = flags
    result.SiderealMode = siderealMode
    result.AltitudeMetres = altitudeMetres
    If body < 0 Then SWRaise "This convenience helper accepts body IDs 0 and above; use SW_SWE_CALC_UT for ecliptic/nutation output."
    If flags < 0 Then SWRaise "Flags must be nonnegative."
    model = flags And 7
    If model <> 0 And model <> 1 And model <> 2 And model <> 4 Then SWRaise "Select one ephemeris model."
    If (flags And SW_FLAG_JPL) <> 0 Then SWRaise "Use SW_SWE_CALC_UT with the jpl_file option for external raw JPL files."
    If siderealMode < 0 Or siderealMode > 46 Then SWRaise "Use a built-in mode 0..46 here; SW_SWE_CALC_UT accepts custom epochs through options."
    If (flags And SW_FLAG_TOPOCENTRIC) <> 0 Then
        If IsMissing(longitudeEast) Or IsMissing(latitudeNorth) Then SWRaise "Topocentric calculations require observer longitude and latitude."
        If IsError(longitudeEast) Or IsError(latitudeNorth) Then SWRaise "Observer coordinates contain an Excel error."
        If IsEmpty(longitudeEast) Or IsEmpty(latitudeNorth) Then SWRaise "Observer coordinates are empty."
        If Not IsNumeric(longitudeEast) Or Not IsNumeric(latitudeNorth) Then SWRaise "Observer coordinates must be numbers."
        result.LongitudeEast = CDbl(longitudeEast)
        result.LatitudeNorth = CDbl(latitudeNorth)
        If Abs(result.LongitudeEast) > 180# Or Abs(result.LatitudeNorth) > 90# Then SWRaise "Observer longitude must be within -180..180 and latitude within -90..90 degrees."
    Else
        result.AltitudeMetres = 0#
    End If
    result.RequestedEphemeris = SWEphemerisName(flags)
    SWBeginCalculation
    acquired = True
    SWApplyCalculationOptions siderealMode, result.LongitudeEast, result.LatitudeNorth, result.AltitudeMetres
    result.ActualFlags = Native_swe_calc_ut(julianDayUT, body, flags, result.Values(0), errorBytes(0))
    result.Warning = SWBufferText(errorBytes)
    result.EnginePath = SWEnginePath()
    result.EngineVersion = SWEngineVersion()
    result.DataPath = SWDataPath()
    If result.ActualFlags < 0 Then
        If Len(result.Warning) = 0 Then result.Warning = "Swiss Ephemeris reported a calculation error."
        SWRaise result.Warning
    End If
    result.ActualEphemeris = SWEphemerisName(result.ActualFlags)
    If result.ActualEphemeris <> result.RequestedEphemeris Then
        result.Status = "FALLBACK"
        If Len(result.Warning) > 0 Then result.Warning = result.Warning & " "
        result.Warning = result.Warning & "Requested " & result.RequestedEphemeris & "; calculated with " & result.ActualEphemeris & ". Numeric helpers return #N/A for this fallback."
    Else
        result.Status = "OK"
    End If
    SWEndCalculation
    SWCalculateUT = result
    Exit Function
Failed:
    message = Err.Description
    If acquired Then SWEndCalculation
    result.Status = "ERROR"
    result.Warning = message
    SWCalculateUT = result
End Function
