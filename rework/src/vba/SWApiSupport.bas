Attribute VB_Name = "SWApiSupport"
Option Explicit
Option Private Module

' SPDX-License-Identifier: AGPL-3.0-or-later
Public Type SWApiResult
    FunctionName As String
    EnginePath As String
    Status As String
    Warning As String
    NativeReturn As Variant
    SunDeclination As Variant
    Fields As Collection
    Values As Collection
    Units As Collection
    Primary As Collection
End Type

Private Declare PtrSafe Function ReadProcessMemory Lib "kernel32" (ByVal process As LongPtr, ByVal source As LongPtr, ByRef destination As Byte, ByVal count As LongPtr, ByRef copied As LongPtr) As Long

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByRef bytes As Byte, ByVal size As Long, ByVal destination As LongPtr, ByVal capacity As Long) As Long

Public Function SWApiOutputText(ByRef bytes() As Byte, ByVal functionName As String) As String
    Dim count As Long, text As String, n As Long
    If functionName <> "swe_cs2degstr" Then
        SWApiOutputText = SWBufferText(bytes)
        Exit Function
    End If
    ' The pinned degree formatter embeds a UTF-8 degree sign, unlike ANSI paths.
    For count = 0 To UBound(bytes)
        If bytes(count) = 0 Then Exit For
    Next count
    text = String$(count, vbNullChar)
    n = MultiByteToWideChar(65001, 8, bytes(0), count, StrPtr(text), count)
    If n = 0 And count > 0 Then SWRaise "Invalid UTF-8 formatter output."
    SWApiOutputText = Left$(text, n)
End Function

Public Function SWApiNumber(ByVal value As Variant) As Double
    If IsObject(value) Then value = value.Value2
    If IsError(value) Or IsEmpty(value) Or IsArray(value) Then SWRaise "A numeric scalar is required."
    If Not IsNumeric(value) Then SWRaise "A numeric scalar is required."
    SWApiNumber = CDbl(value)
    If Abs(SWApiNumber) > 1E+100 Then SWRaise "The input exceeds the supported numerical range."
End Function

Public Function SWApiInteger(ByVal value As Variant) As Long
    Dim number As Double
    number = SWApiNumber(value)
    If number <> Fix(number) Or number < -2147483648# Or number > 2147483647# Then SWRaise "A signed 32-bit integer is required."
    SWApiInteger = CLng(number)
End Function

Public Function SWApiCharacter(ByVal value As Variant) As Long
    Dim s As String
    If IsObject(value) Then value = value.Value2
    If IsError(value) Or IsArray(value) Then SWRaise "A single ASCII character is required."
    s = CStr(value)
    If Len(s) <> 1 Then SWRaise "A single ASCII character is required."
    SWApiCharacter = AscW(s)
    If SWApiCharacter < 1 Or SWApiCharacter > 127 Then SWRaise "A single ASCII character is required."
End Function

Public Sub SWApiPutText(ByRef target() As Byte, ByVal text As String)
    Dim bytes() As Byte, k As Long
    bytes = SWAnsiZ(text, UBound(target) + 1)
    For k = 0 To UBound(bytes)
        target(k) = bytes(k)
    Next k
End Sub

Public Sub SWApiPutVector(ByRef target() As Double, ByVal source As Variant, ByVal count As Long)
    Dim item As Variant, k As Long
    If IsObject(source) Then source = source.Value2
    If Not IsArray(source) Then SWRaise "Pass a row/column range or a VBA array."
    For Each item In source
        If k >= count Then SWRaise "Too many vector elements."
        target(k) = SWApiNumber(item)
        k = k + 1
    Next item
    If k <> count Then SWRaise "Incorrect vector length; expected " & CStr(count) & "."
End Sub

' Copy only readable bytes, stopping at NUL; never free an engine-owned pointer.
Public Function SWApiBorrowedText(ByVal pointer As LongPtr, ByVal capacity As Long) As String
    Dim bytes() As Byte, copied As LongPtr, k As Long
    If pointer = 0 Then SWRaise "No native string is available for this input/state."
    ReDim bytes(0 To capacity - 1)
    For k = 0 To capacity - 1
        If ReadProcessMemory(-1, pointer + k, bytes(k), 1, copied) = 0 Or copied <> 1 Then SWRaise "Unreadable native string pointer."
        If bytes(k) = 0 Then
            SWApiBorrowedText = SWBufferText(bytes)
            Exit Function
        End If
    Next k
    SWRaise "Native string exceeds its reviewed capacity."
End Function

Public Sub SWApiAdd(ByRef result As SWApiResult, ByVal label As String, ByVal value As Variant, ByVal unit As String, Optional ByVal primary As Boolean = True)
    result.Fields.Add label
    result.Values.Add value
    result.Units.Add unit
    result.Primary.Add primary
End Sub

Public Sub SWApiValidateObserver(ByRef vector() As Double)
    If Abs(vector(0)) > 180# Or Abs(vector(1)) > 90# Then SWRaise "Observer longitude/latitude are outside their degree ranges."
    If vector(2) < -500# Or vector(2) > 25000# Then SWRaise "Observer altitude must be within -500..25000 metres."
End Sub

Public Sub SWApiValidateHouse(ByVal code As Long)
    If InStr(1, "ABCDEFGHIiJKLMNOPQRSTUVWXY", Chr$(code), vbBinaryCompare) = 0 Then SWRaise "Unknown house system code."
End Sub

Public Function SWApiHouseCount(ByVal code As Long) As Long
    SWApiHouseCount = 12
    If code = Asc("G") Then SWApiHouseCount = 36
End Function

Public Sub SWApiValidateHeliacal(ByRef atmosphere() As Double, ByRef observer() As Double)
    If atmosphere(0) < 0# Or atmosphere(1) <= -273# Or atmosphere(2) < 0# Or atmosphere(2) > 100# Or atmosphere(3) < 0# Then SWRaise "Invalid atmospheric parameters."
    If observer(0) < 0# Or observer(0) > 120# Or observer(1) < 0# Then SWRaise "Invalid observer age/visual acuity."
    If observer(2) <> 0# And observer(2) <> 1# Then SWRaise "Binocular flag must be 0 or 1."
    If observer(3) < 0# Or observer(4) < 0# Or observer(5) < 0# Or observer(5) > 1# Then SWRaise "Invalid optical observer parameters."
End Sub

Public Sub SWApiValidateModels(ByVal text As String)
    Dim fields As Variant, k As Long, n As Long, maxima As Variant
    If Left$(text, 2) = "SE" Then
        If Len(text) > 20 Or InStr(text, "+") > 0 Then SWRaise "Use an engine version up to 20 characters without '+'."
        Exit Sub
    End If
    fields = Split(text, ",")
    If UBound(fields) <> 7 Then SWRaise "Astronomical models require exactly eight comma-separated model IDs."
    maxima = Array(5, 11, 11, 5, 3, 2, 3, 4)
    For k = 0 To 7
        n = SWApiInteger(fields(k))
        If n < 0 Or n > maxima(k) Then SWRaise "Astronomical model ID exceeds the pinned engine's range."
    Next k
End Sub

Public Sub SWApiApplyOptions(ByVal options As Variant, ByRef result As SWApiResult)
    Dim sid As Long, epoch As Double, ayan As Double, lon As Double, lat As Double, alt As Double
    Dim dt As Double, tidal As Double, lapse As Double, interpolate As Long
    Dim models As String, ephe As String, jpl As String, bytes(0 To 255) As Byte
    Dim data As Variant, row As Long, key As String, value As Variant, used As Object
    Dim pathBytes() As Byte, modelBytes() As Byte, jplBytes() As Byte
    dt = SW_DELTAT_AUTOMATIC: tidal = SW_TIDAL_AUTOMATIC: lapse = 0.0065
    models = "0,0,0,0,0,0,0,0": ephe = SWDataPath(): jpl = "de431.eph"
    If IsObject(options) Then options = options.Value2
    If Not IsEmpty(options) Then
        If Not IsArray(options) Then SWRaise "Options must be a two-column key/value array."
        data = options
        If UBound(data, 2) - LBound(data, 2) <> 1 Then SWRaise "Options must have two columns."
        Set used = CreateObject("Scripting.Dictionary")
        For row = LBound(data, 1) To UBound(data, 1)
            key = LCase$(Trim$(CStr(data(row, LBound(data, 2)))))
            value = data(row, LBound(data, 2) + 1)
            If Len(key) > 0 And key <> "option" Then
                If used.Exists(key) Then SWRaise "Duplicate option: " & key
                used.Add key, True
                Select Case key
                    Case "sun_declination": result.SunDeclination = SWApiNumber(value)
                    Case "sidereal_mode": sid = SWApiInteger(value)
                    Case "sidereal_epoch": epoch = SWApiNumber(value)
                    Case "ayanamsa_epoch": ayan = SWApiNumber(value)
                    Case "longitude": lon = SWApiNumber(value)
                    Case "latitude": lat = SWApiNumber(value)
                    Case "altitude": alt = SWApiNumber(value)
                    Case "delta_t": dt = SWApiNumber(value)
                    Case "tidal_acceleration": tidal = SWApiNumber(value)
                    Case "lapse_rate": lapse = SWApiNumber(value)
                    Case "interpolate_nutation": interpolate = SWApiInteger(value)
                    Case "astro_models": models = CStr(value)
                    Case "ephe_path": ephe = CStr(value)
                    Case "jpl_file": jpl = CStr(value)
                    Case Else: SWRaise "Unknown option: " & key
                End Select
            End If
        Next row
    End If
    If sid < 0 Or ((sid And 255) > 46 And (sid And 255) <> 255) Then SWRaise "Unknown sidereal mode."
    If Abs(lon) > 180# Or Abs(lat) > 90# Or alt < -500# Or alt > 25000# Then SWRaise "Invalid observer options."
    If lapse <= 0# Or lapse > 0.1 Then SWRaise "Lapse rate must be within 0..0.1 K/m."
    If interpolate <> 0 And interpolate <> 1 Then SWRaise "Nutation interpolation must be 0 or 1."
    If Len(ephe) = 0 Or InStr(ephe, ";") > 0 Then SWRaise "Use one absolute ephemeris directory."
    If Mid$(ephe, 2, 2) <> ":\" And Left$(ephe, 2) <> "\\" Then SWRaise "Ephemeris path must be absolute."
    If Len(Dir$(ephe, vbDirectory)) = 0 Then SWRaise "Ephemeris directory does not exist."
    If Len(jpl) = 0 Or InStr(jpl, "\") Or InStr(jpl, "/") Or InStr(jpl, ":") Then SWRaise "JPL filename must be relative to the ephemeris directory."
    SWApiValidateModels models
    pathBytes = SWAnsiZ(ephe, SW_EPHE_PATH_BYTES)
    modelBytes = SWAnsiZ(models, 256)
    jplBytes = SWAnsiZ(jpl, 256)
    Native_swe_set_ephe_path pathBytes(0)
    Native_swe_set_jpl_file jplBytes(0)
    Native_swe_set_astro_models modelBytes(0), SW_DEFAULT_FLAGS
    Native_swe_set_delta_t_userdef dt
    Native_swe_set_tid_acc tidal
    Native_swe_set_lapse_rate lapse
    Native_swe_set_interpolate_nut interpolate
    Native_swe_set_sid_mode sid, epoch, ayan
    Native_swe_set_topo lon, lat, alt
    SWApiAdd result, "Data path", ephe, "directory", False
    SWApiAdd result, "Options sidereal", CStr(sid) & "; epoch=" & CStr(epoch) & "; ayanamsa=" & CStr(ayan), "ID; JD TT; degrees", False
    SWApiAdd result, "Options observer", CStr(lon) & "; " & CStr(lat) & "; " & CStr(alt), "degrees east; north; metres", False
    SWApiAdd result, "Options models", models, "model IDs", False
    SWApiAdd result, "Options deltaT/tidal/lapse", CStr(dt) & "; " & CStr(tidal) & "; " & CStr(lapse), "days; arcsec/century^2; K/m", False
    SWApiAdd result, "Options JPL/interpolation", jpl & "; " & CStr(interpolate), "filename; boolean", False
End Sub

Public Sub SWApiCheckReturn(ByRef result As SWApiResult, ByVal kind As String, ByVal startDate As Double)
    Select Case kind
        Case "status", "flags", "event", "rise", "visibility"
            If CDbl(result.NativeReturn) < 0# Then result.Status = "ERROR"
            If kind = "event" And CDbl(result.NativeReturn) = 0# Then result.Status = "NO_EVENT"
            If kind = "rise" And CDbl(result.NativeReturn) = -2# Then result.Status = "NO_EVENT"
            If kind = "visibility" And CDbl(result.NativeReturn) = -2# Then result.Status = "BELOW_HORIZON"
        Case "crossing"
            If CDbl(result.NativeReturn) < startDate Then result.Status = "ERROR"
        Case "house_position"
            If CDbl(result.NativeReturn) = 0# Then result.Status = "ERROR"
    End Select
    If Len(result.Warning) > 0 And result.Status = "OK" Then result.Status = "WARNING"
    If result.Status <> "OK" And Len(result.Warning) = 0 Then result.Warning = "Native result: " & result.Status
End Sub

Public Sub SWApiCheckFlags(ByRef result As SWApiResult, ByVal requested As Long)
    Dim actualModel As Long, requestedModel As Long
    If CDbl(result.NativeReturn) < 0# Then Exit Sub
    actualModel = CLng(result.NativeReturn) And 7
    requestedModel = requested And 7
    If requestedModel = 0 Then requestedModel = 2
    SWApiAdd result, "Requested ephemeris", requestedModel, "1 JPL; 2 Swiss; 4 Moshier", False
    SWApiAdd result, "Actual ephemeris", actualModel, "1 JPL; 2 Swiss; 4 Moshier", False
    If actualModel <> requestedModel Then
        result.Status = "FALLBACK"
        result.Warning = result.Warning & " Requested ephemeris " & CStr(requestedModel) & "; actual " & CStr(actualModel) & "."
    End If
End Sub

Public Function SWApiReturnUnit(ByVal name As String) As String
    Select Case name
        Case "swe_deltat", "swe_deltat_ex": SWApiReturnUnit = "days (TT minus UT1)"
        Case "swe_julday", "swe_solcross", "swe_mooncross", "swe_mooncross_node": SWApiReturnUnit = "Julian day in input time scale / TT"
        Case "swe_solcross_ut", "swe_mooncross_ut", "swe_mooncross_node_ut": SWApiReturnUnit = "Julian day UT1"
        Case "swe_radnorm", "swe_rad_midp", "swe_difrad2n": SWApiReturnUnit = "radians"
        Case "swe_csnorm", "swe_csroundsec", "swe_difcsn", "swe_difcs2n": SWApiReturnUnit = "centiseconds"
        Case "swe_d2l": SWApiReturnUnit = "rounded integer"
        Case "swe_day_of_week": SWApiReturnUnit = "0 Monday .. 6 Sunday"
        Case "swe_sidtime", "swe_sidtime0": SWApiReturnUnit = "sidereal hours"
        Case "swe_get_tid_acc": SWApiReturnUnit = "arcseconds/century squared"
        Case "swe_house_pos": SWApiReturnUnit = "fractional house/sector number"
        Case Else: SWApiReturnUnit = "degrees"
    End Select
End Function

Private Function RenderResult(ByRef result As SWApiResult, ByVal detail As Boolean) As Variant
    Dim output() As Variant, k As Long, offset As Long, n As Long, j As Long
    If Not detail Then
        If result.Status <> "OK" Then
            RenderResult = CVErr(xlErrNA)
        Else
            For k = 1 To result.Values.Count
                If result.Primary(k) Then n = n + 1
            Next k
            If n = 0 Then
                RenderResult = result.NativeReturn
            Else
                ReDim output(1 To 1, 1 To n)
                For k = 1 To result.Values.Count
                    If result.Primary(k) Then
                        j = j + 1: output(1, j) = result.Values(k)
                    End If
                Next k
                If n = 1 Then
                    RenderResult = output(1, 1)
                Else
                    RenderResult = output
                End If
            End If
        End If
        Exit Function
    End If
    offset = 8
    ReDim output(1 To result.Values.Count + offset, 1 To 3)
    For k = 1 To UBound(output, 1)
        For j = 1 To 3
            output(k, j) = ""
        Next j
    Next k
    output(1, 1) = "Field": output(1, 2) = "Value": output(1, 3) = "Unit"
    output(2, 1) = "Status": output(2, 2) = result.Status
    output(3, 1) = "Warning": output(3, 2) = result.Warning
    output(4, 1) = "Function": output(4, 2) = result.FunctionName
    output(5, 1) = "Native return": output(5, 2) = result.NativeReturn
    output(6, 1) = "Engine version": output(6, 2) = SW_ENGINE_VERSION
    output(7, 1) = "Engine path": output(7, 2) = result.EnginePath
    output(8, 1) = "Result convention": output(8, 2) = "Native indexed outputs; flags and time scales follow the reference"
    For k = 1 To result.Values.Count
        output(k + offset, 1) = result.Fields(k)
        If result.Status = "ERROR" Or result.Status = "NO_EVENT" Or result.Status = "BELOW_HORIZON" Then
            output(k + offset, 2) = CVErr(xlErrNA)
        Else
            output(k + offset, 2) = result.Values(k)
        End If
        output(k + offset, 3) = result.Units(k)
    Next k
    RenderResult = output
End Function

Public Function SWApiExecute(ByVal functionName As String, ByVal values As Variant, ByVal options As Variant, ByVal detail As Boolean) As Variant
    Dim result As SWApiResult, acquired As Boolean
    Set result.Fields = New Collection: Set result.Values = New Collection: Set result.Units = New Collection: Set result.Primary = New Collection
    result.FunctionName = functionName: result.Status = "OK"
    On Error GoTo Failed
    SWBeginCalculation
    acquired = True
    result.EnginePath = SWEnginePath()
    If Left$(functionName, 8) <> "swe_set_" And functionName <> "swe_close" And functionName <> "swe_get_current_file_data" Then SWApiApplyOptions options, result
    SWApiDispatch functionName, values, result
    SWEndCalculation
    acquired = False
    SWApiExecute = RenderResult(result, detail)
    Exit Function
Failed:
    result.Status = "ERROR": result.Warning = Err.Description
    If acquired Then SWEndCalculation
    SWApiExecute = RenderResult(result, detail)
End Function
