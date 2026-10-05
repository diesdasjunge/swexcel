Attribute VB_Name = "SWIntegration"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function IsWow64Process2 Lib "kernel32" (ByVal processHandle As LongPtr, ByRef processMachine As Integer, ByRef nativeMachine As Integer) As Long
#End If

Private mPassed As Long
Private mFailed As Long
Private mNextRow As Long
Private mReport As Worksheet

' Explicit Windows smoke test only. Workbook construction is not acceptance.
' Do not call from Workbook_Open or a worksheet formula.
Public Function SW_RunIntegrationChecks() As String
    Dim previousUpdating As Boolean
    Dim actual As Variant, vector As Variant, detail As Variant
    Dim expectedVersion As String, expectedEngine As String
    Dim found As Boolean, value As Variant
    Dim rowLower As Long, rowUpper As Long, colLower As Long, colUpper As Long
    Dim col As Long, allDoubles As Boolean, sheetVector As Variant
    Dim detailRange As Object, spilledDetail As Variant
    Dim invalidResult As Variant, recoveredResult As Variant
    Dim fatalNumber As Long, fatalDescription As String

    previousUpdating = Application.ScreenUpdating
    On Error GoTo Fatal
    Application.ScreenUpdating = False
    Set mReport = ThisWorkbook.Worksheets("Integration")
    mReport.Range("A5:E23").ClearContents
    mReport.Range("G4:H16").ClearContents
    ' Preserve diagnostic strings; Excel must not reinterpret decimal commas
    ' or timestamps using a different regional format from VBA's CStr/Format.
    mReport.Range("A5:E23").NumberFormat = "@"
    mReport.Range("G4:H16").NumberFormat = "@"
    mPassed = 0
    mFailed = 0
    mNextRow = 5
    WriteEnvironment

#If Win64 Then
    RecordCheck "Host architecture", "64-bit Excel / VBA", "Win64", True, ProcessArchitecture()
#Else
    RecordCheck "Host architecture", "64-bit Excel / VBA", "Not Win64", False, "This checkpoint requires Windows 64-bit Excel."
    GoTo Finish
#End If

    expectedVersion = Trim$(CStr(ThisWorkbook.Worksheets("Setup").Range("B8").Value2))
    expectedEngine = ThisWorkbook.Path & Application.PathSeparator & CStr(ThisWorkbook.Worksheets("Setup").Range("B9").Value2)
    RecordCheck "Expected engine metadata", "Resolved runtime version", expectedVersion, Len(expectedVersion) > 0 And InStr(expectedVersion, "@") = 0, "Read from package manifest during build."

    actual = SW_JULDAY(2000, 1, 1, 12)
    RecordCheck "SW_JULDAY", "2451545", SafeText(actual), NearNumber(actual, 2451545#, 0.000000001), "Calendar date/time to J2000 UT; direct wrapper call."
    actual = SW_DEGNORM(-1)
    RecordCheck "SW_DEGNORM", "359", SafeText(actual), NearNumber(actual, 359#, 0.000000000001), "Normalization scalar."
    actual = SW_VERSION()
    RecordCheck "SW_VERSION", expectedVersion, SafeText(actual), TextMatches(actual, expectedVersion), "Native pointer-to-string return."
    actual = SW_BODY_NAME(0)
    RecordCheck "SW_BODY_NAME", "Sun", SafeText(actual), TextMatches(actual, "Sun"), "Native output buffer plus pointer return."
    actual = SWEnginePath()
    RecordCheck "Verified module path", NormalizePath(expectedEngine), SafeText(actual), NormalizePath(SafeText(actual)) = NormalizePath(expectedEngine), "Runtime must report the actual loaded module path."

    vector = SW_CALC_UT(2451545#, 0)
    allDoubles = ArrayBounds2D(vector, rowLower, rowUpper, colLower, colUpper)
    If allDoubles Then
        allDoubles = rowUpper - rowLower = 0 And colUpper - colLower = 5
        If allDoubles Then
            For col = colLower To colUpper
                If VarType(vector(rowLower, col)) <> vbDouble Then allDoubles = False
            Next col
        End If
    End If
    RecordCheck "SW_CALC_UT return", "1 x 6 Double array", DescribeArray(vector), allDoubles, "Longitude, latitude, distance and three speeds."
    actual = SW_LONGITUDE(2451545#, 0)
    If allDoubles Then
        RecordCheck "SW_LONGITUDE parity", SafeText(vector(rowLower, colLower)), SafeText(actual), NearNumber(actual, CDbl(vector(rowLower, colLower)), 0.000000001), "Same inputs/options as the first CALC_UT result."
    Else
        RecordCheck "SW_LONGITUDE parity", "Valid CALC_UT vector", SafeText(actual), False, "CALC_UT failed; scalar parity cannot be established."
    End If
    RecordCheck "Sun J2000 broad reference", "280.369 degrees +/- 0.01", SafeText(actual), NearNumber(actual, 280.369, 0.01), "Independent broad README reference; not exact pinned-data numerical acceptance."

    invalidResult = SW_JULDAY(2000, 2, 30, 12)
    RecordCheck "Invalid calendar date", "Excel error", SafeText(invalidResult), IsError(invalidResult), "February 30 must fail rather than silently roll into March."
    ' Body 1000 lies between upstream SE_FICT_MAX (999) and SE_PLMOON_OFFSET (9000).
    ' It reaches the native invalid-body path after the calculation guard is acquired.
    invalidResult = SW_CALC_UT(2451545#, 1000)
    recoveredResult = SW_LONGITUDE(2451545#, 0)
    RecordCheck "Native failure then recovery", "Error for body 1000; valid Sun calculation", SafeText(invalidResult) & "; then " & SafeText(recoveredResult), IsError(invalidResult) And NearNumber(recoveredResult, 280.369, 0.01), "Valid calculation immediately after native failure confirms calculation guard release."

    detail = SW_POSITION_DETAIL(2451545#, 0)
    CheckDetail detail, "Direct detail", actual, expectedEngine

    Application.CalculateFullRebuild
    RecordCheck "Formula2 scalar cells", "JULDAY 2451545; DEGNORM 359; version; Sun", "See B27:B30", NearNumber(ThisWorkbook.Names("SW_CHECK_JULDAY").RefersToRange.Value2, 2451545#, 0.000000001) And NearNumber(ThisWorkbook.Names("SW_CHECK_DEGNORM").RefersToRange.Value2, 359#, 0.000000000001) And TextMatches(ThisWorkbook.Names("SW_CHECK_VERSION").RefersToRange.Value2, expectedVersion) And TextMatches(ThisWorkbook.Names("SW_CHECK_BODY_NAME").RefersToRange.Value2, "Sun"), "Stored worksheet formulas, independently recalculated."
    If IsNumeric(actual) And Not IsError(actual) Then
        RecordCheck "Formula2 longitude", SafeText(actual), SafeText(ThisWorkbook.Names("SW_CHECK_LONGITUDE").RefersToRange.Value2), NearNumber(ThisWorkbook.Names("SW_CHECK_LONGITUDE").RefersToRange.Value2, CDbl(actual), 0.000000001), "Worksheet formula versus direct call."
    Else
        RecordCheck "Formula2 longitude", "Numeric direct result", SafeText(actual), False, "Direct scalar result is invalid."
    End If

    If allDoubles Then
        sheetVector = ThisWorkbook.Names("SW_CHECK_CALC").RefersToRange.Resize(1, 6).Value2
        For col = 1 To 6
            If Not NearNumber(sheetVector(1, col), CDbl(vector(rowLower, colLower + col - 1)), 0.000000001) Then allDoubles = False
        Next col
    End If
    RecordCheck "Formula2 calculation spill", "Six adjacent numeric values", SafeText(ThisWorkbook.Names("SW_CHECK_CALC").RefersToRange.Value2), allDoubles, "Compares every spilled value to the direct vector."

    Set detailRange = ThisWorkbook.Names("SW_CHECK_DETAIL").RefersToRange
    On Error Resume Next
    spilledDetail = detailRange.SpillingToRange.Value2
    fatalNumber = Err.Number
    Err.Clear
    On Error GoTo Fatal
    If fatalNumber = 0 Then
        CheckDetail spilledDetail, "Formula2 detail spill", actual, expectedEngine
    Else
        RecordCheck "Formula2 detail spill", "Spilled field/value/unit table", "Unable to read spill", False, "SpillingToRange failed; error " & CStr(fatalNumber)
    End If
    value = ThisWorkbook.Names("SW_CHECK_BLOCKED_SPILL").RefersToRange.Value2
    found = False
    If IsError(value) Then found = CStr(value) = CStr(CVErr(2045))
    RecordCheck "Blocked spill demonstration", "#SPILL! (2045)", SafeText(value), found, "C63 intentionally blocks output. This expected error is a PASS."
    GoTo Finish

Fatal:
    fatalNumber = Err.Number
    fatalDescription = Err.Description
    On Error Resume Next
    If Not mReport Is Nothing Then RecordCheck "Fatal smoke error", "No unhandled errors", CStr(fatalNumber), False, fatalDescription
    If mFailed = 0 Then mFailed = 1

Finish:
    On Error Resume Next
    SW_RunIntegrationChecks = "PASS=" & CStr(mPassed) & ";FAIL=" & CStr(mFailed)
    If Not mReport Is Nothing Then
        mReport.Range("G15").Value2 = "Smoke summary"
        mReport.Range("H15").Value2 = SW_RunIntegrationChecks
        mReport.Range("G16").Value2 = "Coverage"
        mReport.Range("H16").Value2 = "18 smoke checks only; full API comparison is recorded separately."
        mReport.Range("E5:E23").WrapText = True
        mReport.Range("A5:E23").Rows.AutoFit
    End If
    Application.ScreenUpdating = previousUpdating
    Set mReport = Nothing
End Function

Private Sub CheckDetail(ByRef detail As Variant, ByVal testName As String, ByVal longitude As Variant, ByVal expectedEngine As String)
    Dim foundLongitude As Boolean, foundFlags As Boolean, foundPath As Boolean
    Dim foundWarning As Boolean, foundStatus As Boolean
    Dim actualLongitude As Variant, flags As Variant, engine As Variant
    Dim warning As Variant, status As Variant, valid As Boolean
    actualLongitude = DetailField(detail, "Longitude", foundLongitude)
    flags = DetailField(detail, "Actual flags", foundFlags)
    engine = DetailField(detail, "Engine path", foundPath)
    warning = DetailField(detail, "Warning", foundWarning)
    status = DetailField(detail, "Status", foundStatus)
    valid = foundLongitude And foundFlags And foundPath And foundWarning And foundStatus
    If valid Then
        valid = IsNumeric(flags) And Not IsError(flags) And Not IsError(status)
        If valid Then valid = CLng(flags) = 258
        If valid Then valid = TextMatches(status, "OK")
        If valid Then valid = NormalizePath(SafeText(engine)) = NormalizePath(expectedEngine)
        If valid Then
            If IsNumeric(longitude) And Not IsError(longitude) Then
                valid = NearNumber(actualLongitude, CDbl(longitude), 0.000000001)
            Else
                valid = False
            End If
        End If
    End If
    RecordCheck testName, "Required fields; flags 258; Status OK; scalar/path parity", DescribeArray(detail), valid, "Status=" & SafeText(status) & "; flags=" & SafeText(flags) & "; warning=" & SafeText(warning)
End Sub

Private Function DetailField(ByRef table As Variant, ByVal label As String, ByRef found As Boolean) As Variant
    Dim r0 As Long, r1 As Long, c0 As Long, c1 As Long, row As Long
    found = False
    If Not ArrayBounds2D(table, r0, r1, c0, c1) Then Exit Function
    If c1 - c0 < 1 Then Exit Function
    For row = r0 To r1
        If TextMatches(table(row, c0), label) Then
            DetailField = table(row, c0 + 1)
            found = True
            Exit Function
        End If
    Next row
End Function

Private Function ArrayBounds2D(ByRef value As Variant, ByRef r0 As Long, ByRef r1 As Long, ByRef c0 As Long, ByRef c1 As Long) As Boolean
    On Error GoTo InvalidArray
    If Not IsArray(value) Then Exit Function
    r0 = LBound(value, 1): r1 = UBound(value, 1)
    c0 = LBound(value, 2): c1 = UBound(value, 2)
    ArrayBounds2D = True
InvalidArray:
End Function

Private Function DescribeArray(ByRef value As Variant) As String
    Dim r0 As Long, r1 As Long, c0 As Long, c1 As Long
    If ArrayBounds2D(value, r0, r1, c0, c1) Then
        DescribeArray = CStr(r1 - r0 + 1) & " x " & CStr(c1 - c0 + 1) & " array"
    Else
        DescribeArray = SafeText(value)
    End If
End Function

Private Function NearNumber(ByVal value As Variant, ByVal expected As Double, ByVal tolerance As Double) As Boolean
    On Error GoTo InvalidValue
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    NearNumber = Abs(CDbl(value) - expected) <= tolerance
InvalidValue:
End Function

Private Function TextMatches(ByVal value As Variant, ByVal expected As String) As Boolean
    If IsError(value) Or IsNull(value) Or IsArray(value) Then Exit Function
    TextMatches = StrComp(Trim$(Replace(CStr(value), vbNullChar, "")), Trim$(expected), vbTextCompare) = 0
End Function

Private Function SafeText(ByVal value As Variant) As String
    On Error GoTo InvalidValue
    If IsArray(value) Then
        SafeText = "Array"
    ElseIf IsNull(value) Then
        SafeText = "Null"
    ElseIf IsEmpty(value) Then
        SafeText = "Empty"
    Else
        SafeText = CStr(value)
    End If
    Exit Function
InvalidValue:
    SafeText = "Unreadable value"
End Function

Private Function NormalizePath(ByVal value As String) As String
    value = Replace(value, "/", "\")
    If Left$(value, 4) = "\\?\" Then value = Mid$(value, 5)
    NormalizePath = LCase$(value)
End Function

Private Sub RecordCheck(ByVal name As String, ByVal expected As String, ByVal actual As String, ByVal passed As Boolean, ByVal notes As String)
    If passed Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
    End If
    mReport.Cells(mNextRow, 1).Value2 = name
    mReport.Cells(mNextRow, 2).Value2 = expected
    mReport.Cells(mNextRow, 3).Value2 = actual
    mReport.Cells(mNextRow, 4).Value2 = IIf(passed, "PASS", "FAIL")
    mReport.Cells(mNextRow, 5).Value2 = notes
    If passed Then
        mReport.Cells(mNextRow, 4).Interior.Color = RGB(221, 237, 232)
    Else
        mReport.Cells(mNextRow, 4).Interior.Color = RGB(249, 220, 214)
    End If
    mNextRow = mNextRow + 1
End Sub

Private Sub WriteEnvironment()
    mReport.Range("G4").Value2 = "Environment"
    mReport.Range("H4").Value2 = "Actual value"
    mReport.Range("G5").Value2 = "Application.Version"
    mReport.Range("H5").Value2 = CStr(Application.Version)
    mReport.Range("G6").Value2 = "Application.Build"
    mReport.Range("H6").Value2 = CStr(Application.Build)
    mReport.Range("G7").Value2 = "OperatingSystem"
    mReport.Range("H7").Value2 = Application.OperatingSystem
    mReport.Range("G8").Value2 = "Host process architecture"
    mReport.Range("H8").Value2 = ProcessArchitecture()
    mReport.Range("G9").Value2 = "Workbook path"
    mReport.Range("H9").Value2 = ThisWorkbook.FullName
    mReport.Range("G10").Value2 = "Native module path"
    mReport.Range("H10").Value2 = SWEnginePath()
    mReport.Range("G11").Value2 = "Data path"
    mReport.Range("H11").Value2 = SWDataPath()
    mReport.Range("G12").Value2 = "Runtime status"
    mReport.Range("H12").Value2 = SWRuntimeStatus()
    mReport.Range("G13").Value2 = "Run time (local)"
    mReport.Range("H13").Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Function ProcessArchitecture() As String
#If Win64 Then
    Dim processMachine As Integer, nativeMachine As Integer, result As Long
    On Error GoTo ApiUnavailable
    result = IsWow64Process2(GetCurrentProcess(), processMachine, nativeMachine)
    If result <> 0 Then
        ProcessArchitecture = "Win64 VBA; processMachine=0x" & Hex$(CLng(processMachine) And &HFFFF&) & "; nativeMachine=0x" & Hex$(CLng(nativeMachine) And &HFFFF&)
        Exit Function
    End If
ApiUnavailable:
    ProcessArchitecture = "Win64 VBA; IsWow64Process2 unavailable or failed"
#Else
    ProcessArchitecture = "Not Win64 VBA"
#End If
End Function
