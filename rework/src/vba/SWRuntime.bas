Attribute VB_Name = "SWRuntime"
Option Explicit
Option Private Module

' SPDX-License-Identifier: AGPL-3.0-or-later
' The DLL stays loaded for the Excel process lifetime. VBA Declare calls cache
' addresses: unloading it could leave a cached address pointing at freed memory.
#If Mac Then
#ElseIf Win64 Then
Private Declare PtrSafe Function LoadLibraryExW Lib "kernel32" (ByVal fileName As LongPtr, ByVal fileHandle As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal moduleName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleFileNameW Lib "kernel32" (ByVal moduleHandle As LongPtr, ByVal fileName As LongPtr, ByVal capacity As Long) As Long
Private Declare PtrSafe Function GetFullPathNameW Lib "kernel32" (ByVal fileName As LongPtr, ByVal capacity As Long, ByVal result As LongPtr, ByVal filePart As LongPtr) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal moduleHandle As LongPtr, ByVal procedureName As String) As LongPtr
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal wideText As LongPtr, ByVal wideLength As Long, ByVal destination As LongPtr, ByVal capacity As Long, ByVal defaultCharacter As LongPtr, ByVal usedDefaultCharacter As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal text As LongPtr, ByVal textLength As Long, ByVal destination As LongPtr, ByVal capacity As Long) As Long
Private mEngine As LongPtr
#End If

Private mBusy As Boolean
Private mReady As Boolean
Private mVerifiedPath As String
Private mVerifiedVersion As String

Public Sub SWRaise(ByVal message As String)
    Err.Raise vbObjectError + 2103, "SWExcel", message
End Sub

Private Function PackagePath() As String
    If Len(ThisWorkbook.Path) = 0 Then SWRaise "Save the workbook in the extracted SWExcel package first."
    PackagePath = ThisWorkbook.Path
End Function

Private Function CanonicalPath(ByVal path As String) As String
#If Mac Then
    SWRaise "SWExcel requires 64-bit Microsoft 365 Excel on Windows."
#ElseIf Win64 Then
    Dim buffer As String, count As Long
    buffer = String$(32768, vbNullChar)
    count = GetFullPathNameW(StrPtr(path), Len(buffer), StrPtr(buffer), 0)
    If count = 0 Or count >= Len(buffer) Then SWRaise "Cannot resolve package path."
    CanonicalPath = Left$(buffer, count)
#Else
    SWRaise "SWExcel requires 64-bit Microsoft 365 Excel on Windows."
#End If
End Function

#If Mac Then
#ElseIf Win64 Then
Private Function ModulePath(ByVal handle As LongPtr) As String
    Dim buffer As String, count As Long
    buffer = String$(32768, vbNullChar)
    count = GetModuleFileNameW(handle, StrPtr(buffer), Len(buffer))
    If count = 0 Or count >= Len(buffer) Then SWRaise "Cannot verify the loaded DLL location."
    ModulePath = CanonicalPath(Left$(buffer, count))
End Function
#End If

Public Function SWAnsiZ(ByVal value As String, ByVal maximumBytes As Long) As Byte()
    Dim bytes() As Byte
#If Mac Then
    SWRaise "Windows is required."
#ElseIf Win64 Then
    Dim count As Long, usedDefault As Long, conversionFlags As Long, defaultPointer As LongPtr
    If InStr(1, value, vbNullChar, vbBinaryCompare) <> 0 Then SWRaise "C strings cannot contain a NUL character."
    If Len(value) = 0 Then
        ReDim bytes(0 To 0)
    Else
        ' CP_ACP matches the Windows C runtime. Reject lossy/best-fit encoding.
        conversionFlags = &H400
        defaultPointer = VarPtr(usedDefault)
        If GetACP() = 65001 Then
            conversionFlags = 0
            defaultPointer = 0
        End If
        count = WideCharToMultiByte(0, conversionFlags, StrPtr(value), Len(value), 0, 0, 0, defaultPointer)
        If count <= 0 Or usedDefault <> 0 Then SWRaise "The package path cannot be represented by the Windows ANSI code page. Move it to an ASCII path."
        If count + 1 > maximumBytes Then SWRaise "The encoded C string exceeds the engine buffer limit. Use a shorter package path."
        ReDim bytes(0 To count)
        usedDefault = 0
        If WideCharToMultiByte(0, conversionFlags, StrPtr(value), Len(value), VarPtr(bytes(0)), count, 0, defaultPointer) <> count Or usedDefault <> 0 Then SWRaise "C string encoding failed."
    End If
#Else
    SWRaise "64-bit Windows Excel is required."
#End If
    SWAnsiZ = bytes
End Function

Public Function SWBufferText(ByRef bytes() As Byte) As String
#If Mac Then
    SWRaise "Windows is required."
#ElseIf Win64 Then
    Dim start As Long, count As Long, wideCount As Long, result As String
    start = LBound(bytes)
    For count = 0 To UBound(bytes) - start
        If bytes(start + count) = 0 Then Exit For
    Next count
    If count > UBound(bytes) - start Then SWRaise "The engine returned an unterminated C string."
    If count = 0 Then Exit Function
    wideCount = MultiByteToWideChar(0, 0, VarPtr(bytes(start)), count, 0, 0)
    If wideCount <= 0 Then SWRaise "Cannot decode the engine string."
    result = String$(wideCount, vbNullChar)
    If MultiByteToWideChar(0, 0, VarPtr(bytes(start)), count, StrPtr(result), wideCount) <> wideCount Then SWRaise "Cannot decode the engine string."
    SWBufferText = result
#Else
    SWRaise "64-bit Windows Excel is required."
#End If
End Function

Public Sub SWEnsureEngine()
#If Mac Then
    SWRaise "SWExcel requires 64-bit Microsoft 365 Excel on Windows."
#ElseIf Win64 Then
    Dim expected As String, existing As LongPtr, loaded As String
    Dim name As String, data As String, pathBytes() As Byte, file As Variant
    Dim versionBuffer(0 To 255) As Byte, versionPointer As LongPtr
    expected = CanonicalPath(PackagePath() & "\runtime\engine\" & SW_ENGINE_FILE)
    data = CanonicalPath(PackagePath() & "\runtime\ephe")
    If Len(Environ$("SE_EPHE_PATH")) <> 0 Then SWRaise "SE_EPHE_PATH overrides the package data location. Clear it before starting Excel."
    If InStr(1, data, ";", vbBinaryCompare) <> 0 Then SWRaise "The engine treats semicolons as data-path separators. Move the package to a folder without semicolons."
    pathBytes = SWAnsiZ(data, SW_EPHE_PATH_BYTES)
    For Each file In Array("sepl_18.se1", "semo_18.se1", "seas_18.se1", "sefstars.txt", "seasnam.txt", "seorbel.txt")
        If Len(Dir$(data & "\" & CStr(file))) = 0 Then SWRaise "Required package data is missing: " & CStr(file)
    Next file
    If Len(Dir$(expected)) = 0 Then SWRaise "The package DLL is missing: " & expected
    name = SW_ENGINE_FILE
    existing = GetModuleHandleW(StrPtr(name))
    If existing <> 0 Then
        loaded = ModulePath(existing)
        If StrComp(loaded, expected, vbTextCompare) <> 0 Then SWRaise "A different SWExcel package DLL is already loaded: " & loaded & ". Open this package in a separate Excel process."
    End If
    If mEngine <> 0 Then
        If StrComp(mVerifiedPath, expected, vbTextCompare) <> 0 Then SWRaise "This workbook moved after loading its DLL. Restart Excel before using the new location."
        If existing <> mEngine Then SWRaise "The loaded engine identity changed. Restart Excel."
        If mReady Then Exit Sub
    Else
        ' Only this absolute directory and System32 participate in dependency lookup.
        mEngine = LoadLibraryExW(StrPtr(expected), 0, &H900)
        If mEngine = 0 Then SWRaise "Windows could not load the x64 engine (error " & CStr(GetLastError()) & "). Check Excel process architecture and the package files."
        mVerifiedPath = ModulePath(mEngine)
    End If
    If StrComp(mVerifiedPath, expected, vbTextCompare) <> 0 Then SWRaise "Windows loaded an unexpected engine path."
    For Each file In Array("swe_version", "swe_calc_ut", "swe_julday", "swe_get_planet_name", "swe_set_ephe_path")
        If GetProcAddress(mEngine, CStr(file)) = 0 Then SWRaise "The engine is missing export " & CStr(file)
    Next file
    versionPointer = Native_swe_version(versionBuffer(0))
    ' Upstream returns the exact caller buffer. Do not scan an arbitrary pointer.
    If versionPointer <> VarPtr(versionBuffer(0)) Then SWRaise "The engine returned an unexpected version pointer."
    mVerifiedVersion = SWBufferText(versionBuffer)
    If mVerifiedVersion <> SW_ENGINE_VERSION Then SWRaise "Unexpected engine runtime version: " & mVerifiedVersion
    Native_swe_set_ephe_path pathBytes(0)
    mReady = True
#Else
    SWRaise "SWExcel requires 64-bit Microsoft 365 Excel on Windows."
#End If
End Sub

Public Function SWEnginePath() As String
    SWEnsureEngine
    SWEnginePath = mVerifiedPath
End Function

Public Function SWDataPath() As String
    SWEnsureEngine
    SWDataPath = CanonicalPath(PackagePath() & "\runtime\ephe")
End Function

Public Function SWEngineVersion() As String
    SWEnsureEngine
    SWEngineVersion = mVerifiedVersion
End Function

Public Function SWRuntimeStatus() As String
    On Error GoTo Failed
    SWEnsureEngine
    SWRuntimeStatus = "Ready for integration testing; runtime " & mVerifiedVersion
    Exit Function
Failed:
    SWRuntimeStatus = "ERROR: " & Err.Description
End Function

Public Sub SWBeginCalculation()
    If mBusy Then SWRaise "A nested Swiss Ephemeris calculation was blocked."
    SWEnsureEngine
    mBusy = True
End Sub

Public Sub SWEndCalculation()
    mBusy = False
End Sub

Public Sub SWApplyCalculationOptions(ByVal siderealMode As Long, ByVal longitudeEast As Double, ByVal latitudeNorth As Double, ByVal altitudeMetres As Double)
    Dim pathBytes() As Byte, modelBytes() As Byte
    pathBytes = SWAnsiZ(SWDataPath(), SW_EPHE_PATH_BYTES)
    Native_swe_set_ephe_path pathBytes(0)
    ' Zero selects each current engine default, clearing raw model overrides.
    modelBytes = SWAnsiZ("0,0,0,0,0,0,0,0", 256)
    Native_swe_set_astro_models modelBytes(0), SW_DEFAULT_FLAGS
    Native_swe_set_delta_t_userdef SW_DELTAT_AUTOMATIC
    Native_swe_set_tid_acc SW_TIDAL_AUTOMATIC
    Native_swe_set_interpolate_nut 0
    Native_swe_set_sid_mode siderealMode, 0#, 0#
    Native_swe_set_topo longitudeEast, latitudeNorth, altitudeMetres
End Sub

' VBA command, not a worksheet function. Keep the DLL reference alive.
Public Sub SW_CLOSE()
    If mBusy Then SWRaise "Cannot close the engine during a calculation."
    SWEnsureEngine
    Native_swe_close
End Sub
