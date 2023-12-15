'###################################################################
Public Function dll_version()

Dim svers As String
svers = String(256, vbNullChar)
swe_version (svers)
dll_version = svers
End Function