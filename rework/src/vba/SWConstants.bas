Attribute VB_Name = "SWConstants"
Option Explicit
Option Private Module

' SPDX-License-Identifier: AGPL-3.0-or-later
' Values taken from the pinned Swiss Ephemeris public header.
Public Const SW_ENGINE_FILE As String = "swexcel-se-2.10.3b-x64.dll"
Public Const SW_ENGINE_VERSION As String = "2.10.03"
Public Const SW_DEFAULT_FLAGS As Long = 258
Public Const SW_FLAG_JPL As Long = 1
Public Const SW_FLAG_SWISS As Long = 2
Public Const SW_FLAG_MOSHIER As Long = 4
Public Const SW_FLAG_EQUATORIAL As Long = 2048
Public Const SW_FLAG_XYZ As Long = 4096
Public Const SW_FLAG_RADIANS As Long = 8192
Public Const SW_FLAG_TOPOCENTRIC As Long = 32768
Public Const SW_FLAG_SIDEREAL As Long = 65536
Public Const SW_ERROR_BYTES As Long = 256
Public Const SW_NAME_BYTES As Long = 256
' swe_set_ephe_path reserves 13 bytes inside AS_MAXCH for generated filenames.
Public Const SW_EPHE_PATH_BYTES As Long = 243
Public Const SW_TIDAL_AUTOMATIC As Double = 999999#
Public Const SW_DELTAT_AUTOMATIC As Double = -0.0000000001
