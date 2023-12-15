Global Natal(0 To 15) As Double, Planets(0 To 15) As Double
Global NatalDecl(0 To 15) As Double
Global NatalHouse(13) As Double
Global HC(0 To 13) As Double, HC3(0 To 13) As Double
Global NatalPH$
Global HCX1$, PH$, PN$(15), PX1$, SX1$
Global epsilon As Double, crlf As String
Global Result_JD As Double

Global MyiDate%, JDDate$, JDTime$

Global SORT() As Double, SORTPOS() As Double

Global OutputLines$(), TextOnForm$(), TextLineCounter%, LineCounter%, Buffer As String

Global p1_pos As Double, p1_pos_later As Double, p2_pos As Double, p2_pos_later As Double

Global XHousePower!(), XHousePowerVariation!(), DegreesInEachHouse() As Double
Global DistanceFromCusp() As Double, PlanetHousePower() As Double, PlanetAspectPower() As Double
Global PlanetHarmony() As Double, PlanetDiscord() As Double, PlanetTotalPower() As Double
Global InterceptedSign%(), InterceptedHouse%(), Cosmo!()
Global SignPower() As Double, SignHarmony() As Double, SignDiscord() As Double
Global HousePower() As Double, HouseHarmony() As Double, HouseDiscord() As Double

' Swiss Ephemeris Release 1.60  9-jan-2000
' Declarations for Visual Basic 5.0
' The DLL file must exist in the same directory as the VB executable, or in a system
' directory where it can be found at runtime




Declare PtrSafe Function swe_calc Lib "C:\sweph\bin\swedll64.dll" _
        ( _
          ByVal tjd As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef x As Double, _
          ByVal serr As String _
        ) As Long   ' x must be first of six array elements
                    ' serr must be able to hold 256 bytes

Declare PtrSafe Function swe_calc_ut Lib "C:\sweph\bin\swedll64.dll" _
        ( _
          ByVal tjd_ut As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef x As Double, _
          ByVal serr As String _
        ) As Long   ' x must be first of six array elements
                    ' serr must be able to hold 256 bytes

Declare PtrSafe Function swe_close Lib "C:\sweph\bin\swedll64.dll" _
        () As Long
        'Alias "_swe_close@0" () As Long

Declare PtrSafe Function swe_date_conversion Lib "C:\sweph\bin\swedll64.dll" _
        ( _
          ByVal Year As Long, _
          ByVal Month As Long, _
          ByVal Day As Long, _
          ByVal utime As Double, _
          ByVal cal As Byte, _
          ByRef tjd As Double _
        ) As Long

Declare PtrSafe Function swe_deltat Lib "C:\sweph\bin\swedll64.dll" _
        (ByVal JD As Double) As Double

Declare PtrSafe Function swe_houses Lib "C:\sweph\bin\swedll64.dll" _
        ( _
          ByVal tjd_ut As Double, _
          ByVal geolat As Double, _
          ByVal geolon As Double, _
          ByVal ihsy As Long, _
          ByRef hcusps As Double, _
          ByRef ascmc As Double _
        ) As Long       ' hcusps must be first of 13 array elements
                        ' ascmc must be first of 10 array elements

Declare PtrSafe Function swe_house_pos Lib "C:\sweph\bin\swedll64.dll" _
        ( _
          ByVal armc As Double, _
          ByVal geolat As Double, _
          ByVal eps As Double, _
          ByVal ihsy As Long, _
          ByRef xpin As Double, _
          ByVal serr As String _
        ) As Double
                        ' xpin must be first of 2 array elements



Declare PtrSafe Function swe_julday Lib "C:\sweph\bin\swedll64.dll" _
        ( _
          ByVal Year As Long, _
          ByVal Month As Long, _
          ByVal Day As Long, _
          ByVal hour As Double, _
          ByVal gregflg As Long _
        ) As Double

Declare PtrSafe Sub swe_revjul Lib "C:\sweph\bin\swedll64.dll" _
        Alias "_swe_revjul" ( _
          ByValtjd As Double, _
          ByVal gregflg As Long, _
          ByRef Year As Long, _
          ByRef Month As Long, _
          ByRef Day As Long, _
          ByRef hour As Double _
        )

Public Declare PtrSafe Function swe_version Lib "C:\sweph\bin\swedll64.dll" _
        ( _
          ByVal svers As String _
        ) As String






' values for gregflag in swe_julday() and swe_revjul()
 Const SE_JUL_CAL As Integer = 0
 Const SE_GREG_CAL As Integer = 1

' planet and body numbers (parameter ipl) for swe_calc()
 Const SE_ECL_NUT As Integer = -1
 Const SE_SUN As Integer = 0
 Const SE_MOON As Integer = 1
 Const SE_MERCURY As Integer = 2
 Const SE_VENUS As Integer = 3
 Const SE_MARS As Integer = 4
 Const SE_JUPITER As Integer = 5
 Const SE_SATURN As Integer = 6
 Const SE_URANUS As Integer = 7
 Const SE_NEPTUNE As Integer = 8
 Const SE_PLUTO   As Integer = 9
 Const SE_MEAN_NODE As Integer = 10
 Const SE_TRUE_NODE As Integer = 11
 Const SE_MEAN_APOG As Integer = 12
 Const SE_OSCU_APOG As Integer = 13
 Const SE_EARTH     As Integer = 14
 Const SE_CHIRON    As Integer = 15

' points returned by swe_houses() and swe_houses_armc() in array ascmc(0...10)
 Const SE_ASC       As Integer = 0
 Const SE_MC        As Integer = 1
 Const SE_ARMC      As Integer = 2
 Const SE_VERTEX    As Integer = 3
 Const SE_EQUASC    As Integer = 4      ' "equatorial ascendant"
 Const SE_NASCMC    As Integer = 5      ' number of such points

' iflag values for swe_calc()/swe_calc_ut() and swe_fixstar()/swe_fixstar_ut()
Const SEFLG_JPLEPH As Long = 1
Const SEFLG_SWIEPH As Long = 2
Const SEFLG_MOSEPH As Long = 4
Const SEFLG_SPEED As Long = 256
Const SEFLG_HELCTR As Long = 8
Const SEFLG_TRUEPOS As Long = 16
Const SEFLG_J2000 As Long = 32
Const SEFLG_NONUT As Long = 64
Const SEFLG_NOGDEFL As Long = 512
Const SEFLG_NOABERR As Long = 1024
Const SEFLG_EQUATORIAL As Long = 2048
Const SEFLG_RADIANS As Long = 8192
Const SEFLG_TOPOCTR As Long = 32768
Const SEFLG_SIDEREAL As Long = 65536