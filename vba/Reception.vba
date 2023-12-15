Sub Reception()
    For y = 1 To 10: For x = 1 To 10
    sy = Int(Natal(y) / 30) + 1: SX = Int(Natal(x) / 30) + 1
    If y = 1 And (sy = 4 Or sy = 2) And x = 10 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 1 And (sy = 3 Or sy = 6 Or sy = 11) And x = 2 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 1 And (sy = 2 Or sy = 7 Or sy = 12) And x = 3 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 1 And (sy = 1 Or sy = 8 Or sy = 10) And x = 4 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 1 And (sy = 4 Or sy = 9 Or sy = 12) And x = 5 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 1 And (sy = 7 Or sy = 10 Or sy = 11) And x = 6 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 1 And (sy = 3 Or sy = 11) And x = 7 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 1 And (sy = 9 Or sy = 12) And x = 8 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 1 And (sy = 5 Or sy = 8) And x = 9 And (SX = 1 Or SX = 5) Then GoSub MRSE
    If y = 10 And (sy = 3 Or sy = 6 Or sy = 11) And x = 2 And (SX = 2 Or SX = 4) Then GoSub MRSE
    If y = 10 And (sy = 2 Or sy = 7 Or sy = 12) And x = 3 And (SX = 2 Or SX = 4) Then GoSub MRSE
    If y = 10 And (sy = 1 Or sy = 8 Or sy = 10) And x = 4 And (SX = 2 Or SX = 4) Then GoSub MRSE
    If y = 10 And (sy = 4 Or sy = 9 Or sy = 12) And x = 5 And (SX = 2 Or SX = 4) Then GoSub MRSE
    If y = 10 And (sy = 7 Or sy = 10 Or sy = 11) And x = 6 And (SX = 2 Or SX = 4) Then GoSub MRSE
    If y = 10 And (sy = 3 Or sy = 11) And x = 7 And (SX = 2 Or SX = 4) Then GoSub MRSE
    If y = 10 And (sy = 9 Or sy = 12) And x = 8 And (SX = 2 Or SX = 4) Then GoSub MRSE
    If y = 10 And (sy = 5 Or sy = 8) And x = 9 And (SX = 2 Or SX = 4) Then GoSub MRSE
    If y = 2 And (sy = 2 Or sy = 7 Or sy = 12) And x = 3 And (SX = 3 Or SX = 6 Or SX = 11) Then GoSub MRSE
    If y = 2 And (sy = 1 Or sy = 8 Or sy = 10) And x = 4 And (SX = 3 Or SX = 6 Or SX = 11) Then GoSub MRSE
    If y = 2 And (sy = 4 Or sy = 9 Or sy = 12) And x = 5 And (SX = 3 Or SX = 6 Or SX = 11) Then GoSub MRSE
    If y = 2 And (sy = 7 Or sy = 10 Or sy = 11) And x = 6 And (SX = 3 Or SX = 6 Or SX = 11) Then GoSub MRSE
    If y = 2 And (sy = 3 Or sy = 11) And x = 7 And (SX = 3 Or SX = 6 Or SX = 11) Then GoSub MRSE
    If y = 2 And (sy = 9 Or sy = 12) And x = 8 And (SX = 3 Or SX = 6 Or SX = 11) Then GoSub MRSE
    If y = 2 And (sy = 5 Or sy = 8) And x = 9 And (SX = 3 Or SX = 6 Or SX = 11) Then GoSub MRSE
    If y = 3 And (sy = 1 Or sy = 8 Or sy = 10) And x = 4 And (SX = 2 Or SX = 7 Or SX = 12) Then GoSub MRSE
    If y = 3 And (sy = 4 Or sy = 9 Or sy = 12) And x = 5 And (SX = 2 Or SX = 7 Or SX = 12) Then GoSub MRSE
    If y = 3 And (sy = 7 Or sy = 10 Or sy = 11) And x = 6 And (SX = 2 Or SX = 7 Or SX = 12) Then GoSub MRSE
    If y = 3 And (sy = 3 Or sy = 11) And x = 7 And (SX = 2 Or SX = 7 Or SX = 12) Then GoSub MRSE
    If y = 3 And (sy = 9 Or sy = 12) And x = 8 And (SX = 2 Or SX = 7 Or SX = 12) Then GoSub MRSE
    If y = 3 And (sy = 5 Or sy = 8) And x = 9 And (SX = 2 Or SX = 7 Or SX = 12) Then GoSub MRSE
    If y = 4 And (sy = 4 Or sy = 9 Or sy = 12) And x = 5 And (SX = 1 Or SX = 8 Or SX = 10) Then GoSub MRSE
    If y = 4 And (sy = 7 Or sy = 10 Or sy = 11) And x = 6 And (SX = 1 Or SX = 8 Or SX = 10) Then GoSub MRSE
    If y = 4 And (sy = 3 Or sy = 11) And x = 7 And (SX = 1 Or SX = 8 Or SX = 10) Then GoSub MRSE
    If y = 4 And (sy = 9 Or sy = 12) And x = 8 And (SX = 1 Or SX = 8 Or SX = 10) Then GoSub MRSE
    If y = 4 And (sy = 5 Or sy = 8) And x = 9 And (SX = 1 Or SX = 8 Or SX = 10) Then GoSub MRSE
    If y = 5 And (sy = 7 Or sy = 10 Or sy = 11) And x = 6 And (SX = 4 Or SX = 9 Or SX = 12) Then GoSub MRSE
    If y = 5 And (sy = 3 Or sy = 11) And x = 7 And (SX = 4 Or SX = 9 Or SX = 12) Then GoSub MRSE
    If y = 5 And (sy = 9 Or sy = 12) And x = 8 And (SX = 4 Or SX = 9 Or SX = 12) Then GoSub MRSE
    If y = 5 And (sy = 5 Or sy = 8) And x = 9 And (SX = 4 Or SX = 9 Or SX = 12) Then GoSub MRSE
    If y = 6 And (sy = 3 Or sy = 11) And x = 7 And (SX = 7 Or SX = 10 Or SX = 11) Then GoSub MRSE
    If y = 6 And (sy = 9 Or sy = 12) And x = 8 And (SX = 7 Or SX = 10 Or SX = 11) Then GoSub MRSE
    If y = 6 And (sy = 5 Or sy = 8) And x = 9 And (SX = 7 Or SX = 10 Or SX = 11) Then GoSub MRSE
    If y = 7 And (sy = 9 Or sy = 12) And x = 8 And (SX = 3 Or SX = 11) Then GoSub MRSE
    If y = 7 And (sy = 5 Or sy = 8) And x = 9 And (SX = 3 Or SX = 11) Then GoSub MRSE
    If y = 8 And (sy = 5 Or sy = 8) And x = 9 And (SX = 9 Or SX = 12) Then GoSub MRSE
    Next x: Next y: Exit Sub
'
MRSE:   PlanetHarmony(y) = PlanetHarmony(y) + 5: PlanetHarmony(x) = PlanetHarmony(x) + 5
    Return
End Sub