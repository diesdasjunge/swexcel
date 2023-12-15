Sub Dignities()
'find dignity
    For x = 1 To 10: TEMP7 = Int(Natal(x) / 30) + 1
    If x = 1 And TEMP7 = 5 Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 10 And TEMP7 = 4 Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 2 And (TEMP7 = 3 Or TEMP7 = 6) Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 3 And (TEMP7 = 2 Or TEMP7 = 7) Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 4 And (TEMP7 = 1 Or TEMP7 = 8) Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 5 And (TEMP7 = 9 Or TEMP7 = 12) Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 6 And (TEMP7 = 10 Or TEMP7 = 11) Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 7 And TEMP7 = 11 Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 8 And TEMP7 = 12 Then PlanetHarmony(x) = PlanetHarmony(x) + 2
    If x = 9 And TEMP7 = 8 Then PlanetHarmony(x) = PlanetHarmony(x) + 2
'find detriment
    If x = 1 And TEMP7 = 11 Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 10 And TEMP7 = 10 Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 2 And (TEMP7 = 9 Or TEMP7 = 12) Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 3 And (TEMP7 = 1 Or TEMP7 = 8) Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 4 And (TEMP7 = 2 Or TEMP7 = 7) Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 5 And (TEMP7 = 3 Or TEMP7 = 6) Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 6 And (TEMP7 = 4 Or TEMP7 = 5) Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 7 And TEMP7 = 5 Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 8 And TEMP7 = 6 Then PlanetDiscord(x) = PlanetDiscord(x) + 2
    If x = 9 And TEMP7 = 2 Then PlanetDiscord(x) = PlanetDiscord(x) + 2
'find exaltation
    If x = 1 And TEMP7 = 1 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 10 And TEMP7 = 2 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 2 And TEMP7 = 11 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 3 And TEMP7 = 12 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 4 And TEMP7 = 10 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 5 And TEMP7 = 4 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 6 And TEMP7 = 7 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 7 And TEMP7 = 3 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 8 And TEMP7 = 9 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
    If x = 9 And TEMP7 = 5 Then PlanetHarmony(x) = PlanetHarmony(x) + 3
'find fall
    If x = 1 And TEMP7 = 7 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 10 And TEMP7 = 8 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 2 And TEMP7 = 5 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 3 And TEMP7 = 6 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 4 And TEMP7 = 4 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 5 And TEMP7 = 10 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 6 And TEMP7 = 1 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 7 And TEMP7 = 9 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 8 And TEMP7 = 3 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
    If x = 9 And TEMP7 = 11 Then PlanetDiscord(x) = PlanetDiscord(x) + 3
'find harmony
    If x = 1 And TEMP7 = 9 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 10 And TEMP7 = 12 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 2 And TEMP7 = 8 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 3 And TEMP7 = 11 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 4 And TEMP7 = 5 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 5 And TEMP7 = 2 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 6 And TEMP7 = 6 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 7 And TEMP7 = 7 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 8 And TEMP7 = 4 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
    If x = 9 And TEMP7 = 1 Then PlanetHarmony(x) = PlanetHarmony(x) + 1
'find inharmony
    If x = 1 And TEMP7 = 3 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 10 And TEMP7 = 6 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 2 And TEMP7 = 2 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 3 And TEMP7 = 5 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 4 And TEMP7 = 11 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 5 And TEMP7 = 8 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 6 And TEMP7 = 12 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 7 And TEMP7 = 1 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 8 And TEMP7 = 10 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    If x = 9 And TEMP7 = 7 Then PlanetDiscord(x) = PlanetDiscord(x) + 1
    Next x
End Sub