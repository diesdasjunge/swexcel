Sub PrintBuffer()
    On Error GoTo PBErr1
    
    ReDim Preserve OutputLines$(0 To LineCounter% + 1)
    OutputLines$(LineCounter%) = Buffer
    LineCounter% = LineCounter% + 1

    TextOnForm$(TextLineCounter%) = Buffer
    TextLineCounter% = TextLineCounter% + 1
    ReDim Preserve TextOnForm$(0 To TextLineCounter%)

PBErr2:
    Exit Sub

PBErr1: MsgBox "Buffer cannot hold any more data. Please print and/or clear data buffer before displaying any more data. This message may appear up to 35 times depending on how full the buffer is and how much the program wants to write to it.", 16, "Warning!"
    Resume PBErr2
End Sub