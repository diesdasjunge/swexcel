Sub AdjustForm()
    Dim x%, t$

    If LineCounter% = 0 Then Exit Sub
    
    Call LineFeed

    For x% = 0 To TextLineCounter%
        t$ = t$ & TextOnForm$(x%) & crlf
    Next x%
    
    ActiveSheet.TextBox1.Text = t$
End Sub