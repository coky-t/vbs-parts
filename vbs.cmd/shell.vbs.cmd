REM: %~nx0 & @cscript.exe //e:vbscript //nologo "%~f0" %* & GOTO :EOF

If WScript.Arguments.Length = 1 Then
    Exec Replace(WScript.Arguments(0),"`","""")
    WScript.Quit
End If

Do While True
    WScript.StdOut.Write "vbs>"
    Line = WScript.StdIn.ReadLine
    If LCase(Trim(Line)) = "exit" Then Exit Do
    Exec Line
Loop

Sub Exec(Line)
    On Error Resume Next
    Err.Clear
    Execute Line
    If Err.Number <> 0 Then WScript.StdOut.WriteLine(Err.Description)
    On Error Goto 0
End Sub
