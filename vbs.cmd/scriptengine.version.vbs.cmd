REM: %~nx0 & @cscript.exe //e:vbscript //nologo "%~f0" %* & GOTO :EOF

' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/zk6wed09(v=vs.84)
' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/b5zx2btt(v=vs.84)
' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/f7z0ys3a(v=vs.84)
' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/b4scfs96(v=vs.84)

WScript.Echo _
    ScriptEngine & Space(1) & _
    ScriptEngineMajorVersion & "." & _
    ScriptEngineMinorVersion & "." & _
    ScriptEngineBuildVersion
