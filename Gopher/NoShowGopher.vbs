Dim WinScriptHost
Set WinScriptHost = CreateObject("WScript.Shell")
WinScriptHost.Run Chr(34) & "D:\Run\Gopher\Gopher.exe" & Chr(34), 0
Set WinScriptHost = Nothing