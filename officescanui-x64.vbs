Set WshShell = WScript.CreateObject("WScript.Shell")
set FSO = CreateObject("Scripting.FileSystemObject")
strApp = "C:\Program Files (x86)\Trend Micro\OfficeScan Client\ntrmv.exe"
strPara1 = "-991334"
strPara2 = "-442"
 
Dim myExit, return
myExit = 0
 
currentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
 
' Run UnInstall of TrendMicro
WshShell.run Chr(34) & strApp & Chr(34) & " " & Chr(34) & strPara1 & Chr(34) & " " & Chr(34) & strPara2 & Chr(34), 0, True
 
' Activate the loop until result is "myExit" = 1
Do Until myExit = 1
' Triggers the check on the active "ntrmv.exe" process
	CheckTrendMicro
Loop
 
SUB CheckTrendMicro()
 
myExit = 1
set service = GetObject ("winmgmts:")
' Check for active ntrmv.exe process.
for each Process in Service.InstancesOf ("Win32_Process")
	If Process.Name = "ntrmv.exe" then
			myExit = 0
			' wait for X time before checking for running process again.
			Wscript.sleep(60000)
	End if
NEXT
End SUB