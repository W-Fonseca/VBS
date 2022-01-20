oi
sub oi
Set WshShell = CreateObject("WScript.Shell")
On Error Resume Next
WshShell.Run("CMD /K query session |clip")
WshShell.Run("taskkill /f /im cmd.exe")
On Error Goto 0

texto = CreateObject("HTMLFile").parentWindow.clipboardData.getData("Text")

strName = wshShell.ExpandEnvironmentStrings("%USERNAME%")
if InStr(texto, strName) > 0 then
numero = InStr(texto, strName)

For i = numero to 99999

'WScript.Echo Mid(texto,i)

if IsNumeric(Mid(texto,i,1)) Then
WScript.Echo Mid(texto,i,1)
NumeroSessao = Mid(texto,i,1) 

exit for
end if 


Next 
end if 

strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process") 'caso queria expecificar um programa, pode por "SELECT * FROM Win32_Process WHERE Name = 'Database.exe'"
For Each objProcess in colProcessList
if objProcess.SessionId = NumeroSessao then

 Wscript.Echo "Process: " & objProcess.Name 'nome do procsso
 Wscript.Echo "Process ID: " & objProcess.ProcessID
 Wscript.Echo "Thread Count: " & objProcess.ThreadCount
 Wscript.Echo "Page File Size: " & objProcess.PageFileUsage
 Wscript.Echo "Page Faults: " & objProcess.PageFaults
 Wscript.Echo "Working Set Size: " & objProcess.WorkingSetSize
end if
Next

'Material https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-process
end sub