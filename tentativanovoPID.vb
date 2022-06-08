dim meuID = System.Diagnostics.Process.GetCurrentProcess().SessionID
Dim objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Dim colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chrome.exe'")
dim contagem = 0
For Each objProcess In colProcesses
if objProcess.SessionID = meuID then
contagem = contagem + 1 
if contagem = 1 ' 
PIDGoogle = objProcess.ProcessID
exit for
end if
end if
Next
