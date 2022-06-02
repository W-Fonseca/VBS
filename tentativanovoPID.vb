'System.Diagnostics.Process
'dim remoteAll = Process.GetCurrentProcess()
'ProcessStartInfo
'on error goto 0
'dim remoteAll = System.Diagnostics.Process.GetProcesses(Environment.MachineName)
'dim remoteAll = Process.GetProcessesByName("Chrome")
'for each remot in remoteAll
'Environment.MachineName
'Dim WshShell = CreateObject("WScript.Shell")

'dim strSessionName = WshShell.ExpandEnvironmentStrings("%QUERY  PROCESS WELLINGTONFONSECA%")
'dim strSessionName = WshShell.ExpandEnvironmentStrings("SELECT sessionid")
'System.Windows.Forms.Messagebox.Show("oi")
'System.Windows.Forms.Messagebox.Show(strSessionName)
'System.Windows.Forms.Messagebox.Show(remot.ID)
'System.Windows.Forms.Messagebox.Show(remot.SessionID)
'System.Windows.Forms.Messagebox.Show(remot.SessionName)
'System.Windows.Forms.Messagebox.Show(remot.ProcessName)
'System.Windows.Forms.Messagebox.Show(remot.Domain)
'System.Windows.Forms.Messagebox.Show(Environment.UserInteractive)
'System.Windows.Forms.Messagebox.Show(Environment.UserIP)
'System.Windows.Forms.Messagebox.Show(System.Security.Principal.WindowsIdentity.GetCurrent().Number)
'System.Windows.Forms.Messagebox.Show(System.Security.Principal.WindowsIdentity.GetCurrent().Token)
'System.Windows.Forms.Messagebox.Show(System.Security.Principal.SecurityIdentifier)
'System.Windows.Forms.Messagebox.Show(remot.System.Diagnostics.ProcessStartInfo.UserName)
'System.Windows.Forms.Messagebox.Show(Environment.CommandLine("query user"))
'ID = id do processo
'SessionID = id do usuario
'ProcessName = nome do processo
'System.Windows.Forms.Messagebox.Show(WindowsIdentity.GetCurrent)

'next

'System.Windows.Forms.Messagebox.Show(WTSQuerySessionInformation.SessionID)
'System.Windows.Forms.Messagebox.Show(sessionId)
        'Dim ff As System.Security.Principal.WindowsIdentity
       ' ff = System.Security.Principal.WindowsIdentity.GetCurrent
'Dim ee = ff.User.ToString()
'dim AccSid as string = ff.User.AccountDomainSid.tostring
'System.Windows.Forms.Messagebox.Show(ee)
'on error resume next
'dim strNameOfUser
'Dim objWMIService = GetObject("winmgmts:\\.\root\cimv2") 'declara a variavel o local da query procurada
'Dim colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chrome.exe'")
'For Each objProcess In colProcesses 
'    colProperties = processes.GetOwner(user)
'System.Windows.Forms.Messagebox.Show(objProcess.GetOwner())
'if objProcess.GetOwner() <> 0 then
'System.Windows.Forms.Messagebox.Show(objProcess.Name)
'System.Windows.Forms.Messagebox.Show(objProcess.ProcessID)
'System.Windows.Forms.Messagebox.Show(objProcess.ParentProcessId)
'System.Windows.Forms.Messagebox.Show(objProcess.CSName)
'System.Windows.Forms.Messagebox.Show(objProcess.Caption)
'System.Windows.Forms.Messagebox.Show(objProcess.Handle)
'System.Windows.Forms.Messagebox.Show(objProcess.SessionId)
'System.Windows.Forms.Messagebox.Show(objProcess.GetOwner())
'end if
'next

strComputer = "."
dim colProcesses = GetObject("winmgmts:" & _
   "{impersonationLevel=impersonate}!\\" & strComputer & _
   "\root\cimv2").ExecQuery("Select * from Win32_Process")

For Each objProcess in colProcesses

    dim Returno = objProcess.GetOwner(strNameOfUser)
    If Returno <> 0 Then
        System.Windows.Forms.Messagebox.Show( "Could not get owner info for process " & objProcess.Name & VBNewLine & "Error = " )
    Else 
        System.Windows.Forms.Messagebox.Show("Process " & objProcess.Name & " is owned by " & "\" & strNameOfUser & ".")
    End If
Next
