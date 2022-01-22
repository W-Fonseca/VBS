
Try
   If Not File.Exists("C:\Temp\vba.bas") Then
      Throw New ApplicationException("The specified file at does not exist.")
   Else
	   File.Delete("C:\Temp\vba.bas")
   End If
Catch e As Exception
End Try


Dim FSO = CreateObject("Scripting.FileSystemObject")
Dim OutPutFile = FSO.OpenTextFile("C:\Temp\vba.bas" ,8 , True)
OutPutFile.WriteLine("Attribute VB_Name = " & """VBA199"""  & Environment.NewLine & Script)
'set objExcel = GetObject("C:\Users\wellingtonfonseca\Desktop\wellington.xlsx").Application
Dim objExcel = GetObject(WorkbookPath).Application

'CreateObject("Wscript.Shell").AppActivate(objExcel.Caption)

'objExcel.Visible = true

objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Temp\vba.bas")
objExcel.Application.Run (NameScriptRun)
objExcel.VBE.ActiveVBProject.VBComponents.Remove (objExcel.VBE.ActiveVBProject.VBComponents.Item("VBA199"))

'Catch ex As Exception
'Dim objExcel = CreateObject("Excel.Application")
'Dim objWorkbook = objExcel.Workbooks.Open(WorkbookPath)
'objExcel.Visible = true
'objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Temp\vba.bas")
'objExcel.Application.Run (NameScriptRun)
'objExcel.VBE.ActiveVBProject.VBComponents.Remove(objExcel.VBE.ActiveVBProject.VBComponents.Item("VBA199"))
'End Try

Try
   If Not File.Exists("C:\Temp\vba.bas") Then
      Throw New ApplicationException("The specified file at does not exist.")
   Else
	   File.Delete("C:\Temp\vba.bas")
	   
   End If
Catch e As Exception
End Try

'CreateObject("Scripting.FileSystemObject").DeleteFile("C:\Temp\vba.bas")

