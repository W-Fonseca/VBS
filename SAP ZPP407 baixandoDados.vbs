exportSAP

sub exportSAP()

dim Lista
set Lista = CreateObject("System.Collections.ArrayList")
Lista.add "C:\Users\wellington.fonseca\Documents\"
Lista.add "wellbot"
Lista.add ".xlsx"
  
  on error resume next 'executa sem apurar erros 
CreateObject("Scripting.FileSystemObject").DeleteFile(Lista(0) & Lista(1) & Lista(2)) 'tenta excluir o arquivo / para caso o mesmo ja exista, ai não aparece a pergunta se deseja sobrescrever
CreateObject("Scripting.FileSystemObject").DeleteFile("C:\export.XML")
CreateObject("Scripting.FileSystemObject").DeleteFile("C:\Export.xlsx")
  on error goto 0 'habilita novamente o apurador de erros  
  
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

dim ExcelM
set ExcelM = CreateObject("Excel.Application")
ExcelM.Visible = True

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/radRB_OTHERS").setFocus
session.findById("wnd[1]/usr/radRB_OTHERS").select
session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus
session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "04"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "export.XML"
session.findById("wnd[1]/tbar[0]/btn[0]").press

WScript.Sleep(10000)

CreateObject("WScript.Shell").Run "taskkill /im EXCEL.exe", , True


dim ExcelMHTML
set ExcelMHTML = CreateObject("Excel.Application")
ExcelMHTML.Visible = True
Set objWorkbookMHTML = ExcelMHTML.Workbooks.Open("C:\export.XML")
Set objWorksheetMHTML = objWorkbookMHTML.Worksheets("Sheet1")
'ExcelMHTML.ActiveSheet.Copy
'objWorkbookMHTML.SaveAs Lista(0) & Lista(1) & Lista(2)

ExcelMHTML.Range("A:D").Select
ExcelMHTML.Selection.Copy

dim ExcelApp
set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible = True
Set objWorkbook = ExcelApp.Workbooks.add
Set objWorksheet = objWorkbook.Worksheets(1)
objWorksheet.Range("a1").Select
ExcelApp.ActiveSheet.Paste

      'Lista(0)LocalSalvamento
      'Lista(1)NomeArquivo
      'Lista(2)TipoArquivo 
objWorkbook.SaveAs Lista(0) & Lista(1) & Lista(2) 
objWorkbook.Close 'fecha o excel

on error resume next
objWorkbookMHTML.Application.DisplayAlerts = False 
objWorkbookMHTML.Close
objWorkbookMHTML.Application.DisplayAlerts = True
CreateObject("Scripting.FileSystemObject").DeleteFile("C:\export.XML")
on error goto 0

CreateObject("WScript.Shell").Run "taskkill /im EXCEL.exe", , True

end sub
