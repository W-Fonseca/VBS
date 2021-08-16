copiarDados("Export-05-08-2021")
sub copiarDados(NomeArquivo)

CreateObject("Scripting.FileSystemObject").DeleteFile("\\gowacd01\Grupos\RPA\DEV\ExcelenciaOperacional\R014\ArquivoZPP407\" & NomeArquivo & ".xlsx")

Set oExcel = CreateObject("Excel.Application")  
oExcel.Visible = true  
Set oBook = oExcel.Workbooks.Add
         
Set oSheet = oBook.Worksheets(1)    
 
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

Linha = 0
Contagem = 0
For i = 1 To 1000
Linha = Linha + 1

'msgbox session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").GetCellValue(Contagem,"AUFNR") > ""
on error resume next
if session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").GetCellValue(Contagem,"AUFNR") < "" then
exit For
end if
on error goto 0
session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").firstVisibleRow = Contagem
oSheet.Range("A" & Linha).Value = session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").GetCellValue(Contagem,"AUFNR")
oSheet.Range("B" & Linha).Value = session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").GetCellValue(Contagem,"ARBPL")
oSheet.Range("C" & Linha).Value = "'" & session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").GetCellValue(Contagem,"VORNR") 
oSheet.Range("D" & Linha).Value = session.findById("wnd[0]/usr/cntlBCALV_GRID_DEMO_0100_CONT1/shellcont/shell").GetCellValue(Contagem,"LTXA1")

    
Contagem = Contagem + 1
next
oBook.SaveAs "\\gowacd01\Grupos\RPA\DEV\ExcelenciaOperacional\R014\ArquivoZPP407\" & NomeArquivo & ".xlsx"
oBook.Close

CreateObject("WScript.Shell").Run "taskkill /im EXCEL.exe", , True
End sub
