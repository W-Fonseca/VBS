copiarPDF()

Sub copiarPDF() 'mencionar a pasta dos arquivos PDF

set shellApp = CreateObject("Shell.Application")
set folder = shellApp.BrowseForFolder(0, "Selecione a Pasta", 0, myStartFolder)
pasta = folder.Self.Path


'OBS: antes de executar esse script, salve todos os arquivos excel e word, pois ele mata todos esses processos antes de executar.

Dim oShell : Set oShell = CreateObject("WScript.Shell") 'cria um objeto powershell para executar ações no mesmo.

on error resume next 'desliga apurador de erros
oShell.Run "taskkill /im EXCEL.EXE /F", , True 'mata todos os aplicativos excel
oShell.Run "taskkill /im WINWORD.EXE /F", , True 'mata todos os aplicativos word
on error goto 0 'liga o apurador de erros

Set oFSO = CreateObject("Scripting.FileSystemObject") 'cria um objeto de sistema de arquivos

For Each nomearquivo In oFSO.GetFolder(pasta).Files
on error resume next 'desliga apurador de erros
arquivodelete = replace(nomearquivo,".pdf",".xlsx")
oFSO.DeleteFile(arquivodelete)
on error goto 0 'liga o apurador de erros
next

For Each nomearquivo In oFSO.GetFolder(pasta).Files 'para cada arquivo na pasta


Dim doc 'cria variavel doc
Dim wa  'cria variavel wa
Set wa = CreateObject("word.application") 'variavel wa é um aplicativo word
'wa.visible = True
Set doc = wa.Documents.Open(nomearquivo.path,False) 'variavel doc é a variavel wa aberta
pasta = replace(nomearquivo,doc.name,"") 'variavel pasta é o caminho da pasta
wa.Selection.WholeStory 'rola todo o arquivo word selecionando, é igual apertar um CTRL + T
wa.Selection.Copy 'copia toda a seleção

nome = Replace(doc.name,".pdf","") 'variavel nome é igual o nome do arquivo word sem .pdf
dim ExcelApp 'declara uma nova variavel
set ExcelApp = CreateObject("Excel.Application") 'digo que essa variavel ExcellApp é um 'aplicativo Excel

dim wb 'declara uma nova variavel
set wb = ExcelApp.workbooks.add 'digo que a variavel wb é uma nova aba do excel

ExcelApp.WorkSheets(1).Range("A1").Select 'digo que quero selecionar a celula A1 da primeira aba

ExcelApp.ActiveSheet.Paste 'cola na aba
doc.Close 'fecha o word
wa.Quit 'fecha a aplicação word

wb.SaveAs(pasta & nome & ".xlsx")

Next 'passa para o proximo

on error resume next 'desliga apurador de erros
oShell.Run "taskkill /im EXCEL.EXE /F", , True 'mata todos os aplicativos excel
oShell.Run "taskkill /im WINWORD.EXE /F", , True 'mata todos os aplicativos word
on error goto 0 'liga o apurador de erros

End Sub
