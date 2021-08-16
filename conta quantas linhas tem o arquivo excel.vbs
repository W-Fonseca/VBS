Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Users\wellington.fonseca\Desktop\Apontamento Cogtive - P19 - SOLIDOS.xlsx")
Set objWorksheet = objWorkbook.Worksheets("RELATORIO DE APONTAMENTOS")

Const xlByRows = 1
Const xlPrevious = 2


Set objExcel = GetObject(,"Excel.Application")
LastRow = objExcel.ActiveSheet.Cells.Find("*", , , , xlByRows, xlPrevious).Row


MsgBox LastRow






