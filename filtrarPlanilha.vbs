Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Users\wellington.fonseca\Desktop\teste.xlsx")
Set objWorksheet = objWorkbook.Worksheets("Planilha1")

objExcel.Selection.AutoFilter
objExcel.ActiveSheet.Range("A1:Z23").AutoFilter 1,"01.02.2022"
