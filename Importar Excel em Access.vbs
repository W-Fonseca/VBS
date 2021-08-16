Sub CarregarAccess(LISTA)

dim accessApp
set accessApp = createObject("Access.Application")
accessApp.visible = true
'accessApp.DoCmd.SetWarnings False

  'LISTA(0)    PathAccess = "C:\Users\marcelo.fonseca\Desktop\SAP Cogtive teste.accdb"
  'LISTA(1) PathExcel = "C:\Users\marcelo.fonseca\Desktop\Apontamentos Cogtive - P35.xlsx"
  'LISTA(2) TableAccess = "BaseCogtive"

accessApp.OpenCurrentDataBase(LISTA(0))
accessApp.CurrentDb.Execute "DELETE * FROM " & LISTA(2)
AccessApp.DoCmd.TransferSpreadsheet acImport, 10, LISTA(2), (LISTA(1)), True

End Sub
