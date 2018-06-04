'HASH: 3DE4FB1EEEEA47FDAE06521EBC724EB2
 
 
Public Sub EXPORTAR_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_WIDGETSEXPORTADOR" 
	form.Show 
End Sub 
 
Public Sub IMPORTAR_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_WIDGETSIMPORTADOR" 
	form.Show 
End Sub 
