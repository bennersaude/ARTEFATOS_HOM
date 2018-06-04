'HASH: 7E17141E89F1F95810317A56FB6B5A64
 
 
Public Sub EXPORTARMENUS_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_MENUSEXPORTADOR" 
	form.Show 
End Sub 
 
Public Sub IMPORTARMENUS_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_MENUSIMPORTADOR" 
	form.Show 
End Sub 
