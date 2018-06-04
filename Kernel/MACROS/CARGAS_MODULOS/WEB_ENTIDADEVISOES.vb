'HASH: B5A969A48C515EAC1266D7B667B9860D
 
 
Public Sub EXPORTARVISOES_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_ENTIDADEVISOESEXPORTADOR" 
	form.Show 
End Sub 
 
Public Sub IMPORTARVISOES_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_ENTIDADEVISOESIMPORTADOR" 
	form.Show 
 
End Sub 
