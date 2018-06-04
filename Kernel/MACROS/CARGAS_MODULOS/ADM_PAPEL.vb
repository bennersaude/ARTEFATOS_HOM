'HASH: 0BEBAB384ED6065A6647CA77E7FF2EDC
 
Public Sub EXPORTAUTHORIZATION_OnClick() 
 Dim form As CSVirtualForm 
 Set form = NewVirtualForm 
 form.TableName = "Z_ROLEDEFINITIONEXPORTER" 
	form.Show 
End Sub 
 
Public Sub IMPORTAUTHORIZATION_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "Z_ROLEDEFINITIONIMPORTER" 
	form.Show 
End Sub 
