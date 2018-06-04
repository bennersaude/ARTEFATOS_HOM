'HASH: B2608C0CAD5FC2FA40323F7B6A3C819F
 
Public Sub GERARPAGINAS_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_PAGEGENERATOR" 
	form.Caption = "Gerar páginas" 
	form.Show 
	Set form = Nothing 
End Sub 
 
Public Sub IMPORTARPAGINAS_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_PAGINASIMPORTADOR" 
	form.Show 
	Set form = Nothing 
End Sub 
 
Public Sub EXPORTARPAGINAS_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_PAGINASEXPORTADOR" 
	form.Show 
	Set form = Nothing 
End Sub 
