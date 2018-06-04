'HASH: A30D1858669CFD3CCAF1A5791563CD60
 
Public Sub AUTOGERADA_OnChange() 
  If CurrentQuery.FieldByName("AUTOGERADA").AsBoolean Then 
	URLAJUDA.Text = "Exemplo de URL:  ~/rh/a/Paises/Grid.aspx" 
  Else 
	URLAJUDA.Text = "Exemplo de URL:  ~/rh/e/Funcionario/Admissao.aspx" 
  End If 
End Sub 
 
Public Sub BOTAOGERARPAGINA_OnClick() 
	Dim form As CSVirtualForm 
	Set form = NewVirtualForm 
	form.TableName = "W_PAGEGENERATOR" 
	form.Caption = "Gerar páginas" 
	'Passa o handle da página atual, para gerar apenas uma página 
	form.TransitoryVars("PAGE") = CurrentQuery.FieldByName("HANDLE").AsInteger 
	form.Show 
	Set form = Nothing 
End Sub 
 
Public Sub NOME_OnExit() 
  If CurrentQuery.FieldByName("CODIGO").AsString = "" And CurrentQuery.State = 3 Then 
    CurrentQuery.FieldByName("CODIGO").AsString = CurrentQuery.FieldByName("NOME").AsString 
  End If 
End Sub 
