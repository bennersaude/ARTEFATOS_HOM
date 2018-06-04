'HASH: 30C1D5C89FF82ED628F09547363F1BEB
 
Public Sub EDITARATRIBUTOS_OnClick() 
 
  Dim query As BPesquisa 
  Dim virtualTable As String 
 
  Set query = NewQuery 
  query.Text = "SELECT TABELAVIRTUAL, NOME FROM W_WIDGETS WHERE HANDLE = :PARAMETRO" 
  query.ParamByName("PARAMETRO").AsInteger = CurrentQuery.FieldByName("WIDGET").AsInteger 
  query.Active = True 
 
  If query.FieldByName("TABELAVIRTUAL").AsInteger > 0 Then 
    virtualTable = NewMemoryTableByHandle(query.FieldByName("TABELAVIRTUAL").AsInteger).Nome 
 
    Dim form As CSVirtualForm 
    Set form = NewVirtualForm 
 
    form.TableName = virtualTable 
    form.Caption = "Atributos do widget " + query.FieldByName("NOME").AsString 
    form.Width = 600 
 
    form.TransitoryVars("Xml") = CurrentQuery.FieldByName("ATRIBUTOS").AsString 
 
 
    If form.Show = 0 Then 
 
    If CurrentQuery.InReading Then 
      CurrentQuery.Edit 
    End If 
 
      CurrentQuery.FieldByName("ATRIBUTOS").AsString = form.TransitoryVars("Xml") 
  End If 
 
  Else 
    MsgBox("No Widget não foi informada a tabela virtual para editar os atributos.") 
  End If 
 
End Sub 
 
Public Sub NOME_OnExit() 
  If CurrentQuery.FieldByName("CODIGO").AsString = "" And CurrentQuery.State = 3 Then 
    CurrentQuery.FieldByName("CODIGO").AsString = CurrentQuery.FieldByName("NOME").AsString 
  End If 
End Sub 
