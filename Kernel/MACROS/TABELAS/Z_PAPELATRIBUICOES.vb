'HASH: 8784180FED4707D6C65B96757E2850DC
 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  ' Exige o preenchimento do grupo ou do usuário 
  If (CurrentQuery.FieldByName("GRUPO").IsNull And CurrentQuery.FieldByName("USUARIO").IsNull) Then 
    CanContinue = False 
    If (VisibleMode) Then 
      MsgBox("Informe o grupo ou o usuário") 
    Else 
      CancelDescription = "Informe o grupo ou o usuário" 
    End If 
    Exit Sub 
  End If 
 
  ' Atribui o tipo de objeto (AssigneeType) 
  If (CurrentQuery.FieldByName("USUARIO").IsNull) Then 
    CurrentQuery.FieldByName("TIPOOBJETO").AsInteger = 1 
  Else 
    CurrentQuery.FieldByName("TIPOOBJETO").AsInteger = 2 
    CurrentQuery.FieldByName("GRUPO").Clear  ' Se tem usuário não deve ter grupo 
  End If 
 
  ' Validação para tratar índice único 
  Dim query As BPesquisa 
  Set query = NewQuery 
    query.Add("SELECT COUNT(HANDLE) COUNT FROM  Z_PAPELATRIBUICOES WHERE HANDLE <> :HANDLE AND PAPEL = :PAPEL AND GRUPO = :GRUPO AND USUARIO = :USUARIO") 
    query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
    query.ParamByName("PAPEL").AsInteger = CurrentQuery.FieldByName("PAPEL").AsInteger 
 
  If CurrentQuery.FieldByName("GRUPO").IsNull Then 
    query.ParamByName("GRUPO").DataType = ftInteger 
    query.ParamByName("GRUPO").Clear 
  Else 
    query.ParamByName("GRUPO").AsInteger = CurrentQuery.FieldByName("GRUPO").AsInteger 
  End If 
  If CurrentQuery.FieldByName("USUARIO").IsNull Then 
    query.ParamByName("USUARIO").DataType = ftInteger 
    query.ParamByName("USUARIO").Clear 
  Else 
    query.ParamByName("USUARIO").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger 
  End If 
    query.Active = True 
    If query.FieldByName("COUNT").AsInteger > 0 Then 
      CanContinue = False 
      If (VisibleMode) Then 
        MsgBox("Esta atribuição já existe") 
       Else 
        CancelDescription = "Esta atribuição já existe" 
      End If 
  End If 
  query.Active = False 
    Set query = Nothing 
 
End Sub 
