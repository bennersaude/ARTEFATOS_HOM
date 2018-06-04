'HASH: E6C75D21E9158E2839BBF5E11C73307F
Dim Sql As Object, SqlContador As Object

Dim Contador As Integer



'Public Sub ULTIMONIVEL_OnChange()
'  CurrentyQuery("NIVELSUPERIOR") = 0

'End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Set Sql = NewQuery
  Sql.Add "SELECT CARGO FROM ST_AMBIENTECARGOS WHERE AMBIENTE = " + CurrentQuery.FieldByName("HANDLE").AsString
  Sql.Active = True
  If Not (Sql.EOF) Then
    Set SqlContador = NewQuery
    Contador = 0
    While Not (Sql.EOF)
      SqlContador.Add "SELECT CARGO, COUNT (*) CONTA FROM MS_DADOSEMPRESA WHERE CARGO = " + Sql.FieldByName("CARGO").AsInteger + "GROUP BY CARGO "
      SqlContador.Active = True
      If Not (SqlContador.EOF) Then
        Contador = Contador + SqlContador.FieldByName("CONTA").AsInteger
      End If
      Sql.Next

    Wend
    CurrentQuery.FieldByName("NUMEROFUNCIONARIOS").AsInteger = Contador
  End If

  '	Set Sql = NewQuery
  '    Sql.Add "SELECT CARGO FROM ST_AMBIENTECARGOS WHERE AMBIENTE = " +CurrentQuery.FieldByName("HANDLE").AsString
  '    Sql.Active = True
  '    If Not (Sql.EOF) Then
  '    	Set SqlContador = NewQuery
  '    	Contador = 0
  '    	While Not (Sql.EOF)
  '    		 SqlContador.Add "SELECT CARGO, COUNT (*) CONTA FROM DO_FUNCIONARIOS WHERE CARGO = " + Sql.FieldByName("CARGO").AsInteger + " AND SITUACAO = 2 GROUP BY CARGO "
  '    		 SqlContador.Active = True
  '    		 If Not (SqlContador.EOF) Then
  '    		 	Contador = Contador + SqlContador.FieldByName("CONTA").AsInteger
  '    		 End If
  '    		 Sql.Next
  '
  '		Wend
  '		CurrentQuery.FieldByName("NUMEROFUNCIONARIOS").AsInteger = Contador
  '    End If
End Sub


