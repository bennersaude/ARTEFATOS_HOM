'HASH: 4820C03CF2305A7AA3DD11CA487D8006
Sub Main
  Dim Sql As Object
  Set Sql = NewQuery

  Dim Del As Object
  Set Del = NewQuery

  Sql.Add("SELECT T.NOME ")
  Sql.Add("  FROM Z_TABELAS T")
  Sql.Add("  JOIN Z_CAMPOS  C ON C.TABELA = T.HANDLE")
  Sql.Add(" WHERE T.NOME LIKE 'TMP%'")
  Sql.Active = True

  While Not Sql.EOF

    On Error GoTo DELETE 'Go To para caso não tenha permissão para dar truncate table (usuário do banco de dados)
      Del.Clear
      Del.Add("TRUNCATE TABLE "+ Sql.FieldByName("NOME").AsString)
      Del.ExecSQL

    DELETE:

      On Error GoTo FIM 'Go To para caso de termos a tabela na Z_TABELAS e não termos ela fisicamente

      Del.Clear
      Del.Add("DELETE FROM "+ Sql.FieldByName("NOME").AsString)
      Del.ExecSQL

    FIM:

    Sql.Next
  Wend

  Set Sql = Nothing
  Set Del = Nothing

End Sub
