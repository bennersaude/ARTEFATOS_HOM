'HASH: D65D9006316FBF415689783E6F675B99
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT TABTIPOLIMITE,VLRLIMITE ,QTDLIMITE")
  SQL.Add("FROM SAM_CONTRATO_LIMITACAO")
  SQL.Add("WHERE HANDLE = :HCONTRATOLIMITACAO")
  SQL.ParamByName("HCONTRATOLIMITACAO").Value = CurrentQuery.FieldByName("CONTRATOLIMITACAO").AsInteger
  SQL.Active = True


  If SQL.FieldByName("TABTIPOLIMITE").AsInteger = 1 Then ' Por quantidade
    If CurrentQuery.FieldByName("QUANTIDADE").AsInteger > SQL.FieldByName("QTDLIMITE").AsInteger Then
      bsShowMessage("A quantidade não pode ser maior do que a configurada na Limitação do contrato", "E")
      Set SQL = Nothing
      CanContinue = False
      Exit Sub
    End If
  Else 'por valor
    If CurrentQuery.FieldByName("QUANTIDADE").AsInteger > SQL.FieldByName("VLRLIMITE").AsInteger Then
      bsShowMessage("O Valor não pode ser maior do que a configurada na Limitação Do CONTRATO", "E")
      Set SQL = Nothing
      CanContinue = False
      Exit Sub
    End If

  End If


End Sub

