'HASH: EB31443E34B7C7E002DFB827CD9F4DF9
'Macro: SFN_AGENCIA
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterDelete()
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT TABINTEGRACAOAGENCIASBANCARIAS FROM ADM_PARAMINTEGRACAOCORPBENNER")
  SQL.Active = True

  If (SQL.FieldByName("TABINTEGRACAOAGENCIASBANCARIAS").AsInteger = 1) Then
    Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
    Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

    TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SFN_AGENCIA")
    TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "Z")

    TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_AfterPost()
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT TABINTEGRACAOAGENCIASBANCARIAS FROM ADM_PARAMINTEGRACAOCORPBENNER")
  SQL.Active = True

  If (SQL.FieldByName("TABINTEGRACAOAGENCIASBANCARIAS").AsInteger = 1) Then
    Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
    Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

    TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SFN_AGENCIA")
    TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

    TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("CNPJ").IsNull Then
    If Not IsValidCGC(CurrentQuery.FieldByName("CNPJ").AsString) Then
      CanContinue = False
      bsShowMessage("CNPJ inválido", "E")
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "A" Then
    Dim sqlAgencia As Object
    Set sqlAgencia = NewQuery
    sqlAgencia.Add("SELECT HANDLE            ")
    sqlAgencia.Add("  FROM SFN_AGENCIA       ")
    sqlAgencia.Add(" WHERE BANCO   = :BANCO  ")
    sqlAgencia.Add("   AND AGENCIA = :AGENCIA")
    sqlAgencia.Add("   AND HANDLE <> :HANDLE ")
    sqlAgencia.Add("   AND SITUACAO = 'A'    ")
    sqlAgencia.ParamByName("BANCO").AsInteger = CurrentQuery.FieldByName("BANCO").AsInteger
    sqlAgencia.ParamByName("AGENCIA").AsString = CurrentQuery.FieldByName("AGENCIA").AsString
    sqlAgencia.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sqlAgencia.Active = True

    If Not sqlAgencia.EOF Then
      CanContinue = False
      bsShowMessage("Agência já cadastrada para esse banco", "E")
      Set sqlAgencia = Nothing
      Exit Sub
    End If
    Set sqlAgencia = Nothing
  End If

  Dim sqlBanco As Object
  Set sqlBanco = NewQuery

  sqlBanco.Active = False
  sqlBanco.Add("SELECT PERMITEAGENCIASEMDV FROM SFN_BANCO WHERE HANDLE = :BANCO")
  sqlBanco.ParamByName("BANCO").AsInteger = CurrentQuery.FieldByName("BANCO").AsInteger
  sqlBanco.Active = True

  If sqlBanco.FieldByName("PERMITEAGENCIASEMDV").AsString = "N" And CurrentQuery.FieldByName("DV").AsString = "" Then
    CanContinue = False
    bsShowMessage("Necessário digitar o DV para a agência.","E")
    Set sqlBanco = Nothing
    Exit Sub
  End If

  Set sqlBanco = Nothing

End Sub
