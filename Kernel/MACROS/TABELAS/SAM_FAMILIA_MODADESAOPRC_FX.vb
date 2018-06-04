'HASH: 60B16AF540FD9470953D5B37C526B428
'Macro: SAM_FAMILIA_MODADESAOPRC_FX
'Mauricio -19/10/2000
'#Uses "*bsShowMessage"

'Dim CheckContinua As String
Dim InclusaoForaDoPadrao As Boolean
Dim LogInclusaoForaDoPadrao As String

Public Function CheckCodigoTabelaPrcContrato As Boolean

  Dim SQL As Object

  CheckCodigoTabelaPrcContrato = True

  Set SQL = NewQuery

  SQL.Add("SELECT CODIGOTABELAPRC FROM SAM_CONTRATO_MODADESAO c, ")
  SQL.Add("SAM_CONTRATO_MODADESAOPRC d, SAM_CONTRATO_MODADESAOPRC_FX e ")
  SQL.Add("where c.CONTRATOMODULO = :HCONTRATOMODULO       And")
  SQL.Add("      c.HANDLE         = d.CONTRATOMODADESAO    And")
  SQL.Add("      d.HANDLE         = e.CONTRATOMODADESAOPRC And")
  SQL.Add("      e.CODIGOTABELAPRC = :CODIGOTABELAPRC")

  SQL.ParamByName("HCONTRATOMODULO").Value = CurrentQuery.FieldByName("MODULO").AsInteger
  SQL.ParamByName("CODIGOTABELAPRC").Value = CurrentQuery.FieldByName("CODIGOTABELAPRC").AsString
  SQL.RequestLive = False
  SQL.Active = True
  If SQL.EOF Then
	  If bsShowMessage("Código da tabela de preço não encontrada no CONTRATO. " + (Chr(13)) + _
              " Continua?", "Q") = vbYes Then


  	   Set SQL = Nothing
       CheckCodigoTabelaPrcContrato = False
	  Else

	      CheckCodigoTabelaPrcContrato = True
      	  Set SQL = Nothing
  	  End If

 End If

End Function


Public Function CheckCoParticipacao As Boolean

  Dim SQL As Object
  Dim vlContrato As Long

  CheckCoParticipacao = True

  Set SQL = NewQuery

  ' vlContrato = _BSistema.RecordHandleOfTable(SAM_CONTRATO)
  vlContrato = RecordHandleOfTable("SAM_CONTRATO")


  SQL.Add("Select a.handle from sam_contrato_mod a,         ")
  SQL.Add("              sam_contrato_modadesao b,   ")
  SQL.Add("              sam_contrato_modadesaoprc c ")
  SQL.Add("where a.contrato = :CONTRATO          And ")
  SQL.Add("      a.handle   = b.contratomodulo   And ")
  SQL.Add("      b.handle   = c.contratomodadesao   ")

  SQL.ParamByName("CONTRATO").Value = vlContrato 'CurrentQuery.FieldByName("CONTRATO").AsInteger
  SQL.RequestLive = False
  SQL.Active = True
  If Not SQL.EOF Then
    CheckCoParticipacao = True
    Set SQL = Nothing
    Exit Function
  End If

  Set SQL = Nothing

  CheckCoParticipacao = False


End Function

Public Sub CONTRATOTPDEP_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_CONTRATO_TPDEP|SAM_TIPODEPENDENTE[SAM_CONTRATO_TPDEP.TIPODEPENDENTE = SAM_TIPODEPENDENTE.HANDLE]", "DESCRICAO", 1, "Descrição", "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO")), "Procura por Tipo dependente", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOTPDEP").Value = handlexx
  End If
  Set Procura = Nothing

End Sub

Public Sub TABLE_AfterDelete()
  Dim vPrimeiraCompetencia As Date
  Dim vUltimaCompetencia As Date
  Dim vCompetencia As Date
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT COMPETENCIA")
  SQL.Add("FROM SAM_FAMILIA_MODADESAOPRC")
  SQL.Add("WHERE HANDLE = :HFAMILIAMODADESAOPRC")
  SQL.ParamByName("HFAMILIAMODADESAOPRC").Value = CurrentQuery.FieldByName("FAMILIAMODADESAOPRC").AsInteger
  SQL.Active = True
  vCompetencia = CurrentQuery.FieldByName("FAMILIAMODADESAOPRC").AsInteger

  'SMS 10255 ->DANIELA
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT MODULO FROM SAM_FAMILIA_MOD WHERE HANDLE = :HFAMILIAMOD")
  q1.ParamByName("HFAMILIAMOD").AsInteger = RecordHandleOfTable("SAM_FAMILIA_MOD")
  q1.Active = False
  q1.Active = True
  'Fim Daniela

  SQL.Clear
  SQL.Add("SELECT MIN(COMPETENCIA) PRIMEIRACOMPETENCIA, MAX(COMPETENCIA) ULTIMACOMPETENCIA")
  SQL.Add("FROM SFN_FATURA_LANC_MOD A, SAM_BENEFICIARIO_MOD BM")
  SQL.Add("WHERE BM.MODULO = :HCONTRATOMOD")
  SQL.Add("  AND (   BM.DATACANCELAMENTO IS NULL")
  SQL.Add("       OR BM.DATACANCELAMENTO >= :HOJE)")
  SQL.Add("  AND A.BENEFICIARIOMOD = BM.HANDLE")
  SQL.Add("  AND A.COMPETENCIA >= :COMPETENCIA")
  SQL.ParamByName("HCONTRATOMOD").Value = q1.FieldByName("MODULO").AsInteger 'RecordHandleOfTable("SAM_CONTRATO_MOD")
  SQL.ParamByName("HOJE").Value = ServerDate
  SQL.ParamByName("COMPETENCIA").Value = vCompetencia
  SQL.Active = True


  If Not SQL.FieldByName("PRIMEIRACOMPETENCIA").IsNull Then

    vPrimeiraCompetencia = SQL.FieldByName("PRIMEIRACOMPETENCIA").AsDateTime
    vUltimaCompetencia = SQL.FieldByName("ULTIMACOMPETENCIA").AsDateTime

    SQL.Clear
    SQL.Add("SELECT HANDLE")
    SQL.Add("FROM SAM_ROTINARECALCULOMENSALID")
    SQL.Add("WHERE SITUACAO = 'A'")
    SQL.Add("  AND TABRECALCULAR = 3")
    SQL.Add("  AND COMPETENCIAINICIAL = :COMPETENCIAINICIAL")
    SQL.Add("  AND COMPETENCIAFINAL = :COMPETENCIAFINAL")
    SQL.Add("  AND CONTRATO = :HCONTRATO")
    SQL.Add("  AND FAMILIAINICIAL = :HFAMILIA")
    SQL.Add("  AND FAMILIAFINAL = :HFAMILIA")
    SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
    SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
    SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
    SQL.ParamByName("HFAMILIA").Value = RecordHandleOfTable("SAM_FAMILIA")
    SQL.Active = True

    If SQL.EOF Then

      SQL.Clear
      SQL.Add("INSERT INTO SAM_ROTINARECALCULOMENSALID")
      SQL.Add("(HANDLE, CODIGO, DESCRICAO, DATAROTINA, TABRECALCULAR,")
      SQL.Add(" COMPETENCIAINICIAL, COMPETENCIAFINAL, CONTRATO, FAMILIAINICIAL, FAMILIAFINAL,")
      SQL.Add(" USUARIO, DATAINCLUSAO, SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO)")
      SQL.Add("VALUES")
      SQL.Add("(:HANDLE, :HANDLE, :DESCRICAO, :DATAROTINA, 3,")
      SQL.Add(" :COMPETENCIAINICIAL, :COMPETENCIAFINAL, :HCONTRATO, :HFAMILIA, :HFAMILIA,")
      SQL.Add(" :USUARIO, :DATAINCLUSAO, '1', '1')")

      SQL.ParamByName("HANDLE").Value = NewHandle("SAM_ROTINARECALCULOMENSALID")
      SQL.ParamByName("DATAROTINA").Value = ServerDate
      SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
      SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
      SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
      SQL.ParamByName("HFAMILIA").Value = RecordHandleOfTable("SAM_FAMILIA")
      SQL.ParamByName("USUARIO").Value = CurrentUser
      SQL.ParamByName("DATAINCLUSAO").Value = ServerDate
      SQL.ParamByName("DESCRICAO").Value = "Alteração nas faixas de preço do módulo"

      SQL.ExecSQL

    End If

  End If

  Set SQL = Nothing
  Set q1 = Nothing
End Sub

Public Sub TABLE_AfterPost()
  Dim vPrimeiraCompetencia As Date
  Dim vUltimaCompetencia As Date
  Dim vCompetencia As Date
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT COMPETENCIA")
  SQL.Add("FROM SAM_FAMILIA_MODADESAOPRC")
  SQL.Add("WHERE HANDLE = :HFAMILIAMODADESAOPRC")
  SQL.ParamByName("HFAMILIAMODADESAOPRC").Value = RecordHandleOfTable("SAM_FAMILIA_MODADESAOPRC")
  SQL.Active = True
  vCompetencia = CurrentQuery.FieldByName("FAMILIAMODADESAOPRC").AsInteger

  SQL.Clear
  SQL.Add("SELECT MIN(COMPETENCIA) PRIMEIRACOMPETENCIA, MAX(COMPETENCIA) ULTIMACOMPETENCIA")
  SQL.Add("FROM SFN_FATURA_LANC_MOD A, SAM_BENEFICIARIO_MOD BM")
  SQL.Add("WHERE BM.MODULO = :HCONTRATOMOD")
  SQL.Add("  AND (   BM.DATACANCELAMENTO IS NULL")
  SQL.Add("       OR BM.DATACANCELAMENTO >= :HOJE)")
  SQL.Add("  AND A.BENEFICIARIOMOD = BM.HANDLE")
  SQL.Add("  AND A.COMPETENCIA >= :COMPETENCIA")
  SQL.ParamByName("HCONTRATOMOD").Value = RecordHandleOfTable("SAM_CONTRATO_MOD")
  SQL.ParamByName("HOJE").Value = ServerDate
  SQL.ParamByName("COMPETENCIA").Value = vCompetencia
  SQL.Active = True

  vPrimeiraCompetencia = SQL.FieldByName("PRIMEIRACOMPETENCIA").AsDateTime
  vUltimaCompetencia = SQL.FieldByName("ULTIMACOMPETENCIA").AsDateTime

  If Not SQL.FieldByName("PRIMEIRACOMPETENCIA").IsNull Then
    SQL.Clear
    SQL.Add("SELECT HANDLE")
    SQL.Add("FROM SAM_ROTINARECALCULOMENSALID")
    SQL.Add("WHERE SITUACAO = 'A'")
    SQL.Add("  AND TABRECALCULAR = 3")
    SQL.Add("  AND COMPETENCIAINICIAL = :COMPETENCIAINICIAL")
    SQL.Add("  AND COMPETENCIAFINAL = :COMPETENCIAFINAL")
    SQL.Add("  AND CONTRATO = :HCONTRATO")
    SQL.Add("  AND FAMILIAINICIAL = :HFAMILIA")
    SQL.Add("  AND FAMILIAFINAL = :HFAMILIA")
    SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
    SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
    SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
    SQL.ParamByName("HFAMILIA").Value = RecordHandleOfTable("SAM_FAMILIA")
    SQL.Active = True

    If SQL.EOF Then

      SQL.Clear
      SQL.Add("INSERT INTO SAM_ROTINARECALCULOMENSALID")
      SQL.Add("(HANDLE, CODIGO, DESCRICAO, DATAROTINA, TABRECALCULAR,")
      SQL.Add(" COMPETENCIAINICIAL, COMPETENCIAFINAL, CONTRATO, FAMILIAINICIAL, FAMILIAFINAL,")
      SQL.Add(" USUARIO, DATAINCLUSAO, SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO)")
      SQL.Add("VALUES")
      SQL.Add("(:HANDLE, :HANDLE, :DESCRICAO, :DATAROTINA, 3,")
      SQL.Add(" :COMPETENCIAINICIAL, :COMPETENCIAFINAL, :HCONTRATO, :HFAMILIA, :HFAMILIA,")
      SQL.Add(" :USUARIO, :DATAINCLUSAO, '1', '1')")

      SQL.ParamByName("HANDLE").Value = NewHandle("SAM_ROTINARECALCULOMENSALID")
      SQL.ParamByName("DATAROTINA").Value = ServerDate
      SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
      SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
      SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
      SQL.ParamByName("HFAMILIA").Value = RecordHandleOfTable("SAM_FAMILIA")
      SQL.ParamByName("USUARIO").Value = CurrentUser
      SQL.ParamByName("DATAINCLUSAO").Value = ServerDate
      SQL.ParamByName("DESCRICAO").Value = "Alteração nas faixas de preço do módulo"

      SQL.ExecSQL

    End If

  End If

  Set SQL = Nothing

  If InclusaoForaDoPadrao Then
    WriteAudit("|", HandleOfTable("SAM_FAMILIA_MODADESAOPRC_FX"), CurrentQuery.FieldByName("HANDLE").Value, LogInclusaoForaDoPadrao)
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT HANDLE")
  SQL.Add("FROM SAM_FAMILIA_REPACTUACAO")
  SQL.Add("WHERE FAMILIAMODADESAOPRCFX = :HMODADESAOPRCFX")
  SQL.ParamByName("HMODADESAOPRCFX").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Existem faixas de repactuação. Exclusão não permitida", "E")
    CanContinue = False
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim SQL
  Set SQL = NewQuery
  SQL.Add("SELECT DATACANCELAMENTO FROM SAM_FAMILIA_MOD WHERE HANDLE = :HFAMILIAMOD")
  SQL.ParamByName("HFAMILIAMOD").Value = RecordHandleOfTable("SAM_FAMILIA_MOD")
  SQL.Active = True
  If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("Módulo cancelado não permite manutenções", "I")
    CurrentQuery.Cancel
    RefreshNodesWithTable("SAM_FAMILIA_MODADESAO_FX")
  End If
End Sub


Public Sub TABLE_Beforepost(CanContinue As Boolean)

  'SMS: 9579
  'Não há motivo para tal verificação
  'CheckContinua ="S"
  'If CheckCoParticipacao Then
  '  If CheckCodigoTabelaPrcContrato Then
  '      CanContinue =False
  '  Else
  '      CanContinue =True
  '  End If
  'End If

  'If CheckContinua ="N" Then
  '  CanContinue =False
  '  Exit Sub
  'End If

  Dim TipoCalculoPreco As String

  TipoCalculoPreco = VerificaPreco

  If TipoCalculoPreco <>"2" And _
      Not(CurrentQuery.FieldByName("CONTRATOTPDEP").IsNull)Then
    CanContinue = False
    bsShowMessage("A configuração do módulo NÃO permite que se informe o Tipo dependente", "E")
  End If

  If TipoCalculoPreco <>"3" And _
      Not(CurrentQuery.FieldByName("GRUPODEPENDENTE").IsNull)Then
    CanContinue = False
    bsShowMessage("A configuração do módulo NÃO permite que se informe o Grupo dependente", "E")
    CurrentQuery.FieldByName("GRUPODEPENDENTE").Clear
  End If

  If TipoCalculoPreco = "3" And _
                        CurrentQuery.FieldByName("GRUPODEPENDENTE").IsNull Then
    CanContinue = False
    bsShowMessage("A configuração do módulo exige que se informe o Grupo dependente", "E")
  End If

  InclusaoForaDoPadrao = False

  If CurrentQuery.State = 3 Then
    Dim SQL As Object
    Dim vCompetencia As Date
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT F.VALORMINIMO, F.VALORCUSTO")
    SQL.Add("FROM SAM_FAMILIA_MODADESAOPRC A, SAM_FAMILIA_MODADESAO B, SAM_FAMILIA_MOD C, SAM_CONTRATO_MOD D, SAM_MODULO_PRECO E, SAM_MODULO_PRECO_FX F")
    SQL.Add("WHERE A.HANDLE = :HFAMILIAMODADESAOPRC")
    SQL.Add("  AND B.HANDLE = A.FAMILIAMODADESAO")
    SQL.Add("  AND C.HANDLE = B.FAMILIAMODULO")
    SQL.Add("  AND D.HANDLE = C.MODULO")
    SQL.Add("  AND E.MODULO = D.MODULO")
    SQL.Add("  AND E.COMPETENCIAINICIAL <= A.COMPETENCIA")
    SQL.Add("  AND (E.COMPETENCIAFINAL IS NULL OR E.COMPETENCIAFINAL >= A.COMPETENCIA)")
    SQL.Add("  AND F.MODULOPRECO = E.HANDLE")
    SQL.Add("  AND F.IDADEMAXIMA = (SELECT MIN(IDADEMAXIMA)")
    SQL.Add("                       FROM SAM_MODULO_PRECO_FX")
    SQL.Add("                       WHERE MODULOPRECO = E.HANDLE")
    SQL.Add("                         AND IDADEMAXIMA >= :IDADEMAXIMA)")
    SQL.ParamByName("HFAMILIAMODADESAOPRC").Value = CurrentQuery.FieldByName("FAMILIAMODADESAOPRC").AsInteger
    SQL.ParamByName("IDADEMAXIMA").Value = CurrentQuery.FieldByName("IDADEMAXIMA").AsInteger
    SQL.Active = True

    If Not SQL.EOF Then
      If CurrentQuery.FieldByName("VALOR").AsFloat <SQL.FieldByName("VALORCUSTO").AsFloat Then
          bsShowMessage("O valor está abaixo do valor de custo padrão do módulo", "I" )
          InclusaoForaDoPadrao = True
          LogInclusaoForaDoPadrao = "Inclusão de faixa abaixo do valor de custo padrão" + Chr(13) + "Custo: " + SQL.FieldByName("VALORCUSTO").AsString + "  Valor: " + CurrentQuery.FieldByName("VALOR").AsString
      ElseIf CurrentQuery.FieldByName("VALOR").AsFloat <SQL.FieldByName("VALORMINIMO").AsFloat Then
          bsShowMessage("O valor está abaixo do valor mínimo padrão do módulo", "I")
          InclusaoForaDoPadrao = True
          LogInclusaoForaDoPadrao = "Inclusão de faixa abaixo do valor mínimo padrão" + Chr(13) + "Mínimo: " + SQL.FieldByName("VALORMINIMO").AsString + "  Valor: " + CurrentQuery.FieldByName("VALOR").AsString
      End If

      Set SQL = Nothing
    End If

  End If


  'Mochi sms 84108
  If CurrentQuery.FieldByName("VALORCONTRATO").Value <>0 Then
    Dim qContrato As Object
    Set qContrato = NewQuery

    qContrato.Active = False
    qContrato.Clear
    qContrato.Add(" SELECT C.TABTIPOCONTRATO, C.LOCALFATURAMENTO, TF.CODIGO     ")
    qContrato.Add(" FROM SAM_FAMILIA_MODADESAOPRC PRC,           ")
    qContrato.Add("      SAM_FAMILIA_MODADESAO AD,               ")
    qContrato.Add("      SAM_FAMILIA_MOD MOD,                    ")
    qContrato.Add("      SAM_CONTRATO_MOD CM,                    ")
    qContrato.Add("      SAM_CONTRATO C,                         ")
    qContrato.Add("      SIS_TIPOFATURAMENTO TF                  ")
    qContrato.Add("WHERE PRC.FAMILIAMODADESAO = AD.HANDLE AND   ")
    qContrato.Add("      AD.FAMILIAMODULO = MOD.HANDLE AND      ")
    qContrato.Add("      MOD.MODULO = CM.HANDLE AND              ")
    qContrato.Add("      CM.CONTRATO = C.HANDLE AND              ")
    qContrato.Add("      C.TIPOFATURAMENTO = TF.HANDLE AND       ")
    qContrato.Add("      PRC.HANDLE = :HANDLE                     ")
    qContrato.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_FAMILIA_MODADESAOPRC")
    qContrato.Active = True
    'codigo TipoFaturamento 130 =autogestao
    If(qContrato.FieldByName("CODIGO").Value <>130)Or _
       (qContrato.FieldByName("TABTIPOCONTRATO").AsInteger = 1 And qContrato.FieldByName("LOCALFATURAMENTO").AsString = "C")Then
      bsShowMessage(" 'Valor do contrato' deve ser informado quando tipo de faturamento do contrato for autogestão com faturamento na família", "E")
      CanContinue = False
      Set qContrato = Nothing
      Exit Sub
    End If
    Set qContrato = Nothing
  End If


End Sub


Public Function VerificaPreco As String
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT D.TIPOCALCULOPRECO FROM SAM_FAMILIA_MODADESAOPRC A, SAM_FAMILIA_MODADESAO B,")
  SQL.Add("       SAM_FAMILIA_MOD C, SAM_CONTRATO_MOD D")
  SQL.Add("WHERE A.HANDLE = :HANDLE")
  SQL.Add("  AND B.HANDLE = A.FAMILIAMODADESAO")
  SQL.Add("  AND C.HANDLE = B.FAMILIAMODULO")
  SQL.Add("  AND D.HANDLE = C.MODULO")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAMODADESAOPRC").Value
  SQL.Active = True
  VerificaPreco = SQL.FieldByName("TIPOCALCULOPRECO").Value

  SQL.Active = False
  Set SQL = Nothing
End Function

