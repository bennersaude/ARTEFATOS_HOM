'HASH: A7AEEC95806DED5917BF1CE562A0DD5B
'Macro: SAM_CONTRATO_MODADESAOPRC_FX
'#Uses "*bsShowMessage"

Dim InclusaoForaDoPadrao As Boolean
Dim LogInclusaoForaDoPadrao As String

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
  SQL.Add("FROM SAM_CONTRATO_MODADESAOPRC")
  SQL.Add("WHERE HANDLE = :HCONTRATOMODADESAOPRC")
  SQL.ParamByName("HCONTRATOMODADESAOPRC").Value = CurrentQuery.FieldByName("CONTRATOMODADESAOPRC").AsInteger
  SQL.Active = True
  vCompetencia = CurrentQuery.FieldByName("CONTRATOMODADESAOPRC").AsInteger


  SQL.Clear
  SQL.Add("SELECT MIN(COMPETENCIA) PRIMEIRACOMPETENCIA, MAX(COMPETENCIA) ULTIMACOMPETENCIA")
  SQL.Add("FROM SFN_FATURA_LANC_MOD A, SAM_BENEFICIARIO_MOD BM")
  SQL.Add("WHERE BM.MODULO = :HCONTRATOMOD")
  SQL.Add("  AND (   BM.DATACANCELAMENTO IS NULL")
  SQL.Add("       OR BM.DATACANCELAMENTO >= :HOJE)")
  SQL.Add("  AND A.BENEFICIARIOMOD = BM.HANDLE")
  SQL.Add("  AND A.COMPETENCIA >= :COMPETENCIA")
  If WebMode Or visiblemod Then
    SQL.ParamByName("HCONTRATOMOD").Value = RecordHandleOfTable("SAM_CONTRATO_MOD")
  Else
    SQL.ParamByName("HCONTRATOMOD").Value = BuscaHandleContrato("HANDLECONTRATOMOD")
  End If
  SQL.ParamByName("HOJE").Value = ServerDate
  SQL.ParamByName("COMPETENCIA").Value = vCompetencia
  SQL.Active = True

  vPrimeiraCompetencia = SQL.FieldByName("PRIMEIRACOMPETENCIA").AsDateTime
  vUltimaCompetencia = SQL.FieldByName("ULTIMACOMPETENCIA").AsDateTime

  If Not SQL.FieldByName("PRIMEIRACOMPETENCIA").IsNull Then
    SQL.Clear
    SQL.Add("SELECT HANDLE")
    SQL.Add("FROM SAM_ROTINARECALCULOMENSALID")
    SQL.Add("WHERE SITUACAOPROCESSAMENTO = '1'   ")
    SQL.Add("  AND TABRECALCULAR = 2")
    SQL.Add("  AND COMPETENCIAINICIAL = :COMPETENCIAINICIAL")
    SQL.Add("  AND COMPETENCIAFINAL = :COMPETENCIAFINAL")
    SQL.Add("  AND CONTRATOINICIAL = :HCONTRATO")
    SQL.Add("  AND CONTRATOFINAL = :HCONTRATO")
    SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
    SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
    If WebMode Or VisibleMode Then
      SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
    Else
      SQL.ParamByName("HCONTRATO").Value = BuscaHandleContrato("HANDLECONTRATO")
    End If
    SQL.Active = True

    If SQL.EOF Then

      SQL.Clear
      SQL.Add("INSERT INTO SAM_ROTINARECALCULOMENSALID")
      SQL.Add("(HANDLE, CODIGO, DESCRICAO, DATAROTINA, TABRECALCULAR,")
      SQL.Add(" COMPETENCIAINICIAL, COMPETENCIAFINAL, CONTRATOINICIAL, CONTRATOFINAL,")
      SQL.Add(" USUARIO, DATAINCLUSAO, SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO)")
      SQL.Add("VALUES")
      SQL.Add("(:HANDLE, :HANDLE, :DESCRICAO, :DATAROTINA, 2,")
      SQL.Add(" :COMPETENCIAINICIAL, :COMPETENCIAFINAL, :HCONTRATO, :HCONTRATO,")
      SQL.Add(" :USUARIO, :DATAINCLUSAO, '1', '1')")

      SQL.ParamByName("HANDLE").Value = NewHandle("SAM_ROTINARECALCULOMENSALID")
      SQL.ParamByName("DATAROTINA").Value = ServerDate
      SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
      SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
      If WebMode Or VisibleMode Then
        SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
      Else
        SQL.ParamByName("HCONTRATO").Value = BuscaHandleContrato("HANDLECONTRATO")
      End If
      SQL.ParamByName("USUARIO").Value = CurrentUser
      SQL.ParamByName("DATAINCLUSAO").Value = ServerDate
      SQL.ParamByName("DESCRICAO").Value = "Alteração nas faixas de preço do módulo"

      SQL.ExecSQL

    End If

  End If

  Set SQL = Nothing

End Sub

Public Sub TABLE_AfterPost()
  Dim vPrimeiraCompetencia As Date
  Dim vUltimaCompetencia As Date
  Dim vCompetencia As Date
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT COMPETENCIA")
  SQL.Add("FROM SAM_CONTRATO_MODADESAOPRC")
  SQL.Add("WHERE HANDLE = :HCONTRATOMODADESAOPRC")

  If WebMode Or VisibleMode Then
  	SQL.ParamByName("HCONTRATOMODADESAOPRC").Value = RecordHandleOfTable("SAM_CONTRATO_MODADESAOPRC")
  Else
	SQL.ParamByName("HCONTRATOMODADESAOPRC").Value = BuscaHandleContrato("HANDLEMODADESAOPRC")
  End If
  SQL.Active = True
  vCompetencia = CurrentQuery.FieldByName("CONTRATOMODADESAOPRC").AsInteger

  SQL.Clear
  SQL.Add("SELECT MIN(COMPETENCIA) PRIMEIRACOMPETENCIA, MAX(COMPETENCIA) ULTIMACOMPETENCIA")
  SQL.Add("FROM SFN_FATURA_LANC_MOD A, SAM_BENEFICIARIO_MOD BM")
  SQL.Add("WHERE BM.MODULO = :HCONTRATOMOD")
  SQL.Add("  AND (   BM.DATACANCELAMENTO IS NULL")
  SQL.Add("       OR BM.DATACANCELAMENTO >= :HOJE)")
  SQL.Add("  AND A.BENEFICIARIOMOD = BM.HANDLE")
  SQL.Add("  AND A.COMPETENCIA >= :COMPETENCIA")
  If WebMode Or VisibleMode Then
  	SQL.ParamByName("HCONTRATOMOD").Value = RecordHandleOfTable("SAM_CONTRATO_MOD")
  Else
    SQL.ParamByName("HCONTRATOMOD").Value = BuscaHandleContrato("HANDLECONTRATOMOD")
  End If
  SQL.ParamByName("HOJE").Value = ServerDate
  SQL.ParamByName("COMPETENCIA").Value = vCompetencia
  SQL.Active = True

  vPrimeiraCompetencia = SQL.FieldByName("PRIMEIRACOMPETENCIA").AsDateTime
  vUltimaCompetencia = SQL.FieldByName("ULTIMACOMPETENCIA").AsDateTime

  If Not SQL.FieldByName("PRIMEIRACOMPETENCIA").IsNull Then
    SQL.Clear
    SQL.Add("SELECT COUNT(1) QTDE")
    SQL.Add("  FROM SAM_ROTINARECALCULOMENSALID")
    SQL.Add(" WHERE SITUACAOPROCESSAMENTO = '1'")
    SQL.Add("   AND TABRECALCULAR = 2")
    SQL.Add("   AND COMPETENCIAINICIAL = :COMPETENCIAINICIAL")
    SQL.Add("   AND COMPETENCIAFINAL = :COMPETENCIAFINAL")
    SQL.Add("   AND CONTRATOINICIAL = :HCONTRATO")
    SQL.Add("   AND CONTRATOFINAL = :HCONTRATO")
    SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
    SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
    If WebMode Or VisibleMode Then
      SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
    Else
	  SQL.ParamByName("HCONTRATO").Value = BuscaHandleContrato("HANDLECONTRATO")
    End If
    SQL.Active = True

    If SQL.FieldByName("QTDE").AsInteger = 0 Then

      SQL.Clear
      SQL.Add("INSERT INTO SAM_ROTINARECALCULOMENSALID")
      SQL.Add("(HANDLE, CODIGO, DESCRICAO, DATAROTINA, TABRECALCULAR,")
      SQL.Add(" COMPETENCIAINICIAL, COMPETENCIAFINAL, CONTRATOINICIAL, CONTRATOFINAL,")
      SQL.Add(" USUARIO, DATAINCLUSAO, SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO)")
      SQL.Add("VALUES")
      SQL.Add("(:HANDLE, :HANDLE, :DESCRICAO, :DATAROTINA, 2,")
      SQL.Add(" :COMPETENCIAINICIAL, :COMPETENCIAFINAL, :HCONTRATO, :HCONTRATO,")
      SQL.Add(" :USUARIO, :DATAINCLUSAO, '1', '1')")

      SQL.ParamByName("HANDLE").Value = NewHandle("SAM_ROTINARECALCULOMENSALID")
      SQL.ParamByName("DATAROTINA").Value = ServerDate
      SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
      SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
      If WebMode Or VisibleMode Then
        SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
      Else
        SQL.ParamByName("HCONTRATO").Value = BuscaHandleContrato("HANDLECONTRATO")
      End If
      SQL.ParamByName("USUARIO").Value = CurrentUser
      SQL.ParamByName("DATAINCLUSAO").Value = ServerDate
      SQL.ParamByName("DESCRICAO").Value = "Alteração nas faixas de preço do módulo"

      SQL.ExecSQL

    End If

  End If

  Set SQL = Nothing

  If InclusaoForaDoPadrao Then
    WriteAudit("|", HandleOfTable("SAM_CONTRATO_MODADESAOPRC_FX"), CurrentQuery.FieldByName("HANDLE").Value, LogInclusaoForaDoPadrao)
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If WebMode Then
	CONTRATOTPDEP.WebLocalWhere = "A.CONTRATO = " + CStr(RecordHandleOfTable("SAM_CONTRATO"))
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT HANDLE")
  SQL.Add("FROM SAM_CONTRATO_REPACTUACAO")
  SQL.Add("WHERE CONTRATOMODADESAOPRCFX = :HMODADESAOPRCFX")
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

  SQL.Add("SELECT DATACANCELAMENTO FROM SAM_CONTRATO_MOD WHERE HANDLE = :HCONTRATOMOD")

  If VisibleMode Or WebMode Then
  	SQL.ParamByName("HCONTRATOMOD").Value = RecordHandleOfTable("SAM_CONTRATO_MOD")
  Else
    SQL.ParamByName("HCONTRATOMOD").Value = BuscaHandleContrato("HANDLECONTRATOMOD")
  End If

  SQL.Active = True
  If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("Módulo cancelado não permite manutenções", "I")
    CurrentQuery.Cancel
    RefreshNodesWithTable("SAM_CONTRATO_MODADESAOPRC_FX")
  End If

End Sub

Public Function VerificaPreco As String
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT C.TIPOCALCULOPRECO FROM SAM_CONTRATO_MODADESAOPRC A, SAM_CONTRATO_MODADESAO B,")
  SQL.Add("       SAM_CONTRATO_MOD C")
  SQL.Add("WHERE A.HANDLE = :HANDLE")
  SQL.Add("  AND B.HANDLE = A.CONTRATOMODADESAO")
  SQL.Add("  AND C.HANDLE = B.CONTRATOMODULO")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CONTRATOMODADESAOPRC").Value
  SQL.Active = True
  VerificaPreco = SQL.FieldByName("TIPOCALCULOPRECO").AsString

  SQL.Active = False
  Set SQL = Nothing
End Function
Public Function FaixaEtariaInvalida As Boolean
	Dim vSQL As BPesquisa
	Set vSQL = NewQuery
	vSQL.Clear
	vSQL.Add(" SELECT COUNT(1) QTDE")
    vSQL.Add("        FROM SAM_CONTRATO_REPACTUACAO SCR")
    vSQL.Add("        Join SAM_CONTRATO_MODADESAOPRC_FX SCMF On SCMF.Handle = SCR.CONTRATOMODADESAOPRCFX")
    vSQL.Add(" WHERE SCMF.IDADEMAXIMA <> :QTD ")
    vSQL.Add("       And :QTD BETWEEN SCR.IDADE And SCMF.IDADEMAXIMA")
    vSQL.Add("       And SCMF.CONTRATOMODADESAOPRC = :IHANDLE        ")
    vSQL.ParamByName("IHANDLE").Value = CurrentQuery.FieldByName("CONTRATOMODADESAOPRC").AsInteger
    vSQL.ParamByName("QTD").Value = CurrentQuery.FieldByName("IDADEMAXIMA").AsInteger
    vSQL.Active = True

    FaixaEtariaInvalida = False

    If vSQL.FieldByName("QTDE").AsInteger > 0 Then
	  FaixaEtariaInvalida = True
	  bsShowMessage("Não pode ser inserida uma Faixa Etária que abranja uma Pactuação já criada.", "E")
	End If

    vSQL.Active = False
    Set vSQL = Nothing
End Function

Public Function FaixaDuplicada As Boolean
  Dim vSQL As BPesquisa
  Dim especifico As Object

  Set vSQL = NewQuery
  Set especifico = CreateBennerObject("Especifico.uEspecifico")

  vSQL.Clear
  vSQL.Add(" SELECT COUNT(1) AS COUNT                                                ")
  vSQL.Add("   FROM SAM_CONTRATO_MODADESAOPRC_FX                                     ")
  vSQL.Add("  WHERE CONTRATOMODADESAOPRC = :PCODIGO                                  ")
  vSQL.Add("    AND QTDBENEFICIARIOS = :PQTDBENEFICIARIOS                            ")
  vSQL.Add("    AND IDADEMAXIMA = :PIDADEMAXIMA                                      ")
  vSQL.Add("    AND CODIGOTABELAPRC = :PCODIGOTABELAPRC                              ")
  vSQL.Add("    AND (CONTRATOTPDEP IS NULL OR CONTRATOTPDEP = :PCONTRATOTPDEP)       ")
  vSQL.Add("    AND (GRUPODEPENDENTE IS NULL OR GRUPODEPENDENTE = :PGRUPODEPENDENTE) ")
  vSQL.Add("    AND HANDLE <> :HANDLE                                                ")

  vSQL.Add(especifico.BEN_VerificarFaixaEtariaModuloDuplicada(CurrentSystem, CurrentQuery.TQuery))
  vSQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  vSQL.ParamByName("PCODIGO").Value = CurrentQuery.FieldByName("CONTRATOMODADESAOPRC").AsInteger
  vSQL.ParamByName("PQTDBENEFICIARIOS").Value = CurrentQuery.FieldByName("QTDBENEFICIARIOS").AsInteger
  vSQL.ParamByName("PIDADEMAXIMA").Value = CurrentQuery.FieldByName("IDADEMAXIMA").AsInteger
  vSQL.ParamByName("PCODIGOTABELAPRC").Value = CurrentQuery.FieldByName("CODIGOTABELAPRC").AsString
  vSQL.ParamByName("PCONTRATOTPDEP").Value = CurrentQuery.FieldByName("CONTRATOTPDEP").AsInteger
  vSQL.ParamByName("PGRUPODEPENDENTE").Value = CurrentQuery.FieldByName("GRUPODEPENDENTE").AsString

  If Not CurrentQuery.FieldByName("SETOR").IsNull Then
    vSQL.Add("    AND SETOR = :SETOR ")
    vSQL.ParamByName("SETOR").Value = CurrentQuery.FieldByName("SETOR").AsInteger
  Else
    vSQL.Add("    AND SETOR IS NULL ")
  End If

  If Not CurrentQuery.FieldByName("CARGO").IsNull Then
    vSQL.Add("    AND CARGO = :CARGO ")
    vSQL.ParamByName("CARGO").Value = CurrentQuery.FieldByName("CARGO").AsInteger
  Else
    vSQL.Add("    AND CARGO IS NULL ")
  End If

  If Not CurrentQuery.FieldByName("NIVEL").IsNull Then
    vSQL.Add("    AND NIVEL = :NIVEL ")
    vSQL.ParamByName("NIVEL").Value = CurrentQuery.FieldByName("NIVEL").AsInteger
  Else
    vSQL.Add("    AND NIVEL IS NULL ")
  End If

  vSQL.Active = True

  FaixaDuplicada = False

  If vSQL.FieldByName("COUNT").AsInteger > 0 Then
    FaixaDuplicada = True
    bsShowMessage("Já Existe uma faixa cadastrada com os mesmos dados.", "E")
  End If

  Set especifico = Nothing
  Set vSQL = Nothing
End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)

If (Not CurrentQuery.FieldByName("SETOR").IsNull) And (CurrentQuery.FieldByName("CARGO").IsNull Or CurrentQuery.FieldByName("NIVEL").IsNull) Then
  CanContinue = False
  bsShowMessage("Ao informar o Setor é obrigatório informar Cargo e Nível", "E")
End If

If FaixaEtariaInvalida Then
   CanContinue = False
End If
If FaixaDuplicada Then
   CanContinue = False
End If

  Dim TipoCalculoPreco As String

  TipoCalculoPreco = VerificaPreco

  If TipoCalculoPreco <>"2" And _
      Not(CurrentQuery.FieldByName("CONTRATOTPDEP").IsNull)Then
    CanContinue = False
    bsShowMessage("A configuração do módulo NÃO permite que se informe o Tipo dependente", "E")
  End If

  If TipoCalculoPreco <>"3" And _
      Not(CurrentQuery.FieldByName("GRUPODEPENDENTE").IsNull)Then

    If WebMode Then
      InfoDescription = "A configuração do módulo NÃO permite que se informe o Grupo dependente."
      CurrentQuery.FieldByName("GRUPODEPENDENTE").Clear
    Else
      CanContinue = False
      bsShowMessage("A configuração do módulo NÃO permite que se informe o Grupo dependente.", "E")
      CurrentQuery.FieldByName("GRUPODEPENDENTE").Clear
    End If

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
    SQL.Add("SELECT E.VALORMINIMO, E.VALORCUSTO")
    SQL.Add("FROM SAM_CONTRATO_MODADESAOPRC A, SAM_CONTRATO_MODADESAO B, SAM_CONTRATO_MOD C, SAM_MODULO_PRECO D, SAM_MODULO_PRECO_FX E")
    SQL.Add("WHERE A.HANDLE = :HCONTRATOMODADESAOPRC")
    SQL.Add("  AND B.HANDLE = A.CONTRATOMODADESAO")
    SQL.Add("  AND C.HANDLE = B.CONTRATOMODULO")
    SQL.Add("  AND D.MODULO = C.MODULO")
    SQL.Add("  AND D.COMPETENCIAINICIAL <= A.COMPETENCIA")
    SQL.Add("  AND (D.COMPETENCIAFINAL IS NULL OR D.COMPETENCIAFINAL >= A.COMPETENCIA)")
    SQL.Add("  AND E.MODULOPRECO = D.HANDLE")
    SQL.Add("  AND E.IDADEMAXIMA = (SELECT MIN(IDADEMAXIMA)")
    SQL.Add("                       FROM SAM_MODULO_PRECO_FX")
    SQL.Add("                       WHERE MODULOPRECO = D.HANDLE")
    SQL.Add("                         AND IDADEMAXIMA >= :IDADEMAXIMA)")
    SQL.ParamByName("HCONTRATOMODADESAOPRC").Value = CurrentQuery.FieldByName("CONTRATOMODADESAOPRC").AsInteger
    SQL.ParamByName("IDADEMAXIMA").Value = CurrentQuery.FieldByName("IDADEMAXIMA").AsInteger
    SQL.Active = True

    If Not SQL.EOF Then
      If CurrentQuery.FieldByName("VALOR").AsFloat <SQL.FieldByName("VALORCUSTO").AsFloat Then
        bsShowMessage("O valor está abaixo do valor de custo padrão do módulo. Deseja continuar?", "I")
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

  'Andreia sms 22955
  If CurrentQuery.FieldByName("VALORCONTRATO").Value <>0 Then
    Dim qContrato As Object
    Set qContrato = NewQuery

    qContrato.Active = False
    qContrato.Clear
    qContrato.Add(" SELECT C.TABTIPOCONTRATO, C.LOCALFATURAMENTO, TF.CODIGO     ")
    qContrato.Add(" FROM SAM_CONTRATO_MODADESAOPRC_FX FX,        ")
    qContrato.Add("      SAM_CONTRATO_MODADESAOPRC PRC,          ")
    qContrato.Add("      SAM_CONTRATO_MODADESAO AD,              ")
    qContrato.Add("      SAM_CONTRATO_MOD MOD,                   ")
    qContrato.Add("      SAM_CONTRATO C,                         ")
    qContrato.Add("      SIS_TIPOFATURAMENTO TF                  ")
    qContrato.Add("WHERE FX.CONTRATOMODADESAOPRC = PRC.HANDLE AND")
    qContrato.Add("      PRC.CONTRATOMODADESAO = AD.HANDLE AND   ")
    qContrato.Add("      AD.CONTRATOMODULO = MOD.HANDLE AND      ")
    qContrato.Add("      MOD.CONTRATO = C.HANDLE AND             ")
    qContrato.Add("      C.TIPOFATURAMENTO = TF.HANDLE AND       ")
    qContrato.Add("      FX.HANDLE = :HANDLE                     ")
    qContrato.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
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
'sms 22955

End Sub

Public Function BuscaHandleContrato(campo As String) As String

Set qBuscaHandleContrato = NewQuery

qBuscaHandleContrato.Add("SELECT MO.HANDLE HANDLECONTRATOMOD, MA.HANDLE HANDLEMODADESAOPRC, CO.HANDLE HANDLECONTRATO  ")
qBuscaHandleContrato.Add("  FROM SAM_CONTRATO_MODADESAOPRC_FX CM                                                      ")
qBuscaHandleContrato.Add("    JOIN SAM_CONTRATO_MODADESAOPRC MA ON MA.HANDLE = CM.CONTRATOMODADESAOPRC                ")
qBuscaHandleContrato.Add("    JOIN SAM_CONTRATO_MODADESAO CA ON CA.HANDLE = MA.CONTRATOMODADESAO                      ")
qBuscaHandleContrato.Add("    JOIN SAM_CONTRATO_MOD MO ON MO.HANDLE = CA.CONTRATOMODULO                               ")
qBuscaHandleContrato.Add("    JOIN SAM_CONTRATO CO ON CO.HANDLE = MO.CONTRATO                                         ")
qBuscaHandleContrato.Add("  WHERE CM.HANDLE = :HCONTRATOMOD                                                           ")

qBuscaHandleContrato.ParamByName("HCONTRATOMOD").Value = CurrentQuery.FieldByName("HANDLE").AsString
qBuscaHandleContrato.Active = True
BuscaHandleContrato = qBuscaHandleContrato.FieldByName(campo).AsString

End Function
