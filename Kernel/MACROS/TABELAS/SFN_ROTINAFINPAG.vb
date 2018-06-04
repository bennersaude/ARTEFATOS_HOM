'HASH: 15E49F995F78E3A37B32F2B7547CDF88
'Macro: SFN_ROTINAFINPAG
'#Uses "*bsShowMessage"

Option Explicit
Dim qAux As Object

Public Sub ARQUIVOLOG_OnBeforeCommand(CanContinue As Boolean, ByVal idCommand As Long)

  If (idCommand = 3&) Then
	If bsShowMessage("Você tem certeza que deseja excluir este arquivo?", "Q") = vbNo Then
	  CanContinue = False
	End If
  End If

End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Obj As Object
  Dim SQLRotFin As Object
  Dim vMsg As String

  If CurrentQuery.FieldByName("TABSELECAO").AsInteger = 4 Then
    Dim SqlPegs As BPesquisa
    Set SqlPegs = NewQuery

    SqlPegs.Add("SELECT 1 EXISTE FROM SFN_ROTINAFINPAG_PEG WHERE ROTINAFINPAG = :HANDLE")
    SqlPegs.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    SqlPegs.Active = True

    If SqlPegs.FieldByName("EXISTE").AsInteger = 0 Then

      bsShowMessage("Esta rotina está parametrizada para selecionar por 'Seleção de PEG's' e nenhum PEG foi inserido na carga da rotina", "I")
      Exit Sub
    End If
    Set SqlPegs = Nothing

  End If

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "E")
    Exit Sub
  End If

  Set SQLRotFin = NewQuery
  SQLRotFin.Add("SELECT SITUACAO, DATACONTABIL, SEQUENCIA, DATAROTINA, TABDATACONTABIL, DESCRICAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  SQLRotFin.Active = True

  If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < _
                              SQLRotFin.FieldByName("DATACONTABIL").AsDateTime Then
    SQLRotFin.Active = False
    Set SQLRotFin = Nothing
    bsShowMessage("A data de pagamento não pode ser anterior a data contábil", "E")
    Exit Sub
  End If


  If (SQLRotFin.FieldByName("TABDATACONTABIL").AsInteger = 2) Then
    If (CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < SQLRotFin.FieldByName("DATAROTINA").AsDateTime) Then
      bsShowMessage("A data de vencimento não pode ser anterior a data de emissão da rotina!", "E")
      Set SQLRotFin = Nothing
  	  Exit Sub
	End If
  End If

  If (VisibleMode) Then
    Set Obj = CreateBennerObject("BSINTERFACE0040.PAGAMENTO")
    Obj.PROCESSAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").Value, vMsg)

  Else
    Dim vsMensagemErro As String
    Dim viRetorno As Long
    Dim vcContainer As CSDContainer

	Set vcContainer = NewContainer

	vcContainer.AddFields("HANDLE:INTEGER")
    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SAMPAGAMENTO", _
                                     "PAGAMENTO_PAGAR", _
                                     "Rotina de Pagamento de Guias (Processamento)|" + _
                                     "Sequência|" + SQLRotFin.FieldByName("SEQUENCIA").AsString + "|" + _
                                     "Descrição|" + SQLRotFin.FieldByName("DESCRICAO").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINPAG", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagemErro, _
                                     vcContainer)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If

    Set vcContainer = Nothing
  End If

  Set Obj = Nothing

  SQLRotFin.Active = False
  Set SQLRotFin = Nothing

  WriteAudit("P", HandleOfTable("SFN_ROTINAFINPAG"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Faturamento de Pagamentos - Processamento")
  RefreshNodesWithTable("SAM_PEG")

End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object
  Dim vMsg As String
  Dim SQLRotFin As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "E")
    Exit Sub
  End If

  Set SQLRotFin = NewQuery
  SQLRotFin.Add("SELECT SITUACAO, DATACONTABIL, SEQUENCIA, DESCRICAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  SQLRotFin.Active = True

  If (VisibleMode) Then
    Set Obj = CreateBennerObject("BSINTERFACE0040.PAGAMENTO")
    Obj.CANCELAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").Value, vMsg)
  Else
    Dim vsMensagemErro As String
    Dim viRetorno As Long
	Dim vcContainer As CSDContainer

	Set vcContainer = NewContainer

	vcContainer.AddFields("HANDLE:INTEGER")
    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SAMPAGAMENTO", _
                                     "PAGAMENTO_CANCELAR", _
                                     "Rotina de Pagamento de Guias (Cancelamento)|" + _
                                     "Sequência|"+ SQLRotFin.FieldByName("SEQUENCIA").AsString + "|" + _
                                     "Descrição|" + SQLRotFin.FieldByName("DESCRICAO").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINPAG", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "C", _
                                     False, _
                                     vsMensagemErro, _
                                     vcContainer)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If

    Set vcContainer = Nothing
  End If

  SQLRotFin.Active = False
  Set SQLRotFin = Nothing

  Set Obj = Nothing

  WriteAudit("C", HandleOfTable("SFN_ROTINAFINPAG"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Faturamento de Pagamentos - Cancelamento")
  RefreshNodesWithTable("SFN_ROTINAFIN")

End Sub

Public Sub VerificaSeProcessada(CanContinue As Boolean)
  Dim SQLRotFin As Object
  Set SQLRotFin = NewQuery
  SQLRotFin.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  SQLRotFin.Active = True
  If CurrentQuery.FieldByName("SITUACAO").Value <> "1" Then
    CanContinue = False
    SQLRotFin.Active = False
    Set SQLRotFin = Nothing
    bsShowMessage("A Rotina já foi processada", "E")
    Exit Sub
  End If
  SQLRotFin.Active = False
  Set SQLRotFin = Nothing
End Sub


Public Sub TABLE_AfterScroll()
  If (Not WebMode) Then
    If (CurrentQuery.FieldByName("SITUACAO").AsString = "1") Then
      BOTAOPROCESSAR.Enabled = True
      BOTAOCANCELAR.Enabled = False
    ElseIf (CurrentQuery.FieldByName("SITUACAO").AsString = "5") Then
      BOTAOPROCESSAR.Enabled = False
      BOTAOCANCELAR.Enabled = True
    Else
      BOTAOPROCESSAR.Enabled = False
      BOTAOCANCELAR.Enabled = False
    End If
  End If

  If WebMode Then
    Dim SQL As Object
    Dim rp As Integer

    Set SQL = NewQuery
    SQL.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE IN (SELECT TIPOFATURAMENTO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE)")
    SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SFN_ROTINAFIN")
    SQL.Active = True
    'SMS 30184 - Cazangi - Fim
    If SQL.FieldByName("CODIGO").AsInteger = 310 Then
      rp = 2 'reembolso
    Else
      rp = 1
    End If
    Set SQL = Nothing

    Dim STRX As String
    If InStr(SQLServer, "DB2") > 0 Then
      STRX = " AND timestamp_iso(date(A.DATAPAGAMENTO)) = "
    ElseIf InStr(SQLServer, "ORACLE") > 0 Then
      STRX = " AND trunc(A.DATAPAGAMENTO) = "
    ElseIf InStr(SQLServer, "MSSQL") > 0 Then
      STRX = " AND convert(datetime, cast(floor(convert(float, A.DATAPAGAMENTO)) as int)) = "
    Else
      STRX = " AND CONVERT(DATETIME , CAST(A.DATAPAGAMENTO  AS DATE), 103) ="
    End If


    PEGINICIAL.WebLocalWhere = "A.SITUACAO = '3' AND A.TABREGIMEPGTO = " + Str(rp) + " " + STRX + "@CAMPO(DATAPAGAMENTO)"

    PEGFINAL.WebLocalWhere = " A.PEG >= (Select PEG FROM SAM_PEG WHERE HANDLE = @CAMPO(PEGINICIAL) ) " + _
                        " And A.SITUACAO = '3' AND A.TABREGIMEPGTO = " + Str(rp) + " AND A.DATAPAGAMENTO = @CAMPO(DATAPAGAMENTO) "

    'Inicio - SMS 96895 - RODRIGO ANDRADE
    AbreQueryTipoRotina
    If (qAux.FieldByName("EHREAPRESENTACAO").AsString = "S") Then
      PEGINICIAL.WebLocalWhere = PEGINICIAL.WebLocalWhere + " AND A.PEGORIGINAL > 0 "
    ElseIf (qAux.FieldByName("CONTROLEPAGAMENTO").AsString <> "S") And (qAux.FieldByName("REAPRESENTADOJUNTONORMAL").AsString <> "S") Then
      PEGINICIAL.WebLocalWhere = PEGINICIAL.WebLocalWhere + " AND A.PEGORIGINAL IS NULL "
    End If
    'Fim - SMS 96895 - RODRIGO ANDRADE

  End If


   'Seleciona o parametro, para a comparação da carga prestador
   '-SMS-338695-------------------------------------------------
	Dim qArquivoLog As BPesquisa
	Set qArquivoLog = NewQuery

	qArquivoLog.Add("SELECT ST.TIPOROTINAFINANCEIRA                                    ")
	qArquivoLog.Add("  FROM SIS_TIPOFATURAMENTO ST                                     ")
	qArquivoLog.Add("  JOIN SFN_COMPETFIN       SC ON (ST.HANDLE = SC.TIPOFATURAMENTO) ")
	qArquivoLog.Add("  JOIN SFN_ROTINAFIN       SR ON (SC.HANDLE = SR.COMPETFIN)       ")
	qArquivoLog.Add(" WHERE SR.HANDLE = :HANDLE                                        ")
	If Not(WebMode) Then
	  qArquivoLog.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINAFIN")
    Else
      qArquivoLog.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
	End If

	qArquivoLog.Active = True

	If qArquivoLog.FieldByName("TIPOROTINAFINANCEIRA").AsString = "P" Then
	  ARQUIVOLOG.Visible = True
	Else
	  ARQUIVOLOG.Visible = False
	End If

	qArquivoLog.Active = False
	Set qArquivoLog = Nothing
   '------------------------------------------------------------

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If NodeInternalCode <> 802 Then
    VerificaSeProcessada(CanContinue)
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If NodeInternalCode <> 802 Then
    VerificaSeProcessada(CanContinue)
  End If
End Sub

Public Sub TABREGIMEPGTO_OnChanging(AllowChange As Boolean)
  If NodeInternalCode <> 802 Then
    VerificaSeProcessada(AllowChange)
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Verificar se já existe um registro
  Dim qVerificaExiste As Object
  Set qVerificaExiste = NewQuery
  qVerificaExiste.Add("SELECT COUNT(*) QTD")
  qVerificaExiste.Add("  FROM SFN_ROTINAFINPAG")
  qVerificaExiste.Add(" WHERE HANDLE <> :HROTINAFINPAG")
  qVerificaExiste.Add("   AND ROTINAFIN = :HROTINAFIN")
  qVerificaExiste.ParamByName("HROTINAFINPAG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaExiste.ParamByName("HROTINAFIN").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  qVerificaExiste.Active = True

  If qVerificaExiste.FieldByName("QTD").AsInteger > 0 Then
    bsShowMessage("Já existe um registro para essa rotina!", "E")
    CanContinue = False
    Exit Sub
  End If

  qVerificaExiste.Active = False
  Set qVerificaExiste = Nothing

  'Verificar se a data de pagamento está dentro da competência
  Dim SQLCompetFin As Object
  Set SQLCompetFin =NewQuery

  If NodeInternalCode <> 802 Then
    SQLCompetFin.Add("SELECT A.DATACONTABIL, B.COMPETENCIA, A.TABDATACONTABIL, A.DATAROTINA")
    SQLCompetFin.Add("FROM SFN_ROTINAFIN A, SFN_COMPETFIN B")
    SQLCompetFin.Add("WHERE A.HANDLE = :HANDLEROTINAFIN")
    SQLCompetFin.Add("AND B.HANDLE = A.COMPETFIN")
    SQLCompetFin.ParamByName("HANDLEROTINAFIN").Value =CurrentQuery.FieldByName("ROTINAFIN").Value
    SQLCompetFin.Active =True

    If CurrentQuery.FieldByName("TABSELECAO").AsInteger =3 Then
      If CurrentQuery.FieldByName("DATARECEBIMENTOINICIAL").AsDateTime >CurrentQuery.FieldByName("DATARECEBIMENTOFINAL").AsDateTime Then
        CanContinue =False
        bsShowMessage("A data inicial não pode ser MAIOR do que a data final !", "E")
        Exit Sub
      End If

      If (Not SQLCompetFin.FieldByName("DATACONTABIL").IsNull) Then 'Soares - SMS: 68831 - 29/10/2006 - Quando a rotina está como assumir a data contabil, o select acima nao encontra ninguem.
        If SQLCompetFin.FieldByName("DATACONTABIL").AsDateTime <CurrentQuery.FieldByName("DATARECEBIMENTOINICIAL").AsDateTime Or  _
          SQLCompetFin.FieldByName("DATACONTABIL").AsDateTime >CurrentQuery.FieldByName("DATARECEBIMENTOFINAL").AsDateTime Then
          CanContinue =False
          bsShowMessage("A data contábil não pode estar fora do período de recebimento", "E")
          Exit Sub
        End If
      End If
    End If

    If (SQLCompetFin.FieldByName("TABDATACONTABIL").AsInteger = 2) Then
	  If (CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < SQLCompetFin.FieldByName("DATAROTINA").AsDateTime) Then
		bsShowMessage("A data de vencimento não pode ser anterior a data de emissão da rotina!", "E")
		CanContinue = False
		Exit Sub
	  End If
    End If

    If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < _
      SQLCompetFin.FieldByName("DATACONTABIL").AsDateTime Then
      SQLCompetFin.Active =False
      Set SQLCompetFin =Nothing
      CanContinue =False
      bsShowMessage("A data de pagamento não pode ser anterior a data contábil", "E")
      Exit Sub
    End If

    If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime <ServerDate Then
       bsShowMessage("A data de pagamento não pode ser anterior a hoje", "E")
    End If

    If(Month(CurrentQuery.FieldByName("DATAPAGAMENTO").Value)<>Month(SQLCompetFin.FieldByName("COMPETENCIA").Value))Or  _
      (Year(CurrentQuery.FieldByName("DATAPAGAMENTO").Value)<>Year(SQLCompetFin.FieldByName("COMPETENCIA").Value))Then
       bsShowMessage("Data de pagamento fora da competência da rotina financeira", "I")
    End If

    SQLCompetFin.Active =False
    Set SQLCompetFin =Nothing

  End If
End Sub

Public Sub TABSELECAO_OnChanging(AllowChange As Boolean)
  If NodeInternalCode <> 801 Then
    VerificaSeProcessada(AllowChange)
  End If
End Sub

Public Sub PEGINICIAL_OnPopup(ShowPopup As Boolean)
  Dim SQL As Object
  Dim rp As Integer

  'SMS 30184 - Cazangi - Início
  Set SQL = NewQuery
  SQL.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE IN (SELECT TIPOFATURAMENTO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE)")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SFN_ROTINAFIN")
  SQL.Active = True
  'SMS 30184 - Cazangi - Fim
  If SQL.FieldByName("CODIGO").AsInteger = 310 Then
    rp = 2 'reembolso
  Else
    rp = 1
  End If
  Set SQL = Nothing

  Dim datapag As Date
  Dim STRX As String
  datapag = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime
  If InStr(SQLServer, "DB2") > 0 Then
  'início sms 56808 - Edilson.Castro - 21/01/2006 - truncamento do campo SAM_PEG.DATAPAGAMENTO
    STRX = " AND timestamp_iso(date(SAM_PEG.DATAPAGAMENTO)) = "
  ElseIf InStr(SQLServer, "ORACLE") > 0 Then
    STRX = " AND trunc(SAM_PEG.DATAPAGAMENTO) = "
  ElseIf InStr(SQLServer, "MSSQL") > 0 Then
    STRX = " AND convert(datetime, cast(floor(convert(float, SAM_PEG.DATAPAGAMENTO)) as int)) = "
  Else
    STRX = " AND CONVERT(DATETIME , CAST(SAM_PEG.DATAPAGAMENTO  AS DATE), 103) ="
  End If

  PEGINICIAL.LocalWhere = "SAM_PEG.SITUACAO = '3' AND SAM_PEG.TABREGIMEPGTO = " + Str(rp) + " " + STRX + SQLDate(datapag)

  'Inicio - SMS 96895 - RODRIGO ANDRADE
  AbreQueryTipoRotina
  If (qAux.FieldByName("EHREAPRESENTACAO").AsString = "S") Then
    PEGINICIAL.LocalWhere = PEGINICIAL.LocalWhere + " AND SAM_PEG.PEGORIGINAL > 0 "
  ElseIf (qAux.FieldByName("CONTROLEPAGAMENTO").AsString <> "S") And (qAux.FieldByName("REAPRESENTADOJUNTONORMAL").AsString <> "S") Then
    PEGINICIAL.LocalWhere = PEGINICIAL.LocalWhere + " AND SAM_PEG.PEGORIGINAL IS NULL "
  End If
  'Fim - SMS 96895 - RODRIGO ANDRADE


End Sub

Public Sub PEGFINAL_OnPopup(ShowPopup As Boolean)
  Dim SQL As Object
  Dim rp As Integer

  'SMS 30184 - Cazangi - Início
  Set SQL = NewQuery
  SQL.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE IN (SELECT TIPOFATURAMENTO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE)")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SFN_ROTINAFIN")
  SQL.Active = True
  'SMS 30184 - Cazangi - Fim
  If SQL.FieldByName("CODIGO").AsInteger = 310 Then
    rp = 2 'reembolso
  Else
    rp = 1
  End If
  Set SQL = Nothing

  Dim datapag As Date
  Dim STRX As String
  datapag = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime


  If InStr(SQLServer, "DB2") > 0 Then
  'início sms 56808 - Edilson.Castro - 21/01/2006 - truncamento do campo SAM_PEG.DATAPAGAMENTO
    STRX = " AND timestamp_iso(date(SAM_PEG.DATAPAGAMENTO)) = "
  ElseIf InStr(SQLServer, "ORACLE") > 0 Then
    STRX = " AND trunc(SAM_PEG.DATAPAGAMENTO) = "
  ElseIf InStr(SQLServer, "MSSQL") > 0 Then
    STRX = " AND convert(datetime, cast(floor(convert(float, SAM_PEG.DATAPAGAMENTO)) as int)) = "
  Else
    STRX = " AND CONVERT(DATETIME , CAST(SAM_PEG.DATAPAGAMENTO  AS DATE), 103) ="
    'STRX = " AND truncdate(SAM_PEG.DATAPAGAMENTO) = "
  End If

  PEGFINAL.LocalWhere = "(SAM_PEG.PEG >= " + _
                        "(Select PEG FROM SAM_PEG WHERE SAM_PEG.HANDLE = " + _
                        Str(CurrentQuery.FieldByName("PEGINICIAL").AsInteger) + ")) AND " + _
                        "(SAM_PEG.SITUACAO = '3' AND SAM_PEG.TABREGIMEPGTO = " + Str(rp) + " " + STRX + SQLDate(datapag) + ")"

  'Inicio - SMS 96895 - RODRIGO ANDRADE
  AbreQueryTipoRotina
  If (qAux.FieldByName("EHREAPRESENTACAO").AsString = "S") Then
    PEGFINAL.LocalWhere = PEGFINAL.LocalWhere + " AND SAM_PEG.PEGORIGINAL > 0 "
  ElseIf (qAux.FieldByName("CONTROLEPAGAMENTO").AsString <> "S") And (qAux.FieldByName("REAPRESENTADOJUNTONORMAL").AsString <> "S") Then
    PEGFINAL.LocalWhere = PEGFINAL.LocalWhere + " AND SAM_PEG.PEGORIGINAL IS NULL "
  End If
  'Fim - SMS 96895 - RODRIGO ANDRADE

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

 If CommandID = "BOTAOPROCESSAR" Then
    BOTAOPROCESSAR_OnClick
 ElseIf CommandID = "BOTAOCANCELAR" Then
    BOTAOCANCELAR_OnClick
 End If

End Sub

Public Function AbreQueryTipoRotina()
  Set qAux = NewQuery
  qAux.Clear
  qAux.Add("SELECT R.EHREAPRESENTACAO,        ")
  qAux.Add("       R.CONTROLEPAGAMENTO,       ")
  qAux.Add("       P.REAPRESENTADOJUNTONORMAL ")
  qAux.Add("  FROM SFN_ROTINAFIN R,           ")
  qAux.Add("       SAM_PARAMETROSPROCCONTAS P ")
  qAux.Add(" WHERE R.HANDLE = :HANDLE         ")
  qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  qAux.Active = True
End Function
