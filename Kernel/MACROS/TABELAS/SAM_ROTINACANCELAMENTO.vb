'HASH: 1D6C593656A496488D355F93522FD857
'Macro: SAM_ROTINACANCELAMENTO
'#Uses "*bsShowMessage"

Option Explicit

Public Function VerificaDataFechamento()As Boolean
  Dim qFechamento
  Set qFechamento = NewQuery

  qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")

  qFechamento.Active = True

  Dim vMesComp As Integer
  Dim vAnoComp As Integer
  Dim vMesFechamento As Integer
  Dim vAnoFechamento As Integer

  If CurrentQuery.FieldByName("TABDATACANCELAMENTO").AsInteger = 2 Then
	vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIACANCELAMENTO").AsDateTime)
	vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIACANCELAMENTO").AsDateTime)

	vMesFechamento = DatePart("m", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
	vAnoFechamento = DatePart("yyyy", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

	If (vAnoComp < vAnoFechamento) Or (vAnoComp = vAnoFechamento And vMesComp < vMesFechamento) Then
	  bsShowMessage("A competência não pode ser inferior à data de fechamento - Parâmetros Gerais", "I")
	  VerificaDataFechamento = False
	  Exit Function
	End If
  Else
	If CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime < qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
	  VerificaDataFechamento = False
	  bsShowMessage("Não é possível cancelar beneficiários com data de cancelamento inferior a data de fechamento - Parâmetros Gerais", "I")
	  Exit Function
	End If
  End If

  Set qFechamento = Nothing

  VerificaDataFechamento = True
End Function

Public Sub BOTAOEMITIRAVISO_OnClick()
  If (CurrentQuery.FieldByName("SITUACAO").AsString = "1") Or _
     (CurrentQuery.FieldByName("SITUACAOPROCESSAR").AsString = "5") Then
	bsShowMessage("A rotina está em aberto ou já foi processada.", "I")
	Exit Sub
  End If

  If bsShowMessage("Emitir aviso de cancelamento?", "Q") = vbYes Then
	'controle de correspondencia
	Dim RelatorioHandle As Long
	Dim QueryBuscaHandleRelatorio As Object
	Dim HandleRotCanc As Long
	Set QueryBuscaHandleRelatorio = NewQuery

	QueryBuscaHandleRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'BEN020'")

	QueryBuscaHandleRelatorio.Active = False
	QueryBuscaHandleRelatorio.Active = True

	RelatorioHandle = QueryBuscaHandleRelatorio.FieldByName("HANDLE").AsInteger

	Set QueryBuscaHandleRelatorio = Nothing

	HandleRotCanc = CurrentQuery.FieldByName("HANDLE").AsInteger

	ReportPreview(RelatorioHandle, "A.HANDLE=" + Str(HandleRotCanc), False, False)
	WriteAudit("E", HandleOfTable("SAM_ROTINACANCELAMENTO"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
		"Rotina de Cancelamento de Beneficiários - Emissão do Aviso")
  End If
End Sub

Public Sub BOTAOEMITIRCOMUNICADO_OnClick()
  If(CurrentQuery.FieldByName("SITUACAOPROCESSAR").AsString <>"5")Then
	bsShowMessage("A rotina não foi processada.", "I")
	Exit Sub
  End If

  If bsShowMessage("Emitir comunicado de cancelamento?", "Q") = vbYes Then
	'controle de correspondencia
	Dim RelatorioHandle As Long
	Dim QueryBuscaHandleRelatorio As Object
	Dim HandleRotCanc As Long
	Set QueryBuscaHandleRelatorio = NewQuery

	QueryBuscaHandleRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'BEN004'")

	QueryBuscaHandleRelatorio.Active = False
	QueryBuscaHandleRelatorio.Active = True

	RelatorioHandle = QueryBuscaHandleRelatorio.FieldByName("HANDLE").AsInteger

	Set QueryBuscaHandleRelatorio = Nothing

	HandleRotCanc = CurrentQuery.FieldByName("HANDLE").AsInteger

	ReportPreview(RelatorioHandle, "A.HANDLE=" + Str(HandleRotCanc), False, False)
	WriteAudit("E", HandleOfTable("SAM_ROTINACANCELAMENTO"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
		"Rotina de Cancelamento de Beneficiários - Emissão do Comunicado")

	If WebMode Then
		bsShowMessage("Emissão do relatório enviado com sucesso ao servidor.", "I")
	End If
  End If
End Sub

Public Sub BOTAOENVIAREMAIL_OnClick()

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os dados não podem estar em edição!", "E")
	Exit Sub
  End If

  Dim qEmail As BPesquisa
  Set qEmail = NewQuery

  qEmail.Active = False
  qEmail.Add(" SELECT EMAIL             ")
  qEmail.Add("   FROM Z_GRUPOUSUARIOS   ")
  qEmail.Add("  WHERE HANDLE = :HANDLE  ")
  qEmail.ParamByName("HANDLE").AsInteger = CurrentUser
  qEmail.Active = True

  If qEmail.FieldByName("EMAIL").AsString = "" Then
    Set qEmail = Nothing
    bsshowmessage("Usuário não possui e-mail cadastrado!", "E")
    Exit Sub
  End If

  Set qEmail = Nothing


  Dim qAtualizaData As Object

  Set qAtualizaData = NewQuery

  qAtualizaData.Add("  UPDATE SAM_ROTINACANCELAMENTO                     ")
  qAtualizaData.Add("     SET DATAHORAENVIOEMAIL = :DATAHORAENVIOEMAIL,  ")
  qAtualizaData.Add("         USUARIOENVIOEMAIL = :USUARIOENVIOEMAIL     ")
  qAtualizaData.Add("   WHERE HANDLE = :HANDLE                           ")
  qAtualizaData.ParamByName("DATAHORAENVIOEMAIL").AsDateTime = ServerNow
  qAtualizaData.ParamByName("USUARIOENVIOEMAIL").AsInteger = CurrentUser
  qAtualizaData.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qAtualizaData.ExecSQL

  Set qAtualizaData = Nothing

  Dim vSql As BPesquisa
  Set vSql = NewQuery

  vSql.Clear
  vSql.Add("SELECT HANDLE FROM Z_MACROS WHERE NOME = :NOME")
  vSql.ParamByName("NOME").AsString = "Envio_Email_Cancelamento_Beneficiario"
  vSql.Active = True

  Dim sx As CSServerExec
  Set sx = NewServerExec

  sx.Description = "Envio dos Avisos de Cancelamento"
  sx.Process = vSql.FieldByName("HANDLE").AsInteger
  sx.SessionVar("HANDLEROTINA") = CurrentQuery.FieldByName("HANDLE").AsString
  sx.Execute

  Set sx = Nothing

  Set vSql = Nothing

  bsshowmessage("O processamento foi iniciado internamente." & Chr(13) & "Você pode continuar navegando pelo sistema" & Chr(13) & "enquanto aguarda a conclusão.", "I" )
  RefreshNodesWithTable("SAM_ROTINACANCELAMENTO")


End Sub

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"
  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCampos = "Nº do Contrato|Contratante|Data Adesão"
  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, 2, vCampos, vCriterio, "Contratos", True, "")

  If vHandle <> 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("CONTRATO").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub BOTAOGERAR_OnClick()
  If CurrentQuery.State <> 1 Then
	bsShowMessage("Os parâmetros não podem estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
	bsShowMessage("A rotina já foi processada", "I")
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("VERIFICARINADIMPLENCIA").AsString = "S") And _
	 CurrentQuery.FieldByName("DATABASEINADIMPLENCIA").IsNull Then
	bsShowMessage("A Data base para inadimplência deve ser informada", "I")
	DATABASEINADIMPLENCIA.SetFocus
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("VERIFICARINADIMPLENCIA").AsString = "S") And _
	 CurrentQuery.FieldByName("MOTIVOINADIMPLENCIA").IsNull Then
	bsShowMessage("O Motivo para inadimplência deve ser informado", "I")
	MOTIVOINADIMPLENCIA.SetFocus
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("VERIFICARIDADEMAXIMA").AsString = "S") And _
	 CurrentQuery.FieldByName("DATABASEIDADEMAXIMA").IsNull Then
	bsShowMessage("A Data base para idade máxima deve ser informada", "I")
	DATABASEIDADEMAXIMA.SetFocus
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("VERIFICARIDADEMAXIMA").AsString = "S") And _
	 CurrentQuery.FieldByName("MOTIVOIDADEMAXIMA").IsNull Then
	bsShowMessage("O Motivo para idade máxima deve ser informado", "I")
	MOTIVOIDADEMAXIMA.SetFocus
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("VERIFICARDOCUMENTOS").AsString = "S") And _
	 CurrentQuery.FieldByName("DATABASEDOCUMENTOS").IsNull Then
	bsShowMessage("A Data base para documentos deve ser informada", "I")
	DATABASEDOCUMENTOS.SetFocus
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("VERIFICARDOCUMENTOS").AsString = "S") And _
	 CurrentQuery.FieldByName("MOTIVODOCUMENTOS").IsNull Then
	bsShowMessage("O Motivo para documentos deve ser informado", "I")
	MOTIVODOCUMENTOS.SetFocus
	Exit Sub
  End If

 'Daniela Zardo -15/07/2002
  If (CurrentQuery.FieldByName("VERIFICARDECOMPOSICAOFAMILIAR").AsString = "S") And _
	 CurrentQuery.FieldByName("DATABASEDECOMPOSICAOFAMILIAR").IsNull Then
	bsShowMessage("A Data base para decomposição familiar deve ser informada", "I")
	DATABASEDECOMPOSICAOFAMILIAR.SetFocus
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("VERIFICARDECOMPOSICAOFAMILIAR").AsString = "S") And _
	 CurrentQuery.FieldByName("MOTIVODECOMPOSICAOFAMILIAR").IsNull Then
	bsShowMessage("O Motivo para decomposição familiar deve ser informado", "I")
	MOTIVODECOMPOSICAOFAMILIAR.SetFocus
	Exit Sub
  End If

  If VerificaDataFechamento = False Then
	Exit Sub
  End If

  Dim Obj As Object

  If VisibleMode Then

    Set Obj = CreateBennerObject("BSInterface0003.RotinaCancelamento")
    Obj.Gerar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Else

    Dim vsMensagemErro As String
    Dim viRetorno As Long

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen010", _
                                     "RotinaCancelamento_Gerar", _
                                     "Rotina de Cancelamento de Beneficiários (Gerar) -  Código: " + _
                                     CurrentQuery.FieldByName("CODIGO").AsString + " Descrição: " + _
                                     CurrentQuery.FieldByName("DESCRICAO").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACANCELAMENTO", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagemErro, _
                                     Null)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If

  End If

  Set Obj = Nothing

  WriteAudit("G", HandleOfTable("SAM_ROTINACANCELAMENTO"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
	  "Rotina de Cancelamento de Beneficiários - Geração")
  RefreshNodesWithTable("SAM_ROTINACANCELAMENTO")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <> 1 Then
	bsShowMessage("Os parâmetros não podem estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "5" Then
	bsShowMessage("A rotina ainda não foi gerada", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAR").AsString = "5" Then
	bsShowMessage("A rotina já foi processada", "I")
	Exit Sub
  End If

  If VerificaDataFechamento = False Then
	Exit Sub
  End If

  Dim Obj As Object

  If VisibleMode Then

    Set Obj = CreateBennerObject("BSInterface0003.RotinaCancelamento")
    Obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Else

    Dim vsMensagemErro As String
    Dim viRetorno As Long

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen010", _
                                     "RotinaCancelamento_Processar", _
                                     "Rotina de Cancelamento de Beneficiários (Processar) -  Código: " + _
                                     CurrentQuery.FieldByName("CODIGO").AsString + " Descrição: " + _
                                     CurrentQuery.FieldByName("DESCRICAO").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACANCELAMENTO", _
                                     "SITUACAOPROCESSAR", _
                                     "SITUACAO", _
                                     "Geração não foi processada.", _
                                     "P", _
                                     False, _
                                     vsMensagemErro, _
                                     Null)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If

  End If

  Set Obj = Nothing

  WriteAudit("P", HandleOfTable("SAM_ROTINACANCELAMENTO"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
	  "Rotina de Cancelamento de Beneficiários - Processamento")
  RefreshNodesWithTable("SAM_ROTINACANCELAMENTO")
End Sub

Public Sub TABLE_AfterScroll()
  If VisibleMode Then
	CONTRATOFINAL.LocalWhere = "SAM_CONTRATO.CONTRATO >= " + _
							   "(Select CONTRATO FROM SAM_CONTRATO WHERE SAM_CONTRATO.HANDLE = @CONTRATOINICIAL)"
	FAMILIAFINAL.LocalWhere = "SAM_FAMILIA.FAMILIA >= " + _
							  "(SELECT FAMILIA FROM SAM_FAMILIA WHERE SAM_FAMILIA.HANDLE = @FAMILIAINICIAL)"
  Else
	CONTRATOFINAL.WebLocalWhere = "A.CONTRATO >= " + _
								  "(SELECT CONTRATO FROM SAM_CONTRATO WHERE HANDLE = @CAMPO(CONTRATOINICIAL))"
	FAMILIAFINAL.WebLocalWhere = "A.FAMILIA >= " + _
								 "(SELECT FAMILIA FROM SAM_FAMILIA WHERE HANDLE = @CAMPO(FAMILIAINICIAL))"

  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If CurrentUser = CurrentQuery.FieldByName("RESPONSAVEL").AsInteger Then
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Clear

	SQL.Add("DELETE FROM SAM_ROTINACANCELAMENTO_BENEF")
	SQL.Add("WHERE CANCELAMENTO = :HROTINACANCELAMENTO")

	SQL.ParamByName("HROTINACANCELAMENTO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

	SQL.ExecSQL

	Set SQL = Nothing
  Else
	CanContinue = False

	bsShowMessage("Operação cancelada. Usuário não é o Responsável", "E")
  End If
End Sub



Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qFechamento
  Set qFechamento = NewQuery

  qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")

  qFechamento.Active = True

  If CurrentQuery.State = 3 Then
	Dim vMesComp As Integer
	Dim vAnoComp As Integer
	Dim vMesFechamento As Integer
	Dim vAnoFechamento As Integer

	If CurrentQuery.FieldByName("tabdatacancelamento").AsInteger = 2 Then
	  vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIACANCELAMENTO").AsDateTime)
	  vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIACANCELAMENTO").AsDateTime)
	  vMesFechamento = DatePart("m", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
	  vAnoFechamento = DatePart("yyyy", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

	  If (vAnoComp < vAnoFechamento) Or (vAnoComp = vAnoFechamento And vMesComp < vMesFechamento) Then
		CanContinue = False
		bsShowMessage("A competência não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
	  End If
	Else
	  If CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime < qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
		bsShowMessage("Não é possível cancelar beneficiários com data de cancelamento inferior a data de fechamento - Parâmetros Gerais", "E")
		CanContinue = False
	  End If
	End If
  End If

  Set qFechamento = Nothing

  If Not CurrentQuery.FieldByName("DATABASEINADIMPLENCIA").IsNull Then
	If CurrentQuery.FieldByName("FATURASATRASO").IsNull Then
	  bsShowMessage("Preencha o número de faturas em atraso!", "E")
	  CanContinue = False
	End If
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
	Case "BOTAOEMITIRAVISO"
	  BOTAOEMITIRAVISO_OnClick
	Case "BOTAOEMITIRCOMUNICADO"
	  BOTAOEMITIRCOMUNICADO_OnClick
	Case "BOTAOGERAR"
	  BOTAOGERAR_OnClick
	Case "BOTAOPROCESSAR"
	  BOTAOPROCESSAR_OnClick
	Case "BOTAOENVIAREMAIL"
	  BOTAOENVIAREMAIL_OnClick
  End Select
End Sub
