'HASH: 9A3518DEE193566215C804089D433CD8
'#uses "*bsShowMessage"
Dim SitProcessamsento As String

Public Sub ROTINAINICIAL_OnPopup(ShowPopup As Boolean)
	SitProcessamento = Solver(RecordHandleOfTable("SAM_ROTINARECALCULOMENSALID"), "SAM_ROTINARECALCULOMENSALID", "SITUACAOPROCESSAMENTO")
	ROTINAINICIAL.LocalWhere = "SITUACAOPROCESSAMENTO = '" + SitProcessamento + "' AND SITUACAOFATURAMENTO = '1'"
End Sub

Public Sub ROTINAFINAL_OnPopup(ShowPopup As Boolean)
	SitProcessamento = Solver(RecordHandleOfTable("SAM_ROTINARECALCULOMENSALID"), "SAM_ROTINARECALCULOMENSALID", "SITUACAOPROCESSAMENTO")
	ROTINAFINAL.LocalWhere = "SITUACAOPROCESSAMENTO = '" + SitProcessamento + "' AND SITUACAOFATURAMENTO = '1' AND CODIGO >= (SELECT CODIGO FROM SAM_ROTINARECALCULOMENSALID WHERE HANDLE = @CAMPO(ROTINAINICIAL))"
End Sub

Public Sub TABLE_AfterPost()
	If VisibleMode Then
		Exit Sub
	End If

	Dim TituloRotina As String

	SitProcessamento = Solver(RecordHandleOfTable("SAM_ROTINARECALCULOMENSALID"), "SAM_ROTINARECALCULOMENSALID", "SITUACAOPROCESSAMENTO")
	Select Case (SitProcessamento)
		Case "1"
			SitProcessamento = "RotinaProcessamentoRecalculo_ProcessarRotinas"
			TituloRotina = "Rotina Processamento - Recálculo de Mensalidade"
		Case "5"
			SitProcessamento = "RotinaFaturamentoRecalculo_ProcessarRotinas"
			TituloRotina = "Rotina Faturamento - Recálculo de Mensalidade"
	End Select


	Dim viRetorno As Integer
	Dim vsMensagemErro As String
	Dim obj As Object
	Dim vcContainer As CSDContainer

	Set vcContainer = NewContainer
	vcContainer.GetFieldsFromQuery(CurrentQuery.TQuery)
	vcContainer.LoadAllFromQuery(CurrentQuery.TQuery)

	Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRetorno = obj.ExecucaoImediata(CurrentSystem, _
									 "SamRecalcMensal", _
									 SitProcessamento, _
									 TituloRotina + _
									 " - Rotina Inicial: " + CStr(CurrentQuery.FieldByName("ROTINAINICIAL").AsInteger) + _
									 " - Rotina Final: " + CStr(CurrentQuery.FieldByName("ROTINAFINAL").AsInteger), _
									 0, _
									 "", _
									 "", _
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

	Set obj = Nothing
	Set vcContainer = Nothing
End Sub
