'HASH: 6FED1DFB7188F18A5F48180210243555
Option Explicit

Public Sub ALTERARDATAPAGAMENTO_AfterOnClick()
	RefreshNodesWithTable("SAM_AGRUPADORPAGAMENTO")
End Sub

Public Sub EXCLUIRPAGAMENTO_AfterOnClick()
	RefreshNodesWithTable("SAM_AGRUPADORPAGAMENTO")
End Sub

Public Sub EXPORTARPREVIAPAGAMENTO_AfterOnClick()
	RefreshNodesWithTable("SAM_AGRUPADORPAGAMENTO")
End Sub

Public Sub FECHARPAGAMENTO_AfterOnClick()
	RefreshNodesWithTable("SAM_AGRUPADORPAGAMENTO")
End Sub

Public Sub REPROCESSARPAGAMENTO_AfterOnClick()
	RefreshNodesWithTable("SAM_AGRUPADORPAGAMENTO")
End Sub

Public Sub TABLE_AfterScroll()
	VerificarBotaoLiberarPagamento
	VerificarBotoesPagamentoAbertoFechado
	BOTAOCONCILIARDOCFISCAIS.Visible = VerificaPodeExibirBotaoConciliacaoDocumentosFiscais
	REPROCESSARPAGAMENTO.Visible = (CurrentQuery.FieldByName("STATUSPAGAMENTO").AsString <> "4") 'Não Faturado
	VerificarBotaoAtualizarPagamento
End Sub

Public Sub VerificarBotaoAtualizarPagamento()
	Dim bNotFiscal As Boolean
	Dim DocumentoFiscal As BPesquisa
	Set DocumentoFiscal = NewQuery
	DocumentoFiscal.Clear
	DocumentoFiscal.Active = False
	DocumentoFiscal.Add(" SELECT HANDLE, STATUSCONCILIACAO   ")
 	DocumentoFiscal.Add("   FROM SAM_DOCUMENTOFISCAL         ")
	DocumentoFiscal.Add("  WHERE PAGAMENTO = :PAGAMENTO      ")
	DocumentoFiscal.ParamByName("PAGAMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	DocumentoFiscal.Active = True

	bNotFiscal = (Not (DocumentoFiscal.EOF) And (DocumentoFiscal.FieldByName("STATUSCONCILIACAO").AsString = "2")) _
		Or Not ((CurrentQuery.FieldByName("USUARIOFECHAMENTO").IsNull) And (CurrentQuery.FieldByName("DATAHORAFECHAMENTO").IsNull))

	Set DocumentoFiscal = Nothing
	BOTAOATUALIZARPAGAMENTO.Visible = Not(bNotFiscal)
End Sub

Public Function VerificaPodeExibirBotaoConciliacaoDocumentosFiscais As Boolean
	If (CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
		Dim callEntity As CSEntityCall
	  	Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.Pagamentos.SamAgrupadorPagamento, Benner.Saude.Entidades", "PermiteExibirBotaoConciliarDocumentosFiscais")
	  	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
	  	VerificaPodeExibirBotaoConciliacaoDocumentosFiscais = CBool(callEntity.Execute)
		Set callEntity =  Nothing
	Else
		VerificaPodeExibirBotaoConciliacaoDocumentosFiscais = False
	End If
End Function

Public Sub VerificarBotoesPagamentoAbertoFechado()
	Dim bPagamentoAberto As Boolean
	bPagamentoAberto = ((CurrentQuery.FieldByName("USUARIOFECHAMENTO").IsNull) And (CurrentQuery.FieldByName("DATAHORAFECHAMENTO").IsNull))
	If bPagamentoAberto Then ' Pagamento Aberto
		FECHARPAGAMENTO.Visible = True
		EXCLUIRPAGAMENTO.Visible = True
		ALTERARDATAPAGAMENTO.Visible = True
		EXPORTARPREVIAPAGAMENTO.Visible = False
		BOTAOLIBERARTETOALCADA.Enabled = False
	Else ' Pagamento Fechado
		FECHARPAGAMENTO.Visible = False
		EXCLUIRPAGAMENTO.Visible = False
		ALTERARDATAPAGAMENTO.Visible = False
		EXPORTARPREVIAPAGAMENTO.Visible = True
		BOTAOLIBERARTETOALCADA.Enabled = (CurrentQuery.FieldByName("STATUSPAGAMENTO").AsInteger = 1)
	End If
End Sub

Public Sub VerificarBotaoLiberarPagamento()
	Dim bExigeLiberacao As Boolean
	Dim handleAgrup As String
	bExigeLiberacao = (UtilizaExigeLiberacaoGeral And UtilizaExigeLiberacaoPrestador)
	If Not(bExigeLiberacao) Then
		BOTAOLIBERARPAGAMENTO.Visible = bExigeLiberacao
		Exit Sub
	End If
	BOTAOLIBERARPAGAMENTO.Visible = bExigeLiberacao
	handleAgrup = CurrentQuery.FieldByName("HANDLE").AsString
	BOTAOLIBERARPAGAMENTO.Enabled = Not(VerificarExisteLiberacaoPagamento(handleAgrup))
End Sub

Public Function UtilizaExigeLiberacaoGeral As Boolean
	Dim oQry As BPesquisa
	Const sSim = "S"
	Set oQry = NewQuery
	oQry.Clear
	oQry.Add(" SELECT COALESCE(EXIGELIBERACAOPGTO,'N') AS EXIGELIBERACAOPGTO FROM SAM_PARAMETROSPROCCONTAS   ")
	oQry.Active = True
	UtilizaExigeLiberacaoGeral = (oQry.FieldByName("EXIGELIBERACAOPGTO").AsString = sSim)
	Set oQry = Nothing
End Function

Public Function UtilizaExigeLiberacaoPrestador As Boolean
	Dim oQry As BPesquisa
	Const sSim = "S"
	Set oQry = NewQuery
	oQry.Clear
	oQry.Add(" SELECT COALESCE(EXIGELIBERACAOPGTO,'N') AS EXIGELIBERACAOPGTO FROM SAM_PRESTADOR WHERE HANDLE = :PRESTADOR   ")
	oQry.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
	oQry.Active = True
	UtilizaExigeLiberacaoPrestador = (oQry.FieldByName("EXIGELIBERACAOPGTO").AsString = sSim)
	Set oQry = Nothing
End Function

Public Function VerificarExisteLiberacaoPagamento(handle As String) As Boolean
	Dim oQry As BPesquisa
	Set oQry = NewQuery
	oQry.Clear
	oQry.Add(" SELECT USUARIOLIBERACAOPAGAMENTO, DATALIBERACAOPAGAMENTO FROM SAM_AGRUPADORPAGAMENTO WHERE HANDLE = :HANDLE ")
	oQry.ParamByName("HANDLE").AsString = handle
	oQry.Active = True
	VerificarExisteLiberacaoPagamento = Not((oQry.FieldByName("USUARIOLIBERACAOPAGAMENTO").IsNull) And (oQry.FieldByName("DATALIBERACAOPAGAMENTO").IsNull))
	Set oQry = Nothing
End Function
