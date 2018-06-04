'HASH: 64D0CE2D22816AAC0D5C0B653D1CDEB4
'#Uses "*bsShowMessage"
'#Uses "*VerificarBloqueioAlteracoes"
'#Uses "*VerificarBloqueioAlteracoesReapresentacao"
'#Uses "*RecordHandleOfTableInterfacePEG"

Public Sub TABLE_AfterInsert()

	Dim qAux As BPesquisa

	Set qAux = NewQuery
	qAux.Clear
	qAux.Add("SELECT * FROM SAM_GUIA WHERE HANDLE = :HANDLE")
	qAux.Active = False

	If SessionVar("HANDLEGUIA") <> "" Then
      qAux.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HANDLEGUIA"))
      SessionVar("HANDLEGUIA") = ""
	Else
      qAux.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_GUIA")
   	End If

	qAux.Active = True

	Select Case qAux.FieldByName("TABREGIMEPGTO").AsInteger
		Case 1
			CurrentQuery.FieldByName("TABDEVOLUCAO").AsInteger = 1

			If Not qAux.FieldByName("RECEBEDOR").IsNull Then
				CurrentQuery.FieldByName("PRESTADOR").AsInteger = qAux.FieldByName("RECEBEDOR").AsInteger
			End If

			If Not qAux.FieldByName("BENEFICIARIO").IsNull Then
				CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = qAux.FieldByName("BENEFICIARIO").AsInteger
			End If

		Case 2
			CurrentQuery.FieldByName("TABDEVOLUCAO").AsInteger = 2

			If Not qAux.FieldByName("BENEFICIARIO").IsNull Then
				CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = qAux.FieldByName("BENEFICIARIO").AsInteger
			End If

	End Select

	CurrentQuery.FieldByName("PEG").AsInteger = qAux.FieldByName("PEG").AsInteger
	CurrentQuery.FieldByName("GUIA").AsInteger = qAux.FieldByName("HANDLE").AsInteger
	CurrentQuery.FieldByName("PRESTADOR").AsInteger = qAux.FieldByName("RECEBEDOR").AsInteger
	CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser

	Set qAux = Nothing

End Sub

Public Sub TABLE_AfterScroll()

	Dim qSql As BPesquisa
	Set qSql = NewQuery
	Dim qPeg As BPesquisa
	Set qPeg = NewQuery

	qPeg.Clear
	qPeg.Add("SELECT TABREGIMEPGTO FROM SAM_PEG WHERE HANDLE = :HANDLE")
	qPeg.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
	qPeg.Active = True

	qSql.Add("SELECT TABCOMUNICADEVOLUCAO FROM SAM_PARAMETROSPROCCONTAS ")
	qSql.Active = True

	If ((qPeg.FieldByName("TABREGIMEPGTO").AsInteger = 1) And _
	    (qSql.FieldByName("TABCOMUNICADEVOLUCAO").AsInteger = 1)) Then

		TABLE.TabVisible(1) = True

		Dim qDadosGuia As BPesquisa
		Dim qRemetente As BPesquisa
		Dim qMsgPadrao As BPesquisa
		Dim vsMsgTraduzida As String

		Set qDadosGuia = NewQuery
		Set qRemetente = NewQuery
		Set qMsgPadrao = NewQuery

		qDadosGuia.Clear
    	qDadosGuia.Add("SELECT P.EMAIL")
		qDadosGuia.Add("  FROM SAM_GUIA G")
		qDadosGuia.Add("  JOIN SAM_PRESTADOR P ON P.HANDLE = G.RECEBEDOR")
		qDadosGuia.Add(" WHERE G.HANDLE = :HGUIA")
		qDadosGuia.ParamByName("HGUIA").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
    	qDadosGuia.Active = True

		qRemetente.Clear
		qRemetente.Add("SELECT EMAIL")
		qRemetente.Add("  FROM Z_GRUPOUSUARIOS")
		qRemetente.Add(" WHERE HANDLE = :HANDLE")
		qRemetente.ParamByName("HANDLE").AsInteger = CurrentUser
    	qRemetente.Active = True

    	qMsgPadrao.Clear
    	qMsgPadrao.Add("SELECT M.ASSUNTO")
		qMsgPadrao.Add("  FROM SAM_PARAMETROSPROCCONTAS P")
		qMsgPadrao.Add("  JOIN SAM_MENSAGEM_HTML M ON M.HANDLE = P.MENSAGEMDEVOLUCAOGUIA ")
    	qMsgPadrao.Active = True

  		vsMsgTraduzida = TraduzirMensagemHtml

		If (CurrentQuery.State <> 1) Then
		    CurrentQuery.FieldByName("EMAILREMETENTE").AsString = qRemetente.FieldByName("EMAIL").AsString
			CurrentQuery.FieldByName("ASSUNTO").AsString = qMsgPadrao.FieldByName("ASSUNTO").AsString
			CurrentQuery.FieldByName("EMAILDESTINATARIO").AsString = qDadosGuia.FieldByName("EMAIL").AsString
			CurrentQuery.FieldByName("MENSAGEM").AsString = vsMsgTraduzida
		End If
	Else
    	TABLE.TabVisible(1) = False
	End If

	SessionVar("GUIAHANDLE") = ""
	Set qMsgPadrao = Nothing
	Set qDadosGuia = Nothing
	Set qSql = Nothing
	Set qPeg = Nothing

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If VerificarBloqueioAlteracoesReapresentacao(RecordHandleOfTableInterfacePEG("SAM_PEG")) Then
      bsShowMessage("A GUIA não pode ser devolvida porque o PEG de reapresentação não pode ser alterado. ", "E")
      CanContinue = False
	  Exit Sub
  End If

  If VerificarBloqueioAlteracoes(RecordHandleOfTableInterfacePEG("SAM_PEG")) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG está vinculado a um agrupador de pagamento com documentos fiscais conciliados. ", "E")
    CanContinue = False
	Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If CurrentQuery.FieldByName("MOTIVOGLOSA").IsNull Then
		BsShowMessage("O motivo de glosa deve ser informado", "I")
		MOTIVOGLOSA.SetFocus
		CanContinue = False
		Exit Sub
	End If

	CurrentQuery.FieldByName("MENSAGEM").AsString = TraduzirMensagemHtml

	Dim qPeg As BPesquisa
	Set qPeg = NewQuery
	Dim qSql As BPesquisa
	Set qSql = NewQuery

	qPeg.Clear
	qPeg.Add("SELECT TABREGIMEPGTO FROM SAM_PEG WHERE HANDLE = :HANDLE")
	qPeg.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
	qPeg.Active = True

	qSql.Add("SELECT TABCOMUNICADEVOLUCAO FROM SAM_PARAMETROSPROCCONTAS ")
	qSql.Active = True

    If ((qPeg.FieldByName("TABREGIMEPGTO").AsInteger = 1) And _
	    (qSql.FieldByName("TABCOMUNICADEVOLUCAO").AsInteger = 1)) Then

	    Dim mensagemValidacao As String
	    Dim ValidaEmail As CSBusinessComponent
		Dim vsRetorno As String

		If CurrentQuery.FieldByName("ASSUNTO").IsNull Then
			vsRetorno = " - O campo Assunto é obrigatório." + Chr(13) + Chr(10)
	  	End If


		If CurrentQuery.FieldByName("MENSAGEM").IsNull Then
			vsRetorno = vsRetorno + " - O campo Mensagem é obrigatório." + Chr(13) + Chr(10)
	  	End If

		Set ValidaEmail = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.TabelaVirtual.TvForm0148BLL, Benner.Saude.ProcessamentoContas.Business")
		ValidaEmail.AddParameter(pdtString, CurrentQuery.FieldByName("EMAILREMETENTE").AsString)
	  	mensagemValidacao = ValidaEmail.Execute("ValidaEmailRemetentePrestador")

		If (mensagemValidacao <> "") Then
			vsRetorno = vsRetorno + " - " + mensagemValidacao + Chr(13) + Chr(10)
			mensagemValidacao = ""
		End If

		ValidaEmail.ClearParameters
		ValidaEmail.AddParameter(pdtString, CurrentQuery.FieldByName("EMAILDESTINATARIO").AsString)
	  	mensagemValidacao = ValidaEmail.Execute("ValidaEmailDestinatarioPrestador")

	  	Set ValidaEmail = Nothing

	  	If (mensagemValidacao <> "") Then
			vsRetorno = vsRetorno + " - " + mensagemValidacao + Chr(13) + Chr(10)
			mensagemValidacao = ""
		End If

		If vsRetorno <> "" Then
			bsShowMessage(vsRetorno, "E")
			CanContinue = False
			Exit Sub
		End If

		SessionVar("EMAILREMETENTE") = CurrentQuery.FieldByName("EMAILREMETENTE").AsString
		SessionVar("ASSUNTO") = CurrentQuery.FieldByName("ASSUNTO").AsString
		SessionVar("MENSAGEM") = CurrentQuery.FieldByName("MENSAGEM").AsString
		SessionVar("EMAILRECEBEDOR") = CurrentQuery.FieldByName("EMAILDESTINATARIO").AsString
		SessionVar("ORIGEM") = "G"

    End If

	Set qPeg = Nothing
	Set qSql = Nothing

	Dim vsMsg As String

	If VisibleMode Then

		If CurrentQuery.FieldByName("DEVOLVERPEG").AsString = "S" Then
			vbApagarPeg = True
		Else
			vbApagarPeg = False
		End If

		If SessionVar("HANDLEGUIA") <> "" Then
	      viHandleGuia = CLng(SessionVar("HANDLEGUIA"))
	      SessionVar("HANDLEGUIA") = ""
	   	End If

		Set vvSamGuiaDev = CreateBennerObject("SAMDEVOLUCAOGUIA.Rotinas")
		vsMsg = vvSamGuiaDev.DevolverGuia(CurrentSystem, _
										  CurrentQuery.FieldByName("GUIA").AsInteger, _
										  CurrentQuery.FieldByName("TABDEVOLUCAO").AsInteger, _
										  CurrentQuery.FieldByName("PRESTADOR").AsInteger, _
										  CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, _
										  CurrentQuery.FieldByName("ESTADO").AsInteger, _
										  CurrentQuery.FieldByName("MUNICIPIO").AsInteger, _
										  CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger, _
										  CurrentQuery.FieldByName("COMPLEMENTO").AsString, _
										  vbApagarPeg, _
										  CurrentQuery.FieldByName("ACEITAREGULARIZACAO").AsString, _
										  CurrentQuery.FieldByName("GUIAPROCESSADA").AsString)

		If vsMsg <> "" Then
			bsShowMessage(vsMsg, "E")
			CanContinue = False
			Exit Sub
		End If
	End If

End Sub

Public Function TraduzirMensagemHtml As String
    TraduzirMensagemHtml = ""
	Dim TraduzMensagem As CSBusinessComponent

	Set TraduzMensagem = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.TabelaVirtual.TvForm0148BLL, Benner.Saude.ProcessamentoContas.Business")
	TraduzMensagem.AddParameter(pdtString, "G")
  	TraduzMensagem.AddParameter(pdtInteger, CurrentQuery.FieldByName("GUIA").AsInteger)
  	TraduzMensagem.AddParameter(pdtInteger, CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger)
  	TraduzMensagem.AddParameter(pdtString, CurrentQuery.FieldByName("MENSAGEM").AsString)

  	TraduzirMensagemHtml = TraduzMensagem.Execute("TraduzirMensagemPadrao")

  	Set TraduzMensagem = Nothing
End Function
