'HASH: B8711BB077158DE801BE8E14C9BC3F45
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim validaEmail As CSBusinessComponent

	Dim vsRetorno As String
	Dim vsRetornoRemetente As String
	Dim vsRetornoDestinatario As String

	Set validaEmail = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.TabelaVirtual.TvForm0148BLL, Benner.Saude.ProcessamentoContas.Business")

	validaEmail.AddParameter(pdtString, CurrentQuery.FieldByName("EMAILREMETENTE").AsString)
  	vsRetornoRemetente = validaEmail.Execute("ValidaEmailRemetentePrestador")

	validaEmail.ClearParameters
  	validaEmail.AddParameter(pdtString, CurrentQuery.FieldByName("EMAILDESTINATARIO").AsString)
  	vsRetornoDestinatario = validaEmail.Execute("ValidaEmailDestinatarioPrestador")

  	Set validaEmail = Nothing

	If (CurrentQuery.FieldByName("ASSUNTO").IsNull) Then
		vsRetorno = "O campo Assunto é obrigatório." + Chr(13) + Chr(10)
	End If

	If (CurrentQuery.FieldByName("MENSAGEM").IsNull) Then
		vsRetorno = vsRetorno + " - O campo Mensagem é obrigatório"
	End If

	If((vsRetornoRemetente <> "") Or (vsRetornoDestinatario <> "")) Then
		If (vsRetornoRemetente <> "") Then
			vsRetorno = vsRetorno + " - " + vsRetornoRemetente + Chr(13) + Chr(10)
		End If

		If (vsRetornoDestinatario <> "") Then
			vsRetorno = vsRetorno + " - " + vsRetornoDestinatario
		End If
	End If

	If (vsRetorno <> "") Then
 		bsshowmessage(vsRetorno, "I")

        CanContinue = False
		Exit Sub
	End If


	SessionVar("EMAILREMETENTE") = CurrentQuery.FieldByName("EMAILREMETENTE").AsString
	SessionVar("ASSUNTO") = CurrentQuery.FieldByName("ASSUNTO").AsString
	SessionVar("MENSAGEM") = CurrentQuery.FieldByName("MENSAGEM").AsString
	SessionVar("EMAILRECEBEDOR") = CurrentQuery.FieldByName("EMAILDESTINATARIO").AsString

End Sub

Public Sub TABLE_AfterScroll()

	Dim qDadosPegGuia   As BPesquisa
	Dim qMsgPadrao      As BPesquisa
	Dim qRemetente      As BPesquisa
	Dim vMotivoGlosa   As Long
    Dim vHandlePeg     As Long
	Dim vsMsgTraduzida   As String
	Dim TraduzMensagem As CSBusinessComponent

	vMotivoGlosa = CLng(SessionVar("MOTIVO"))
	Set qDadosPegGuia =NewQuery
	Set qMsgPadrao = NewQuery
	Set qRemetente = NewQuery

	vHandlePeg = CLng(SessionVar("HPEG"))
	qDadosPegGuia.Clear
   	qDadosPegGuia.Add("SELECT P.EMAIL")
	qDadosPegGuia.Add("  FROM SAM_PEG PEG")
	qDadosPegGuia.Add("  JOIN SAM_PRESTADOR P ON P.HANDLE = PEG.RECEBEDOR")
	qDadosPegGuia.Add(" WHERE PEG.HANDLE = :HPEG")
	qDadosPegGuia.ParamByName("HPEG").AsInteger = vHandlePeg
   	qDadosPegGuia.Active = True

	qMsgPadrao.Clear
    qMsgPadrao.Add("SELECT M.ASSUNTO")
	qMsgPadrao.Add("  FROM SAM_PARAMETROSPROCCONTAS P")
	qMsgPadrao.Add("  JOIN SAM_MENSAGEM_HTML M ON M.HANDLE = P.MENSAGEMDEVOLUCAO ")
    qMsgPadrao.Active = True

	qRemetente.Clear
    qRemetente.Add("SELECT EMAIL")
    qRemetente.Add("  FROM Z_GRUPOUSUARIOS")
    qRemetente.Add(" WHERE HANDLE = :HANDLE")
    qRemetente.ParamByName("HANDLE").AsInteger = CurrentUser
    qRemetente.Active = True

	Set TraduzMensagem = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.TabelaVirtual.TvForm0148BLL, Benner.Saude.ProcessamentoContas.Business")
	TraduzMensagem.AddParameter(pdtString, "P")
  	TraduzMensagem.AddParameter(pdtInteger, vHandlePeg)
  	TraduzMensagem.AddParameter(pdtInteger, vMotivoGlosa)
  	vsMsgTraduzida = TraduzMensagem.Execute("TraduzirMensagemPadrao")


	If (CurrentQuery.State <> 1) Then
		CurrentQuery.FieldByName("EMAILREMETENTE").AsString = qRemetente.FieldByName("EMAIL").AsString
		CurrentQuery.FieldByName("ASSUNTO").AsString = qMsgPadrao.FieldByName("ASSUNTO").AsString
		CurrentQuery.FieldByName("EMAILDESTINATARIO").AsString = qDadosPegGuia.FieldByName("EMAIL").AsString
		CurrentQuery.FieldByName("MENSAGEM").AsString = vsMsgTraduzida
	End If

  	Set TraduzMensagem = Nothing
	Set qDadosPegGuia = Nothing
	Set qMsgPadrao = Nothing
	Set qRemetente = Nothing

End Sub
