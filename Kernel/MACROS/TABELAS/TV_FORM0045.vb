'HASH: E2030B6D5DC862D2F4F10BAEBA6AA33A
'#uses "*CriaTabelaTemporariaSqlServer"
'#uses "*bsShowMessage"
Option Explicit


Public Sub TABLE_AfterInsert()
	comum
End Sub

Public Sub TABLE_AfterPost()
	Dim dll As Object
	Dim resultado As Integer
	Dim msg As String
	Set dll = CreateBennerObject("ca043.autorizacao")
	If pegarTipoAcomodacao="E" Then
		resultado = dll.gerarDiariasAfterPost(CurrentSystem, CurrentQuery.FieldByName("acomodacaoEvento").AsInteger, CurrentQuery.FieldByName("CODIGOTABELA").AsInteger, msg)
	Else
		resultado = dll.gerarDiariasAfterPost(CurrentSystem, CurrentQuery.FieldByName("acomodacaoGrau").AsInteger, CurrentQuery.FieldByName("CODIGOTABELA").AsInteger, msg)
	End If

	Set dll=Nothing

	'abrir a transação, pois o processo fecha a transação e neste ponto o runner espera uma transação aberta
	If Not InTransaction Then
		StartTransaction
	End If

	If (msg <> "") Then
		Err.Raise(vbsUserException, "", msg)
    Else
      If WebMode Then
		Dim spBSAUT_GERARPROTOCOLOTRANSACAO As BStoredProc
		Set spBSAUT_GERARPROTOCOLOTRANSACAO = NewStoredProc
		spBSAUT_GERARPROTOCOLOTRANSACAO.AutoMode = True
		spBSAUT_GERARPROTOCOLOTRANSACAO.Name = "BSAUT_GERARPROTOCOLOTRANSACAO"
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_ORIGEMPROCESSO",ptInput, ftString)
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_AUTORIZACAO",ptInput, ftInteger)
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_USUARIO",ptInput, ftInteger)
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_NUMEROPROTOCOLOTRANSACAO",ptInputOutput, ftInteger)
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_HANDLEATENDIMENTOCENTRAL",ptInputOutput, ftInteger)

		spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_ORIGEMPROCESSO").AsString = "D" ' simulando como desktop, mesmo sendo no ambiente web, para localizar o tipo dele (complemento, prorrogação, e outros)
		spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_AUTORIZACAO").AsInteger = CLng(SessionVar("HANDLEAUTORIZACAO"))
		spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_USUARIO").AsInteger = CurrentUser

		If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then
		  Dim qBuscaProtocoloPrincipal As Object
		  Set qBuscaProtocoloPrincipal = NewQuery

		  qBuscaProtocoloPrincipal.Clear
		  qBuscaProtocoloPrincipal.Add("SELECT MIN(P1.HANDLE) PRIMEIROPROTOCOLO ")
		  qBuscaProtocoloPrincipal.Add("  FROM SAM_PROTOCOLOTRANSACAOAUTORIZ P1")
		  qBuscaProtocoloPrincipal.Add("  JOIN SAM_PROTOCOLOTRANSACAOAUTORIZ P2 ON P1.AUTORIZACAO = P2.AUTORIZACAO")
		  qBuscaProtocoloPrincipal.Add(" WHERE P2.HANDLE = :HANDLE")
		  qBuscaProtocoloPrincipal.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
		  qBuscaProtocoloPrincipal.Active = True

		  spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsInteger = qBuscaProtocoloPrincipal.FieldByName("PRIMEIROPROTOCOLO").AsInteger

		  Set qBuscaProtocoloPrincipal = Nothing
		Else
		  spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsInteger = 0
		End If

		spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_HANDLEATENDIMENTOCENTRAL").AsInteger = 0

		spBSAUT_GERARPROTOCOLOTRANSACAO.ExecProc
		Set spBSAUT_GERARPROTOCOLOTRANSACAO = Nothing
      End If
	End If

End Sub

Public Sub TABLE_AfterScroll()

	If WebMode Then
		CODIGOTABELA.WebLocalWhere = "A.VERSAOTISS IN (SELECT HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
	Else
 		CODIGOTABELA.LocalWhere = "TIS_TABELAPRECO.VERSAOTISS IN (SELECT HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	'ESTE BEFOREINSERT SOMENTE NA WEB

	Dim qParametro As Object
	Set qParametro = NewQuery

	qParametro.Add("SELECT SOLICITATABELAPRECODIARIA FROM SAM_PARAMETROSATENDIMENTO")
	qParametro.Active = True

	If qParametro.FieldByName("SOLICITATABELAPRECODIARIA").AsString = "N" Then
		CODIGOTABELA.Visible = False
	Else
		CODIGOTABELA.Visible = True
	End If

	Set qParametro = Nothing

	If WebMode Then
		Dim dll As Object
		Dim resultado As Integer
		Dim acomodacoes As String
		Dim msg As String
		Set dll = CreateBennerObject("ca043.autorizacao")

		resultado = dll.GerarDiariasNewRecord(CurrentSystem, acomodacoes, msg)
		Set dll=Nothing

		If pegarTipoAcomodacao = "E" Then
			ACOMODACAOEVENTO.ReadOnly = False
			ACOMODACAOGRAU.ReadOnly   = True
			If acomodacoes<>"" Then
				ACOMODACAOEVENTO.WebLocalWhere="A.HANDLE IN ("+ acomodacoes+") AND (A.EVENTO IN (SELECT EVENTO FROM SAM_TGE_TABELATISS WHERE TABELATISS = @CAMPO(CODIGOTABELA)) OR (@CAMPO(CODIGOTABELA) = -1) )"
			Else
				ACOMODACAOEVENTO.WebLocalWhere="A.HANDLE = -1"
			End If
		Else
			ACOMODACAOEVENTO.ReadOnly =True
			ACOMODACAOGRAU.ReadOnly   =False
			If acomodacoes<>"" Then
				ACOMODACAOGRAU.WebLocalWhere="A.HANDLE IN ("+ acomodacoes+")"
			Else
				ACOMODACAOGRAU.WebLocalWhere="A.HANDLE = -1"
			End If
		End If


		If (msg<>"") Then
			CanContinue=False
			bsShowMessage(msg, "E")
			Exit Sub
		End If

		If (acomodacoes="") Then
			CanContinue=False
			bsShowMessage("Não existem acomodações para escolher", "E")
			Exit Sub
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim vExibirMensagem As String
	Dim sql As BPesquisa
	Set sql = NewQuery
	sql.Add("SELECT SOLICITATABELAPRECODIARIA, EXIBIRMENSAGEMDIASINTERNACAO FROM SAM_PARAMETROSATENDIMENTO")
	sql.Active = True

    vExibirMensagem = sql.FieldByName("EXIBIRMENSAGEMDIASINTERNACAO").AsString

	If sql.FieldByName("SOLICITATABELAPRECODIARIA").AsString  = "S" And CurrentQuery.FieldByName("CODIGOTABELA").IsNull Then
			bsShowMessage("Parâmetro atendimento está marcado para informar tabela de preço.", "E")
			CanContinue = False
			Set sql = Nothing
			Exit Sub
	End If


	If (CurrentQuery.FieldByName("CODIGOTABELA").AsInteger > 0) And (CurrentQuery.FieldByName("ACOMODACAOEVENTO").AsInteger > 0) Then

		Dim qEvento As Object
		Set qEvento = NewQuery

		qEvento.Add("SELECT EVENTO               ")
		qEvento.Add("  FROM SAM_ACOMODACAO_EVENTO")
		qEvento.Add(" WHERE HANDLE =:HANDLE")
		qEvento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ACOMODACAOEVENTO").AsInteger
        qEvento.Active = True

		Dim qTabTiss As Object
		Set qTabTiss = NewQuery
		qTabTiss.Add("SELECT HANDLE FROM SAM_TGE_TABELATISS WHERE EVENTO = :EVENTO AND TABELATISS = :TABELATISS")
		qTabTiss.ParamByName("EVENTO").AsInteger = qEvento.FieldByName("EVENTO").AsInteger
		qTabTiss.ParamByName("TABELATISS").AsInteger = CurrentQuery.FieldByName("CODIGOTABELA").AsInteger
		qTabTiss.Active = True

		If qTabTiss.EOF Then
			bsShowMessage("Evento incompatível com o código tabela selecionado.", "E")
			CanContinue = False
			Set qTabTiss = Nothing
			Exit Sub
		End If

		Set qTabTiss = Nothing
		Set qEvento = Nothing

	End If

	If vExibirMensagem = "S" Then
		If CurrentQuery.FieldByName("qtdsolicitada").AsInteger > CurrentQuery.FieldByName("qtdtge").AsInteger Then
			bsShowMessage("Quantidade de dias solicitados para internação maior que o máximo definido na TGE", "I")
		End If

		If CurrentQuery.FieldByName("qtdliberada").AsInteger > CurrentQuery.FieldByName("qtdtge").AsInteger Then
			bsShowMessage("Quantidade de dias liberados para internação maior que o máximo definido na TGE", "I")
		End If
	End If

	' gravar na autorizacao os valores digitados
	sql.Clear
	sql.Add("UPDATE SAM_AUTORIZ SET DIARIASLIBERADAS=:L, DIARIASSOLICITADAS=:S WHERE HANDLE=:H")
	sql.ParamByName("H").AsFloat = CLng(SessionVar("HANDLEAUTORIZACAO"))
	sql.ParamByName("L").AsInteger = CurrentQuery.FieldByName("QTDLIBERADA").AsInteger
	sql.ParamByName("S").AsInteger = CurrentQuery.FieldByName("QTDSOLICITADA").AsInteger
	sql.ExecSQL
	Set sql = Nothing
End Sub

Public Sub TABLE_NewRecord()
	'ESTE NEW RECORD SOMENTE NO DESKTOP
	CriaTabelaTemporariaSqlServer
	'comum
	If VisibleMode Then
		Dim dll As Object
		Dim resultado As Integer
		Dim acomodacoes As String
		Dim msg As String
		Set dll = CreateBennerObject("ca043.autorizacao")

		resultado = dll.GerarDiariasNewRecord(CurrentSystem, acomodacoes, msg)
		Set dll=Nothing

		If pegarTipoAcomodacao = "E" Then
			ACOMODACAOEVENTO.Visible=True
			ACOMODACAOGRAU.Visible=False
			If acomodacoes<>"" Then
				ACOMODACAOEVENTO.LocalWhere="SAM_ACOMODACAO_EVENTO.HANDLE IN ("+ acomodacoes+") AND (SAM_ACOMODACAO_EVENTO.EVENTO IN (SELECT EVENTO FROM SAM_TGE_TABELATISS WHERE TABELATISS = @CODIGOTABELA) OR (@CODIGOTABELA = -1))"
			Else
				ACOMODACAOEVENTO.LocalWhere="SAM_ACOMODACAO_EVENTO.HANDLE = -1"
			End If
		Else
			ACOMODACAOEVENTO.Visible=False
			ACOMODACAOGRAU.Visible=True
			If acomodacoes<>"" Then
				ACOMODACAOGRAU.LocalWhere="SAM_ACOMODACAO_GRAU.HANDLE IN ("+ acomodacoes+")"
			Else
				ACOMODACAOGRAU.LocalWhere="SAM_ACOMODACAO_GRAU.HANDLE = -1"
			End If
		End If

		If (resultado > 0) Then
			bsShowMessage(msg, "E")
		End If
	End If
End Sub

Public Sub comum
	Dim sql As BPesquisa
	Set sql = NewQuery
	sql.Add("SELECT DIARIASMAXTGE, DIARIASSOLICITADAS, DIARIASLIBERADAS FROM SAM_AUTORIZ WHERE HANDLE=:H")
	sql.ParamByName("H").AsFloat=CLng(SessionVar("HANDLEAUTORIZACAO"))
	sql.Active=True
	CurrentQuery.FieldByName("qtdtge").AsInteger = sql.FieldByName("diariasmaxtge").AsInteger
	CurrentQuery.FieldByName("QTDLIBERADA").AsInteger = sql.FieldByName("DIARIASLIBERADAS").AsInteger
	CurrentQuery.FieldByName("QTDSOLICITADA").AsInteger = sql.FieldByName("DIARIASSOLICITADAS").AsInteger
	Set sql=Nothing
End Sub


Public Function pegarTipoAcomodacao As String
	Dim sql As BPesquisa
	Set sql=NewQuery
	sql.Add("SELECT TIPOACOMODACAO FROM SAM_PARAMETROSATENDIMENTO")
	sql.Active=True
	pegarTipoAcomodacao = sql.FieldByName("TIPOACOMODACAO").AsString
	Set sql=Nothing
End Function
