'HASH: 0589AF6C84C3A356A05C56A06B9606B7
'Macro: SAM_CONTRATO_PRECOREEMBDOT_GER
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"

Option Explicit


Public Sub BOTAOREPLICARPORCBOS_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  If (CurrentQuery.State <> 1)  Then
	  bsShowMessage("O registro não pode estar em edição. Confirme ou cancele as alterações!", "I")
	  Exit Sub
  End If

  SessionVar("TabelaDeDotacao") = "SAM_CONTRATO_PRECOREEMBDOT_GER"
  SessionVar("HandleDotac") = CurrentQuery.FieldByName("HANDLE").AsString

  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
								   1, _
								   "TV_CBOS", _
								   "Replicar por CBO-S", _
								   0, _
								   480, _
								   640, _
								   False, _
								   vsMensagem, _
								   vcContainer)

  RefreshNodesWithTable("SAM_CONTRATO_PRECOREEMBDOT_GER")
End Sub


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim vData As String
	Dim Interface As Object
	Dim vColunas,vCampos,vTabela As String

	Set Interface =CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, "", "Eventos", True, EVENTO.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value =vHandle
	End If

	Set Interface = Nothing
End Sub

Public Sub CBOSPESQUISA_OnPopup(ShowPopup As Boolean)
    Dim Interface As Object
    Dim vHandle As Long
    Dim vCampos As String
    Dim vColunas As String

    ShowPopup = False

    Set Interface = CreateBennerObject("Procura.Procurar")

    vColunas = "TIS_VERSAO.VERSAO|TIS_CBOS.CODIGO|TIS_CBOS.DESCRICAO"
    vCampos = "Versão TISS|Código do CBOS|Descrição do CBOS"
    vHandle = Interface.Exec(CurrentSystem,"TIS_CBOS|TIS_VERSAO[TIS_CBOS.VERSAOTISS = TIS_VERSAO.HANDLE]", vColunas, 2, vCampos, "", "CBOS", True, CBOSPESQUISA.Text)

    If (vHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("CBOSPESQUISA").Value = vHandle
	End If
	Set Interface = Nothing
End Sub

Public Sub TABELAFILME_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaFilme(TABELAFILME.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAFILME").Value = vHandle
	End If
End Sub

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaUS(TABELAUS.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAUS").Value = vHandle
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qCBOS   As BPesquisa
	Dim sMensagem As String

	If CurrentQuery.FieldByName("CBOSPESQUISA").IsNull Then
		CurrentQuery.FieldByName("CBOS").Clear
	Else
		Set qCBOS = NewQuery
		qCBOS.Add("SELECT CODIGO FROM TIS_CBOS WHERE HANDLE = :HANDLE")
		qCBOS.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CBOSPESQUISA").Value
		qCBOS.Active = True
		CurrentQuery.FieldByName("CBOS").Value = qCBOS.FieldByName("CODIGO").Value
		Set qCBOS = Nothing
	End If

	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		If CurrentQuery.FieldByName("DATAINICIAL").Value > CurrentQuery.FieldByName("DATAFINAL").Value Then
			bsShowMessage("A Data Inicial não pode ser maior que a Data Final", "E")
			CanContinue = False
		End If
	End If

	sMensagem = VerificarCBOSNaVigencia
	If sMensagem <> "" Then
		bsShowMessage(sMensagem, "E")
		CanContinue = False
	End If

	sMensagem = VerificarEventoNaVigencia
	If sMensagem <> "" Then
		bsShowMessage(sMensagem, "E")
		CanContinue = False
	End If
End Sub

Public Function VerificarEventoNaVigencia As String
	Dim sContrato As String
	Dim sEvento As String
	Dim sHandle As String
	Dim sCondicao As String
	Dim sLinha As String

	If Not(CurrentQuery.FieldByName("CBOS").IsNull) Then
	  Exit Function
	End If

	Dim Interface As Object
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	sContrato = CurrentQuery.FieldByName("CONTRATO").Value
	sEvento = CurrentQuery.FieldByName("EVENTO").Value
	sHandle = CurrentQuery.FieldByName("HANDLE").Value

	sCondicao = " AND CONTRATO = " + sContrato
	sCondicao = sCondicao + " AND EVENTO = " + sEvento
	sCondicao = sCondicao + " AND (CBOS IS NULL OR CBOS = '')"
	sCondicao = sCondicao + " AND HANDLE <> " + sHandle


	sLinha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_PRECOREEMBDOT_GER", _
		"DATAINICIAL", "DATAFINAL", _
		CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
		"", sCondicao _
	)

  VerificarEventoNaVigencia = sLinha

End Function

Public Function VerificarCBOSNaVigencia As String
	Dim sContrato As String
	Dim sCBOS As String
	Dim sEvento As String
	Dim sHandle As String
	Dim sCondicao As String
	Dim sLinha As String

	If CurrentQuery.FieldByName("CBOS").IsNull Then
	  Exit Function
	End If

	Dim Interface As Object
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	sContrato = CurrentQuery.FieldByName("CONTRATO").Value
	sEvento = CurrentQuery.FieldByName("EVENTO").Value
	sHandle = CurrentQuery.FieldByName("HANDLE").Value
	sCBOS = CurrentQuery.FieldByName("CBOS").Value

	sCondicao = " AND CONTRATO = " + sContrato
	sCondicao = sCondicao + " AND EVENTO = " + sEvento
	sCondicao = sCondicao + " AND COALESCE(CBOS,'') = '" + sCBOS + "' "
	sCondicao = sCondicao + " AND HANDLE <> " + sHandle

	sLinha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_PRECOREEMBDOT_GER", _
		"DATAINICIAL", "DATAFINAL", _
		CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
		"", sCondicao _
	)

	VerificarCBOSNaVigencia = sLinha

End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOREPLICARPORCBOS"
			ArmazenarSecao
	End Select
End Sub

Public Sub ArmazenarSecao
  SessionVar("TabelaDeDotacao") = "SAM_CONTRATO_PRECOREEMBDOT_GER"
  SessionVar("HandleDotac") = CurrentQuery.FieldByName("HANDLE").AsString
End Sub
