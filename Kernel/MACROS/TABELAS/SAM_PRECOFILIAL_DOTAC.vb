'HASH: 0051FC5A85DF88C0B87C8B41D8BD2172
'Macro: SAM_PRECOFILIAL_DOTAC
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"

Option Explicit

Public Sub BOTAOPRECO_OnClick()
  Dim vDataBaseChecagemVigencia As Date


  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > ServerDate Then
    vDataBaseChecagemVigencia = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  Else
	If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
	  vDataBaseChecagemVigencia = ServerDate
	Else
	  vDataBaseChecagemVigencia = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
	End If
  End If


  Dim Interface As Object
	Dim ValorEvento As Currency
	Dim SQL As Object
	Dim Nivel As Integer
	Dim result As String
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_CONFIGURABUSCAPRECO")

	SQL.Active = True

	Nivel = -1

	If SQL.FieldByName("NIVEL1").AsInteger = 8 Then
		Nivel = 1
	ElseIf SQL.FieldByName("NIVEL2").AsInteger = 8 Then
		Nivel = 2
	ElseIf SQL.FieldByName("NIVEL3").AsInteger = 8 Then
		Nivel = 3
	ElseIf SQL.FieldByName("NIVEL4").AsInteger = 8 Then
		Nivel = 4
	ElseIf SQL.FieldByName("NIVEL5").AsInteger = 8 Then
		Nivel = 5
	ElseIf SQL.FieldByName("NIVEL6").AsInteger = 8 Then
		Nivel = 6
	ElseIf SQL.FieldByName("NIVEL7").AsInteger = 8 Then
		Nivel = 7
	ElseIf SQL.FieldByName("NIVEL8").AsInteger = 8 Then
		Nivel = 8
	End If

	If Nivel <> -1 Then
		If CurrentQuery.FieldByName("EVENTO").IsNull Then
			result = ""
		Else
			Set Interface = CreateBennerObject("BSPRE001.Rotinas")

			ValorEvento = Interface.ValorEvento(CurrentSystem, vDataBaseChecagemVigencia, 99, -1, -1, -1, -1, CurrentQuery.FieldByName("FILIAL").Value, -1, -1, -1, CurrentQuery.FieldByName("EVENTO").Value, -1, Nivel, CurrentQuery.FieldByName("CONVENIO").AsInteger, "", CurrentQuery.FieldByName("CBOS").AsString)

			result = "Valor do evento nesta vigência: R$ " + Format(ValorEvento, "#,##0.0000")+" ("+Format(ValorEvento,"#,##0.00")+")"
		End If
	Else
		result = "Na configuração de busca de preço, não foi definido um nível para a Filial!"
	End If
	If VisibleMode Then
		LABELPRECO.Text = result
	Else
		If Nivel <> -1 Then
			bsShowMessage(result, "I")
		Else
			bsShowMessage(result, "E")
		End If
	End If

End Sub

Public Sub BOTAOQUANTIDADE_OnClick()
	'****************************** SMS 13737 DURVAL 26/12/2002 ******************************************************************
	Dim QueryRetorno As Object
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String

	If CurrentQuery.State <> 3 Then
		Exit Sub
	End If

	If CurrentQuery.FieldByName("EVENTO").IsNull Then
		bsShowMessage("Digite um evento !", "I")
		Exit Sub
	End If

	Set Interface = CreateBennerObject("Procura.Procurar")

	vColunas = "SAM_PRECOGENERICO.DESCRICAO|SAM_TGE.DESCRICAO|SAM_PRECOGENERICO_DOTAC.QTDUSHONORARIO|SAM_PRECOGENERICO_DOTAC.QTDUSCUSTOOPERACIONAL|SAM_PRECOGENERICO_DOTAC.FATORFILME|SAM_PRECOGENERICO_DOTAC.PORTEANESTESICO|SAM_PRECOGENERICO_DOTAC.PORTESALA"
	vCriterio = "((SAM_PRECOGENERICO_DOTAC.DATAINICIAL <= " + SQLDate(ServerDate) + ") AND (SAM_PRECOGENERICO_DOTAC.DATAFINAL >= " + SQLDate(ServerDate) + " OR SAM_PRECOGENERICO_DOTAC.DATAFINAL IS NULL))"
	vCriterio = vCriterio + " AND SAM_PRECOGENERICO_DOTAC.EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	vCampos = "Tabela Genérica|Evento|Qtde US Honorário|Qtde US Custo Operacional|Fator de filme|Porte Anestésico|Porte de Sala"
	vHandle = Interface.Exec(CurrentSystem, "SAM_PRECOGENERICO_DOTAC|SAM_PRECOGENERICO[SAM_PRECOGENERICO_DOTAC.TABELAPRECO = SAM_PRECOGENERICO.HANDLE]|SAM_TGE[SAM_PRECOGENERICO_DOTAC.EVENTO=SAM_TGE.HANDLE]", vColunas, 1, vCampos, vCriterio, "Quantidades", True, "", "")

	Set QueryRetorno = NewQuery

	QueryRetorno.Add("SELECT * FROM SAM_PRECOGENERICO_DOTAC WHERE HANDLE =:HANDLE")

	QueryRetorno.ParamByName("HANDLE").Value = vHandle
	QueryRetorno.Active = True

	CurrentQuery.FieldByName("QTDUSHONORARIO").Value = QueryRetorno.FieldByName("QTDUSHONORARIO").Value
	CurrentQuery.FieldByName("QTDUSCUSTOOPERACIONAL").Value = QueryRetorno.FieldByName("QTDUSCUSTOOPERACIONAL").Value
	CurrentQuery.FieldByName("FATORFILME").Value = QueryRetorno.FieldByName("FATORFILME").Value
	CurrentQuery.FieldByName("PORTEANESTESICO").Value = QueryRetorno.FieldByName("PORTEANESTESICO").Value
	CurrentQuery.FieldByName("PORTESALA").Value = QueryRetorno.FieldByName("PORTESALA").Value

	QueryRetorno.Active = False

	Set Interface = Nothing
	'****************************** SMS 13737 DURVAL 26/12/2002 ***************************************************************************88888
End Sub

Public Sub BOTAOREPLICARPORCBOS_OnClick()

  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  If (CurrentQuery.State <> 1)  Then
	  bsShowMessage("O registro não pode estar em edição. Confirme ou cancela as alterações!", "I")
	  Exit Sub
  End If


  SessionVar("TabelaDeDotacao") = "SAM_PRECOFILIAL_DOTAC"
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
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
Public Sub TABELAFILME_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaFilme(TABELAFILME.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAFILME").Value = vHandle
	End If
End Sub

Public Sub CBOSPESQUISA_OnPopup(ShowPopup As Boolean)
    Dim Interface As Object
    Dim vHandle As Long
    Dim vCampos As String
    Dim vColunas As String
    Dim qCBOS As Object

    ShowPopup = False

    Set Interface = CreateBennerObject("Procura.Procurar")

    vColunas = "TIS_VERSAO.VERSAO|TIS_CBOS.CODIGO|TIS_CBOS.DESCRICAO"
    vCampos = "Versão TISS|Código do CBOS|Descrição do CBOS"
    vHandle = Interface.Exec(CurrentSystem,"TIS_CBOS|TIS_VERSAO[TIS_CBOS.VERSAOTISS = TIS_VERSAO.HANDLE]", vColunas, 2, vCampos, "", "", True, "", CBOSPESQUISA.Text)

    If (vHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("CBOSPESQUISA").Value = vHandle
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

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  	Dim Msg As String
    If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
    Dim Msg As String
    If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  	Dim Msg As String
    If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Dim qCBOS As BPesquisa

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


	Condicao = "AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	If CurrentQuery.FieldByName("CBOS").IsNull Then
    	Condicao = Condicao + " AND (CBOS IS NULL OR CBOS = '')"
	Else
		Condicao = Condicao + " AND CBOS = '" + CurrentQuery.FieldByName("CBOS").AsString + "'"
	End If

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOFILIAL_DOTAC", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "FILIAL", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	'sms 36736 - Lopes
	CurrentQuery.FieldByName("PERCENTUALPGTOUS").AsFloat = CDbl(Format(CurrentQuery.FieldByName("PERCENTUALPGTOUS").AsFloat, "##0.00"))
	CurrentQuery.FieldByName("PERCENTUALPGTOFILME").AsFloat = CDbl(Format(CurrentQuery.FieldByName("PERCENTUALPGTOFILME").AsFloat, "##0.00"))

	If CurrentQuery.FieldByName("CBOSPESQUISA").IsNull Then
		CurrentQuery.FieldByName("CBOS").Value = ""
	End If
End Sub

Public Sub TABLE_AfterScroll()
	LABELPRECO.Text = ""
End Sub

Public Sub TABLE_AfterInsert()
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

	SQL.Active = True

	If SQL.FieldByName("TOTAL").AsInteger = 1 Then
		SQL.Active = False

		SQL.Clear

		SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

		SQL.Active = True

		CurrentQuery.FieldByName("CONVENIO").Value = SQL.FieldByName("HANDLE").Value
	End If

	Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
    Select Case CommandID
    	Case "BOTAOPRECO"
    		BOTAOPRECO_OnClick
    	Case "BOTAOQUANTIDADE"
    		BOTAOQUANTIDADE_OnClick
    	Case "BOTAOREPLICARPORCBOS"
            SessionVar("TabelaDeDotacao") = "SAM_PRECOFILIAL_DOTAC"
            SessionVar("HandleDotac") = CurrentQuery.FieldByName("HANDLE").AsString
	End Select
End Sub

