'HASH: 4E64187C261B66D36548E4B33CFE2634
'MACRO SAM_PRECOMUNICIPIO_GRAU
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraGrau"
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEventoGrau"

Option Explicit

Public Sub EVENTO_OnExit()
	Dim vHandle As Long
	Dim SQL As Object
	Set SQL = NewQuery

	If Not CurrentQuery.FieldByName("GRAU").IsNull And Not CurrentQuery.FieldByName("EVENTO").IsNull Then
		SQL.Add("SELECT EVENTO         ")
		SQL.Add("  FROM SAM_TGE_GRAU   ")
		SQL.Add(" WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString)
		SQL.Add("   AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString)

		SQL.Active = True

		If SQL.EOF Then
			CurrentQuery.FieldByName("EVENTO").Value = Null
			bsShowMessage("Este evento não é valido", "I")
			Exit Sub
		End If
	End If

	Set SQL = Nothing
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False

	If CurrentQuery.FieldByName("GRAU").IsNull Then
		bsShowMessage("Escolha o grau primeiro", "I")
		Exit Sub
	End If

	vHandle = ProcuraEventoGrau(True, EVENTO.Text, CurrentQuery.FieldByName("GRAU").AsInteger)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	Dim vHandleGrau As Long
	Dim INTERFACE As Object
	Dim vColunas, vCriterio, vCampos, vTabela As String
	Set INTERFACE = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = "SAM_GRAU.GRAU|SAM_GRAU.Z_DESCRICAO|SAM_TIPOGRAU.DESCRICAO"
	vCriterio = "(SAM_GRAU.PRECOPORGRAU = 'S' OR SAM_GRAU.PRECOPORGRAUDOTACAO = 'S')"
	vCampos = "Código do Grau|Descrição|Tipo do Grau"
	vTabela = "SAM_GRAU|SAM_TIPOGRAU[SAM_TIPOGRAU.HANDLE = SAM_GRAU.TIPOGRAU]"
	vHandleGrau = INTERFACE.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, "Graus de atuação com configuração de 'preço por grau'", True, "")

	If vHandleGrau <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandleGrau
	End If

	Set INTERFACE = Nothing
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
	Dim INTERFACE As Object
	Dim Linha As String
	Dim Condicao As String
	Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")

	If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
		Condicao = " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + " AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString
	Else
		Condicao = " AND EVENTO IS NULL AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString
	End If

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Condicao = "AND MUNICIPIO = " + CurrentQuery.FieldByName("MUNICIPIO").AsString

	'SMS 21799
	If Not CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull Then
		Condicao = Condicao + " AND REGIMEATENDIMENTO = " + CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsString
	Else
		Condicao = Condicao + " AND REGIMEATENDIMENTO IS NULL "
	End If

	Linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRECOMUNICIPIO_GRAU", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "GRAU", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set INTERFACE = Nothing
	Dim vHandle As Long
	Dim SQL As Object
	Set SQL = NewQuery

	If Not CurrentQuery.FieldByName("GRAU").IsNull And Not CurrentQuery.FieldByName("EVENTO").IsNull Then
		SQL.Add("SELECT EVENTO         ")
		SQL.Add("  FROM SAM_TGE_GRAU   ")
		SQL.Add(" WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString)
		SQL.Add("   AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString)

		SQL.Active = True

		If SQL.EOF Then
			bsShowMessage("Este evento não é valido", "E")
			CanContinue = False
			Exit Sub
		End If
	End If

	Set SQL = Nothing
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
