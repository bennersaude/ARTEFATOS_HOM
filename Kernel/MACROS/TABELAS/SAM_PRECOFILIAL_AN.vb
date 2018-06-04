'HASH: D8EBB253F7E685685EB2E6B743F53089
'Macro: SAM_PRECOFILIAL_AN
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"

Public Sub BOTAOVALORPORTE_OnClick()
	Dim INTERFACE As Object
	Dim vValor As Currency
	Dim vValorAuxiliar As Currency
	Dim q As Object
	Dim result As String
	Set q = NewQuery

	q.Add("SELECT NIVEL8 FROM SAM_CONFIGURABUSCAPRECO")

	q.Active = True

	If q.FieldByName("NIVEL8").IsNull Then
		result = "Na configuração de busca de preço, não foi definido um nível para a Filial!"
		If VisibleMode Then
			LBLVALORPORTE.Text = result
		Else
			bsShowMessage(result, "E")
		End If
		Set q = Nothing
		Exit Sub
	End If

	Set q = Nothing
	Set INTERFACE = CreateBennerObject("BSPRE001.ROTINAS")

	INTERFACE.ValorPorteAnestesico(CurrentSystem, 8, CurrentQuery.FieldByName("HANDLE").AsInteger, vValor, vValorAuxiliar)

	result = "Valor do Porte Anestésico nesta Vigência: R$ " + Format(vValor, "#,##0.00") + Chr(13) + Chr(10)
	result = result + "Valor % Auxiliar nesta vigência: R$ " + Format(vValorAuxiliar, "#,##0.00")

	If VisibleMode Then
		LBLVALORPORTE.Text = result
	Else
		bsShowMessage(result, "I")
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

Public Sub TABLE_AfterPost()
	LBLVALORPORTE.Text = ""
End Sub

Public Sub TABLE_AfterScroll()
	LBLVALORPORTE.Text = ""
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim INTERFACE As Object
	Dim Linha As String
	Dim Condicao As String
	Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = " AND PORTEANESTESICO = " + CurrentQuery.FieldByName("PORTEANESTESICO").Value

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRECOFILIAL_AN", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "FILIAL", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage("Linha", "E")
	End If

	Set INTERFACE = Nothing
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
		Case "BOTAOVALORPORTE"
			BOTAOVALORPORTE_OnClick
	End Select
End Sub
