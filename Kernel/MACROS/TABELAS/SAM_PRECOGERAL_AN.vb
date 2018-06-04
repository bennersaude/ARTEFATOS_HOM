'HASH: AA1CDB57D60DAB37787411B5A6C0B3B8
'Macro: SAM_PRECOGERAL_AN
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOVALORPORTE_OnClick()
	Dim interface As Object
	Dim vValor As Currency
	Dim vValorAuxiliar As Currency
	Dim result As String
	Set interface = CreateBennerObject("BSPRE001.ROTINAS")

	interface.ValorPorteAnestesico(CurrentSystem, 9, CurrentQuery.FieldByName("HANDLE").AsInteger, vValor, vValorAuxiliar)

	result = "Valor do Porte Anestésico nesta Vigência: R$ " + Format(vValor, "#,##0.00") + Chr(13) + Chr(10)
	result = result + "Valor % Auxiliar nesta vigência: R$ " + Format(vValorAuxiliar, "#,##0.00")

	If VisibleMode Then
		LVBVALORPORTE.Text = result
	Else
		bsShowMessage(result, "I")
	End If 
	Set interface = Nothing
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
	LVBVALORPORTE.Text = ""

	If CurrentQuery.FieldByName("PERCENTUALPAGTOUS").IsNull Then
		BOTAOVALORPORTE.Visible = False
	Else
		BOTAOVALORPORTE.Visible = True
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T4377" Then
			CONVENIO.ReadOnly = True
		End If
	End If


	LVBVALORPORTE.Text = ""

	If CurrentQuery.FieldByName("PERCENTUALPAGTOUS").IsNull Then
		BOTAOVALORPORTE.Visible = False
	Else
		BOTAOVALORPORTE.Visible = True
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim interface As Object
	Dim Linha As String
	Dim Convencio As String
	Dim Condicao As String
	Set interface = CreateBennerObject("SAMGERAL.Vigencia")

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = interface.Vigencia(CurrentSystem, "SAM_PRECOGERAL_AN", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PORTEANESTESICO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set interface = Nothing
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
