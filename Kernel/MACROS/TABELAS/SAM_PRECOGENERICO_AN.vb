'HASH: 4A335D29502C6BA9C5C3375688F058EC
'Macro: SAM_PRECOGENERICO_AN
'#Uses "*bsShowMessage"

Public Sub BOTAOVALORPORTE_OnClick()
	Dim interface As Object
	Dim vValor As Currency
	Dim vValorAuxiliar As Currency
	Set interface = CreateBennerObject("BSPRE001.ROTINAS")

	interface.ValorPorteAnestesico(CurrentSystem, -1, _
								   CurrentQuery.FieldByName("PORTEANESTESICO").AsInteger, _
								   -1, -1, -1, -1, -1, -1, _
								   vValor, vValorAuxiliar, "", CurrentQuery.FieldByName("TABELAPRECO").AsInteger)

	LBLVALORPORTE.Text = "Valor do Porte Anestésico nesta Vigência: R$ " + Format(vValor, "#,##0.00") + Chr(13) + Chr(10)
	LBLVALORPORTE.Text = LBLVALORPORTE.Text + "Valor % Auxiliar nesta vigência: R$ " + Format(vValorAuxiliar, "#,##0.00")

	Set interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim interface As Object
	Dim Linha As String
	Dim Condicao As String
	Set interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = "AND PORTEANESTESICO = " + CurrentQuery.FieldByName("PORTEANESTESICO").AsString
	Linha = interface.Vigencia(CurrentSystem, "SAM_PRECOGENERICO_AN", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "TABELAPRECO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set interface = Nothing
End Sub

Public Sub TABLE_BeforeScroll()
	LBLVALORPORTE.Text = ""
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOVALORPORTE"
			BOTAOVALORPORTE_OnClick
	End Select
End Sub
