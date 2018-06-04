'HASH: F6CC9FD5F89322A4FAADD73645AB44ED
'Macro: SAM_PRECOPRESTADOR_AN
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"

Public Sub BOTAOVALORPORTE_OnClick()
	Dim interface As Object
	Dim vValor As Currency
	Dim vValorAuxiliar As Currency
	Dim SQL As Object
	Dim SQL2 As Object
	Dim result As String
	Set SQL = NewQuery


	SQL.Add("SELECT * FROM SAM_CONFIGURABUSCAPRECO")

	SQL.Active = True

	Nivel = -1

	Set SQL2 = NewQuery

	SQL2.Add("SELECT ASSOCIACAO FROM SAM_PRESTADOR WHERE HANDLE = " + CurrentQuery.FieldByName("PRESTADOR").AsString + "")

	SQL2.Active = True

	Set interface = CreateBennerObject("BSPRE001.ROTINAS")

	If SQL2.FieldByName("ASSOCIACAO").AsString = "S" Then
		If SQL.FieldByName("NIVEL1").AsInteger = 5 Then
			Nivel = 1
		ElseIf SQL.FieldByName("NIVEL2").AsInteger = 5 Then
			Nivel = 2
		ElseIf SQL.FieldByName("NIVEL3").AsInteger = 5 Then
			Nivel = 3
		ElseIf SQL.FieldByName("NIVEL4").AsInteger = 5 Then
			Nivel = 4
		ElseIf SQL.FieldByName("NIVEL5").AsInteger = 5 Then
			Nivel = 5
		ElseIf SQL.FieldByName("NIVEL6").AsInteger = 5 Then
			Nivel = 6
		ElseIf SQL.FieldByName("NIVEL7").AsInteger = 5 Then
			Nivel = 7
		ElseIf SQL.FieldByName("NIVEL8").AsInteger = 5 Then
			Nivel = 8
		End If

		If Nivel <> -1 Then
			interface.ValorPorteAnestesico(CurrentSystem, 5, CurrentQuery.FieldByName("HANDLE").AsInteger, vValor, vValorAuxiliar)
		Else
			result = "Na configuração de busca de preço, não foi definido um nível para a Associação!"
			If VisibleMode Then
				LBLVALORPORTE.Text = result
			Else
				bsShowMessage(result, "E")
			End If
			Exit Sub
		End If
	Else
		If SQL.FieldByName("NIVEL1").AsInteger = 4 Then
			Nivel = 1
		ElseIf SQL.FieldByName("NIVEL2").AsInteger = 4 Then
			Nivel = 2
		ElseIf SQL.FieldByName("NIVEL3").AsInteger = 4 Then
			Nivel = 3
		ElseIf SQL.FieldByName("NIVEL4").AsInteger = 4 Then
			Nivel = 4
		ElseIf SQL.FieldByName("NIVEL5").AsInteger = 4 Then
			Nivel = 5
		ElseIf SQL.FieldByName("NIVEL6").AsInteger = 4 Then
			Nivel = 6
		ElseIf SQL.FieldByName("NIVEL7").AsInteger = 4 Then
			Nivel = 7
		ElseIf SQL.FieldByName("NIVEL8").AsInteger = 4 Then
			Nivel = 8
		End If

		If Nivel <> -1 Then
			interface.ValorPorteAnestesico(CurrentSystem, 4, CurrentQuery.FieldByName("HANDLE").AsInteger, vValor, vValorAuxiliar)
		Else
			result = "Na configuração de busca de preço, não foi definido um nível para o Prestador!"
			If VisibleMode Then
				LBLVALORPORTE.Text = result
			Else
				bsShowMessage(result, "E")
			End If

			Exit Sub
		End If
	End If

	result = "Valor do Porte Anestésico nesta Vigência: R$ " + Format(vValor, "#,##0.00") + Chr(13) + Chr(10)
	result = result + "Valor % Auxiliar nesta vigência: R$ " + Format(vValorAuxiliar, "#,##0.00")

	If VisibleMode Then
		LBLVALORPORTE.Text = result
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

Public Sub TABLE_AfterEdit()
	UpdateLastUpdate("SAM_CONVENIO")


	Dim vCondicao As String

	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere = vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
	End If
End Sub

Public Sub TABLE_AfterPost()
	LBLVALORPORTE.Text = ""
End Sub

Public Sub TABLE_AfterScroll()
	LBLVALORPORTE.Text = ""
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim interface As Object
	Dim Linha As String
	Dim Condicao As String
	'SMS 49152 - Anderson Lonardoni
	'Esta verificação foi tirada do BeforeInsert e colocada no
	'BeforePost para que, no caso de Inserção, já existam valores
	'no CurrentQuery e para funcionar com o Integrator
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
	'SMS 49152 - Fim

	Set interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = "AND PORTEANESTESICO = " + CurrentQuery.FieldByName("PORTEANESTESICO").AsString
	Condicao = Condicao + " AND CLASSEASSOCIADO = '" + CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString + "'"

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = interface.Vigencia(CurrentSystem, "SAM_PRECOPRESTADOR_AN", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set interface = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
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

	UpdateLastUpdate("SAM_CONVENIO")

	Dim vCondicao As String

	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere = vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CamContinue As Boolean)
	Select Case CommandID
		Case "BOTAOVALORPORTE"
			BOTAOVALORPORTE_OnClick
	End Select
End Sub
