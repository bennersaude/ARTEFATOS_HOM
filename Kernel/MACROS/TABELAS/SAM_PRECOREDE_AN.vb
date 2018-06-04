'HASH: 41F070AD105A1FD72079052E8EBA940B
'Macro: SAM_PRECOREDE_AN
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOVALORPORTE_OnClick()
	Dim INTERFACE As Object
	Dim vValor As Currency
	Dim vValorAuxiliar As Currency
	Dim SQL As Object
	Dim SQL2 As Object
	Dim Nivel As Integer
	Set SQL = NewQuery

	SQL.Add("SELECT NEGOCIACAOPRECO FROM SAM_REDERESTRITA WHERE HANDLE = :HANDLE")

	SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("REDERESTRITA").Value
	SQL.Active = True

	If SQL.FieldByName("NEGOCIACAOPRECO").AsString = "N" Then
		LBLVALORPORTE.Text = "Esta rede não foi parametrizada para possuir negociação de preço!"
		Exit Sub
		Set SQL = Nothing
	End If

	SQL.Active = False

	SQL.Clear

	SQL.Add("SELECT * FROM SAM_CONFIGURABUSCAPRECO")

	SQL.Active = True

	Nivel = -1

	Set INTERFACE = CreateBennerObject("BSPRE001.ROTINAS")

	If CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull Then
		If SQL.FieldByName("NIVEL1").AsInteger = 2 Then
			Nivel = 1
		ElseIf SQL.FieldByName("NIVEL2").AsInteger = 2 Then
			Nivel = 2
		ElseIf SQL.FieldByName("NIVEL3").AsInteger = 2 Then
			Nivel = 3
		ElseIf SQL.FieldByName("NIVEL4").AsInteger = 2 Then
			Nivel = 4
		ElseIf SQL.FieldByName("NIVEL5").AsInteger = 2 Then
			Nivel = 5
		ElseIf SQL.FieldByName("NIVEL6").AsInteger = 2 Then
			Nivel = 6
		ElseIf SQL.FieldByName("NIVEL7").AsInteger = 2 Then
			Nivel = 7
		ElseIf SQL.FieldByName("NIVEL8").AsInteger = 2 Then
			Nivel = 8
		End If

		If Nivel <> -1 Then
			INTERFACE.ValorPorteAnestesico(CurrentSystem, 2, CurrentQuery.FieldByName("HANDLE").AsInteger, vValor, vValorAuxiliar)
		Else
			LBLVALORPORTE.Text = "Na configuração de busca de preço, não foi definido um nível para a rede restrita!"
			Exit Sub
		End If
	Else
		If SQL.FieldByName("NIVEL1").AsInteger = 1 Then
			Nivel = 1
		ElseIf SQL.FieldByName("NIVEL2").AsInteger = 1 Then
			Nivel = 2
		ElseIf SQL.FieldByName("NIVEL3").AsInteger = 1 Then
			Nivel = 3
		ElseIf SQL.FieldByName("NIVEL4").AsInteger = 1 Then
			Nivel = 4
		ElseIf SQL.FieldByName("NIVEL5").AsInteger = 1 Then
			Nivel = 5
		ElseIf SQL.FieldByName("NIVEL6").AsInteger = 1 Then
			Nivel = 6
		ElseIf SQL.FieldByName("NIVEL7").AsInteger = 1 Then
			Nivel = 7
		ElseIf SQL.FieldByName("NIVEL8").AsInteger = 1 Then
			Nivel = 8
		End If

		If Nivel <> -1 Then
			INTERFACE.ValorPorteAnestesico(CurrentSystem, 1, CurrentQuery.FieldByName("HANDLE").AsInteger, vValor, vValorAuxiliar)
		End If
	End If

	If Nivel = -1 Then
		LBLVALORPORTE.Text = "Na configuração de busca de preço, não foi definido um nível para a rede restrita do prestador!"
	Else
		LBLVALORPORTE.Text = "Valor do Porte Anestésico nesta Vigência: R$ " + Format(vValor, "#,##0.00") + Chr(13) + Chr(10)
		LBLVALORPORTE.Text = LBLVALORPORTE.Text + "Valor % Auxiliar nesta vigência: R$ " + Format(vValorAuxiliar, "#,##0.00")
	End If

	Set INTERFACE = Nothing
End Sub

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaUS(TABELAUS.Text)

	If vHandle <>0 Then
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

	Condicao = " AND PORTEANESTESICO = " + CurrentQuery.FieldByName("PORTEANESTESICO").AsString
	Condicao = Condicao + " AND REDERESTRITA = " + CurrentQuery.FieldByName("REDERESTRITA").AsString

	If CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull Then
		Condicao = Condicao + " AND REDERESTRITAPRESTADOR IS NULL "
	Else
		Condicao = Condicao + " AND REDERESTRITAPRESTADOR = " + CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").AsString
	End If

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRECOREDE_AN", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "REDERESTRITA", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set INTERFACE = Nothing
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

	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		CanContinue = False
		bsShowMessage("Registro finalizado não pode ser alterado", "E")
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
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
