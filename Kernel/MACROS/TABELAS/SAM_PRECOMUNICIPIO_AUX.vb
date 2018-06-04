'HASH: 8F27D4B2AB0A3F3CE0F26DC75F034899
'Macro: SAM_PRECOMUNICIPIO_AUX
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "M", RecordHandleOfTable("MUNICIPIOS"), "E") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode excluir", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "M", RecordHandleOfTable("MUNICIPIOS"), "A") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode alterar", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "M", RecordHandleOfTable("MUNICIPIOS"), "I") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode incluir", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = " AND NUMAUXILIAR = " + CurrentQuery.FieldByName("NUMAUXILIAR").Value

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOMUNICIPIO_AUX", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "MUNICIPIO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
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
