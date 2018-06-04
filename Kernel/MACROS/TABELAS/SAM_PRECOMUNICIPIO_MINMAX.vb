'HASH: 19A9B070F60788C73215D536F9BA13AD
'Macro: SAM_PRECOMUNICIPIO_MINMAX
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"
'#Uses "*ProcuraGrau"

Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	If ShowPopup = False Then
		Exit Sub
	End If

	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraGrau(GRAU.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandle
	End If
End Sub

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
	If CurrentQuery.FieldByName("VALORMAXIMO").AsFloat <CurrentQuery.FieldByName("VALORMINIMO").AsFloat Then
		bsShowMessage("O valor máximo deve ser maior que o mínimo", "E")
		CanContinue = False
		Exit Sub
	End If

	Dim Interface As Object
	Dim Linha As String
	Dim vCondicao As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	vCondicao = "AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	vCondicao = vCondicao + " AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		vCondicao = vCondicao + " AND CONVENIO IS NULL"
	Else
		vCondicao = vCondicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOMUNICIPIO_MINMAX", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "MUNICIPIO", vCondicao)

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
