'HASH: CA09CC908F520D7A483F24F8AD59F84B
'Macro: SAM_PRECOESTADO_ACOMODACAO
'#Uses "*ProcuraGrau"
'#Uses "*bsShowMessage"

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraGrau(GRAU.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandle
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "E", CurrentQuery.FieldByName("ESTADO").AsInteger, "E") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode excluir.", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "E", CurrentQuery.FieldByName("ESTADO").AsInteger, "A") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode alterar", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "E", RecordHandleOfTable("ESTADOS"), "I") = "N" Then
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

	Condicao = " AND ACOMODACAO = " + CurrentQuery.FieldByName("ACOMODACAO").AsString

	If CurrentQuery.FieldByName("GRAU").IsNull Then
		Condicao = Condicao + " AND GRAU IS NULL"
	Else
		Condicao = Condicao + " AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString
	End If

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOESTADO_ACOMODACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESTADO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage("Linha", "E")
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
