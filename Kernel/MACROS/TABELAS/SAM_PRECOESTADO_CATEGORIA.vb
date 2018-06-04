'HASH: 265FBE7DB29DF9A9E9FB7ED08F6BDEC1
'Macro: SAM_PRECOESTADO_CATEGORIA
'#Uses "*bsShowMessage"
'#Uses "*ProcuraTabelaGenerica"

Public Sub TABELAPRECO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaGenerica(TABELAPRECO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAPRECO").Value = vHandle
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If checkPermissao(CurrentSystem, CurrentUser, "E", CurrentQuery.FieldByName("ESTADO").AsInteger, "E") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode excluir", "E")
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

	Condicao = "AND CATEGORIAPRECO = " + CurrentQuery.FieldByName("CATEGORIAPRECO").AsString
	Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOESTADO_CATEGORIA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESTADO", Condicao)

	If Linha <> "" Then
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing

End Sub
