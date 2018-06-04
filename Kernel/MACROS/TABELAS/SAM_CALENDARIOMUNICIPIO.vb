﻿'HASH: 6378FBBE1A66FC9915ECF5C464F8AB76
'Macro: SAM_CALENDARIOMUNICIPIO
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
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_CALENDARIOMUNICIPIO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "MUNICIPIO", "")

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
End Sub

Public Sub TABLE_AfterPost()
	TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
	If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		DATAFINAL.ReadOnly = False
	Else
		DATAFINAL.ReadOnly = True
	End If
End Sub
