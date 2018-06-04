﻿'HASH: 03D0976A54FE06A940ABA364FF64010B
'SAM_LIVROCONFIG_FILTROREDE
'#Uses "*bsShowMessage"

Public Sub REDERESTRITA_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim Handlexx As Long
	Dim vCondicao As String

	ShowPopup = False

	Set Interface = CreateBennerObject("Procura.Procurar")

	Handlexx = -1
	vCondicao = ""
	Handlexx = Interface.Exec(CurrentSystem, "SAM_REDERESTRITA", "DESCRICAO", 1, "Descrição", vCondicao, "Tabela redes restritas", True, "")

	If Handlexx > 0 Then
		CurrentQuery.FieldByName("REDERESTRITA").Value = Handlexx
	End If

	Set Interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_LIVROCONFIG_FILTROREDE WHERE REDERESTRITA = :REDERESTRITA AND LIVROCONFIGURACAO = :LIVROCONFIGURACAO AND HANDLE <> :HANDLE")

	SQL.ParamByName("REDERESTRITA").Value = CurrentQuery.FieldByName("REDERESTRITA").Value
	SQL.ParamByName("LIVROCONFIGURACAO").Value = CurrentQuery.FieldByName("LIVROCONFIGURACAO").Value
	SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
	SQL.Active = True

	If Not SQL.EOF Then
		bsShowMessage("Registro duplicado para esta configuração !", "E")
		CanContinue = False
		Exit Sub
	End If

	Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
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

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
