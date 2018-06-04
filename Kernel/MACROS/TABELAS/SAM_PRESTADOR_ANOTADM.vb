'HASH: 7A417FFB6E141EA11E372C5AE8BF121A
'Macro: SAM_PRESTADOR_ANOTADM
'#Uses "*bsShowMessage"
Option Explicit

Dim vgControlaEdicao As Boolean

Public Sub ANOTACAO_OnChange()
 'SMS 96046 Bruno Penteado 16/04/2008
  Dim sql As Object
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT OBSERVACAO FROM SAM_ANOTACAOADMINISTRATIVA WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ANOTACAO").AsInteger
  sql.Active = True

  CurrentQuery.FieldByName("OBSERVACAO").AsString = CurrentQuery.FieldByName("OBSERVACAO").AsString + " " + sql.FieldByName("OBSERVACAO").AsString
End Sub

Public Sub BOTAOALTERARRESPONSAVEL_OnClick()
	Dim sql As Object
	Set sql = NewQuery

	sql.Clear

	sql.Add("SELECT P.HANDLE, T.NOME, C.NOME, P.LER, P.ALTERAR")
	sql.Add("FROM Z_GRUPOTABELACAMPOS P, Z_GRUPOUSUARIOS U, Z_CAMPOS C, Z_TABELAS T")
	sql.Add("WHERE C.HANDLE = P.CAMPO")
	sql.Add("AND T.HANDLE = C.TABELA")
	sql.Add("AND P.GRUPO = U.GRUPO")
	sql.Add(" AND T.NOME = :TABELA")
	sql.Add("AND C.NOME = :CAMPO")
	sql.Add("AND U.HANDLE = :USUARIO")

	sql.ParamByName("TABELA").Value = "SAM_PRESTADOR_ANOTADM"
	sql.ParamByName("CAMPO").Value = "BOTAOALTERARRESPONSAVEL"
	sql.ParamByName("USUARIO").Value = CurrentUser
	sql.Active = True

	If (sql.FieldByName("ALTERAR").AsString = "N") Then
		bsShowMessage( "Usuário não tem permissão para esta operação","I")
		vgControlaEdicao = True
		Set sql = Nothing
		Exit Sub
	Else
		vgControlaEdicao = False
		If VisibleMode Then
	      CurrentQuery.Edit
		  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser

		  CurrentQuery.UpdateRecord

		End If
	End If

	Set sql = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	USUARIO.ReadOnly = True
	vgControlaEdicao = True

	Dim sqlAnotacaoAdm As Object
	Set sqlAnotacaoAdm = NewQuery

	sqlAnotacaoAdm.Clear
	sqlAnotacaoAdm.Add("SELECT PERMITIROUTROUSUARIOALTERAR FROM SAM_ANOTACAOADMINISTRATIVA WHERE HANDLE = :HANDLE")
	sqlAnotacaoAdm.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ANOTACAO").AsInteger
	sqlAnotacaoAdm.Active = True

	If sqlAnotacaoAdm.FieldByName("PERMITIROUTROUSUARIOALTERAR").AsString = "N" And	CurrentUser <> CurrentQuery.FieldByName("USUARIO").AsInteger Then
		BOTAOALTERARRESPONSAVEL.Enabled = False
		ANOTACAO.ReadOnly = True
		DATA.ReadOnly = True
		OBSERVACAO.ReadOnly = True
	Else
		BOTAOALTERARRESPONSAVEL.Enabled = True
		ANOTACAO.ReadOnly = False
		DATA.ReadOnly = False
		OBSERVACAO.ReadOnly = False
	End If

	Set sqlAnotacaoAdm = Nothing

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem,"E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Dim sql As Object
	Set sql = NewQuery

	sql.Clear

	sql.Add("SELECT PERMITIROUTROUSUARIOALTERAR FROM SAM_ANOTACAOADMINISTRATIVA WHERE HANDLE = :HANDLE")

	sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ANOTACAO").AsInteger
	sql.Active = True

	If sql.FieldByName("PERMITIROUTROUSUARIOALTERAR").AsString = "N" Then
		If CurrentUser <> CurrentQuery.FieldByName("USUARIO").AsInteger Then
			bsShowMessage("Usuário sem permissão para excluir anotação.", "E")
			CanContinue = False
		End If
	End If

	Set sql = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String
	If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Dim sql As Object
	Set sql = NewQuery

	If vgControlaEdicao Then
		sql.Clear

		sql.Add("SELECT PERMITIROUTROUSUARIOALTERAR FROM SAM_ANOTACAOADMINISTRATIVA WHERE HANDLE = :HANDLE")

		sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ANOTACAO").AsInteger
		sql.Active = True

		If sql.FieldByName("PERMITIROUTROUSUARIOALTERAR").AsString = "N" Then
			If CurrentUser <> CurrentQuery.FieldByName("USUARIO").AsInteger Then
				bsShowMessage("Usuário sem permissão para alterar anotação.", "E")
				CanContinue = False
			End If
		End If
	End If

	Set sql = Nothing
End Sub

Public Sub TABLE_NewRecord()
	CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOALTERARRESPONSAVEL"
			BOTAOALTERARRESPONSAVEL_OnClick
	End Select
End Sub
