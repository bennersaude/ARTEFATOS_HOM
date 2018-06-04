'HASH: D8BBED2A654E102BCF86A26E1D6F06C8
' SAM_BENEFICIARIO_ANOTADM
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
  sql.ParamByName("TABELA").Value = "SAM_BENEFICIARIO_ANOTADM"
  sql.ParamByName("CAMPO").Value = "BOTAOALTERARRESPONSAVEL"
  sql.ParamByName("USUARIO").Value = CurrentUser
  sql.Active = True

  If (sql.FieldByName("ALTERAR").AsString = "N") Then
    bsShowMessage("Usuário não tem permissão para esta operação", "I")
    vgControlaEdicao = True
    Exit Sub
  Else
	If WebMode Then
		Set sql = Nothing
		Set sql = NewQuery

		If CurrentQuery.State = 3 Then
			bsShowMessage("O registro não pode estar em edição", "I")
			Exit Sub
		End If

		If Not InTransaction Then
			StartTransaction
		End If
		sql.Add("UPDATE SAM_BENEFICIARIO_ANOTADM SET USUARIO=:USUARIO, DATA=:DATA WHERE HANDLE=" + CurrentQuery.FieldByName("HANDLE").AsString)
		sql.ParamByName("USUARIO").Value = CurrentUser
		sql.ParamByName("DATA").Value = ServerNow
		sql.ExecSQL
		If InTransaction Then
			Commit
		End If

		CurrentQuery.Active = False
		CurrentQuery.Active = True

		Set sql = Nothing
		bsShowMessage("Usuário alterado com sucesso!", "I")
		Exit Sub

	Else
		vgControlaEdicao = False
    	CurrentQuery.Edit
    	' SMS - 78294 - Paulo Drummond - 19/03/2007 - Inicio
    	'USUARIO.ReadOnly = False
    	CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
    	CurrentQuery.Post
    	' SMS - 78294 - Paulo Drummond - 19/03/2007 - Fim
	End If
  End If

End Sub

Public Sub TABLE_AfterDelete()
	Dim qAux As Object
	Dim Liminar As Boolean
	Set qAux = NewQuery
	qAux.Add("SELECT COUNT(1) QTDEANOTACAO         ")
	qAux.Add("  FROM SAM_BENEFICIARIO_ANOTADM   A, ")
	qAux.Add("       SAM_ANOTACAOADMINISTRATIVA B  ")
	qAux.Add(" WHERE A.ANOTACAO   = B.HANDLE       ")
	qAux.Add("   AND BENEFICIARIO = :HBENEF        ")
	qAux.ParamByName("HBENEF").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO")
	qAux.Active = True
	Liminar = Not qAux.EOF
	If Liminar Then
		Liminar = qAux.FieldByName("QTDEANOTACAO").AsInteger>0
	End If
	If Not Liminar Then
	  qAux.Active = False
	  qAux.Clear
	  qAux.Add("UPDATE SAM_BENEFICIARIO SET LIMINAR = 'N' WHERE HANDLE = :HANDLE")
	  qAux.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO")
	  qAux.ExecSQL
	End If
	Set qAux = Nothing
End Sub

Public Sub TABLE_AfterInsert()
	If CurrentQuery.FieldByName("BENEFICIARIO").IsNull And SessionVar("HBENEFICIARIO_LIMINAR") <> "" Then
		CurrentQuery.FieldByName("BENEFICIARIO").Value = CLng(SessionVar("HBENEFICIARIO_LIMINAR"))
	End If
End Sub

Public Sub TABLE_AfterScroll()
  USUARIO.ReadOnly = True
  vgControlaEdicao = True

  If WebMode Then
	ANOTACAO.WebLocalWhere = "A.CLASSELIMINAR = 'N'"
  ElseIf VisibleMode Then
	ANOTACAO.LocalWhere = "SAM_ANOTACAOADMINISTRATIVA.CLASSELIMINAR = 'N'"
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

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

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
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
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOALTERARRESPONSAVEL"
			BOTAOALTERARRESPONSAVEL_OnClick
	End Select
End Sub
