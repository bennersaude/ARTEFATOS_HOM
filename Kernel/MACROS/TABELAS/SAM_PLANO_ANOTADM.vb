'HASH: ACD2764D461D8E40C0F146CD52505561
'Macro: SAM_PLANO_ANOTADM
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
  sql.Add(" AND T.HANDLE = :TABELA")
  sql.Add("AND C.NOME = :CAMPO")
  sql.Add("AND U.HANDLE = :USUARIO")
  sql.ParamByName("TABELA").Value = CurrentTable
  sql.ParamByName("CAMPO").Value = "BOTAOALTERARRESPONSAVEL"
  sql.ParamByName("USUARIO").Value = CurrentUser
  sql.Active = True

  If (sql.FieldByName("ALTERAR").AsString = "N") Then
    bsShowMessage("Usuário não tem permissão para esta operação", "E")
    vgControlaEdicao = True
    Exit Sub
  Else
  	If VisibleMode Then
    	vgControlaEdicao = False
    	CurrentQuery.Edit
    	USUARIO.ReadOnly = False
    Else
		Dim sqlUsuario As Object
		Set sqlUsuario = NewQuery

		If Not InTransaction Then StartTransaction

		sqlUsuario.Add("UPDATE SAM_PLANO_ANOTADM SET USUARIO=:USUARIO WHERE HANDLE=" + CurrentQuery.FieldByName("HANDLE").AsString)
		sqlUsuario.ParamByName("USUARIO").Value = CurrentUser
		sqlUsuario.ExecSQL

		If InTransaction Then Commit

		Set sqlUsuario = Nothing
    End If
  End If

End Sub

Public Sub TABLE_AfterScroll()
  USUARIO.ReadOnly = True
  vgControlaEdicao = True
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
