'HASH: 2BF7A5DBBFC66210E25DA7A3E2E2007B
'Macro: SAM_FAMILIA_ANOTADM
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
    bsShowMessage("Usuário não tem permissão para esta operação", "I")
    vgControlaEdicao = True
    Exit Sub
  Else
    vgControlaEdicao = False
    CurrentQuery.Edit
    USUARIO.ReadOnly = False
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
      bsShowMessage("Usuário sem permissão para excluir anotação.", "I")
      CanContinue = False
    End If
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  If vgControlaEdicao Then
    sql.Clear
    SQL.Add("SELECT PERMITIROUTROUSUARIOALTERAR FROM SAM_ANOTACAOADMINISTRATIVA WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ANOTACAO").AsInteger
    SQL.Active = True

    If SQL.FieldByName("PERMITIROUTROUSUARIOALTERAR").AsString = "N" Then
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
    If CommandID = "BOTAOALTERARRESPONSAVEL" Then
		BOTAOALTERARRESPONSAVEL_OnClick
	End If
End Sub
