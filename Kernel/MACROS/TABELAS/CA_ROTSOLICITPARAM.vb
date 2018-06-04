'HASH: 5D2D85501CE756ADFC4351FEC9CDF773
'#Uses "*bsShowMessage"


Option Explicit

Public Sub TABLE_AfterPost()
  RefreshNodesWithTable("CA_ROTSOLICITPARAM")

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").Value <> "A" Then
    bsShowMessage("Operação não permitida.", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").Value <> "A" Then
    bsShowMessage("Operação não permitida.", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Dim selecao As Integer

  Set SQL = NewQuery
  SQL.Add("SELECT HANDLE,tabprocesso,situacao FROM CA_ROTSOLICIT WHERE handle = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("CA_ROTSOLICIT")
  SQL.Active = True

  If SQL.FieldByName("situacao").AsString <> "A" Then
    CanContinue = False
    Exit Sub
  End If

  Set SQL = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Selecao geral e criterio geral
  'selecao Individual e criterio geral e unico
  Dim SQL As Object
  Dim selecao As Integer


  If CurrentQuery.FieldByName("TABCRITERIO").AsInteger = 0 Then
    bsShowMessage("Seleção é obrigatório.", "E")
    CanContinue = False
    Exit Sub
  End If

  Set SQL = NewQuery
  SQL.Add("SELECT HANDLE,tabprocesso FROM CA_ROTSOLICIT WHERE handle = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("rotsolicit").AsInteger
  SQL.Active = True

  selecao = SQL.FieldByName("tabprocesso").AsInteger

  Set SQL = Nothing


  If selecao = 1 Then 'Geral
    If CurrentQuery.FieldByName("TABCRITERIO").Value = 2 Then 'Individual
      bsShowMessage("Opção não válida para tipo do processo - Geral", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  Set SQL = NewQuery
  SQL.Add("SELECT HANDLE, tabcriterio FROM CA_ROTSOLICITPARAM WHERE ROTSOLICIT = :ROTSOLICIT order by TABCRITERIO ")
  SQL.ParamByName("ROTSOLICIT").Value = CurrentQuery.FieldByName("rotsolicit").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    Exit Sub
  End If

  If SQL.FieldByName("TABCRITERIO").Value = 1 Then
    bsShowMessage("Critério não válido.", "E")
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABCRITERIO").Value = 1 Then
    bsShowMessage("Critério não válido.", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

