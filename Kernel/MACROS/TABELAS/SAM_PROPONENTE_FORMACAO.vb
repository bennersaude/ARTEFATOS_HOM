'HASH: F4D6F4498C820DFD1E9FB23A3E903FE2
'Macro: SAM_PROPONENTE_FORMACAO
'Mauricio Ibelli - 14/08/2001 - sms3858 - acesso
 '#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > ServerDate Then

    bsShowMessage("Data inicial não pode ser maior que a data atual!", "E")

    CanContinue = False
  End If

  If Not CurrentQuery.FieldByName("DATACONCLUSAO").IsNull Then
    If CurrentQuery.FieldByName("DATACONCLUSAO").AsDateTime > ServerDate Then

   	  bsShowMessage("Data da conclusão não pode ser maior que a data Atual!", "E")

      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("DATACONCLUSAO").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then

	  bsShowMessage("Data da conclusão não pode ser menor que a data Inicial!", "E")

      CanContinue = False
      Exit Sub
    End If

  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery
    qPermissao.Active = False

    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("FROM   Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add("WHERE  A.FILIAL = :FILIAL")
    qPermissao.Add("AND    A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True
    If qPermissao.FieldByName("EXCLUIR").AsString <> "S" Then

      bsShowMessage("Permissão negada! Usuário não pode excluir", "E")

     CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If
    Set qPermissao = Nothing
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery
    qPermissao.Active = False

    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("FROM   Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add("WHERE  A.FILIAL = :FILIAL")
    qPermissao.Add("AND    A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True
    If qPermissao.FieldByName("ALTERAR").AsString <> "S" Then

  	bsShowMessage("Permissão negada! Usuário não pode alterar", "E")

      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If
    Set qPermissao = Nothing
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery
    qPermissao.Active = False

    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("FROM   Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add("WHERE  A.FILIAL = :FILIAL")
    qPermissao.Add("AND    A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True
    If qPermissao.FieldByName("INCLUIR").AsString <> "S" Then

  	bsShowMessage("Permissão negada! Usuário não pode incluir", "E")

      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If
    Set qPermissao = Nothing
  End If
End Sub
