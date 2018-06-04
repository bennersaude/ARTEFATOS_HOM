'HASH: 5B9A00775C7FF2072F00B37526514104
Dim IsAreaLivroObrigatorio As Boolean
'Mauricio Ibelli - 14/08/2001 - sms3858 - acesso
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
  Dim qAreaLivro As Object
  Set qAreaLivro = NewQuery

  qAreaLivro.Active = False
  qAreaLivro.Clear
  qAreaLivro.Add ("SELECT COUNT(*) AS QTDE")
  qAreaLivro.Add ("FROM   SAM_AREALIVRO")
  qAreaLivro.Active = True

  IsAreaLivroObrigatorio = False

  If qAreaLivro.FieldByName ("QTDE").AsInteger = 0 Then
    PUBLICARNOLIVRO.Visible = False
    CurrentQuery.FieldByName("PUBLICARNOLIVRO").AsBoolean = False

  ElseIf qAreaLivro.FieldByName ("QTDE").AsInteger = 1 Then

    qAreaLivro.Active = False
    qAreaLivro.Clear
    qAreaLivro.Add ("SELECT HANDLE")
    qAreaLivro.Add ("FROM   SAM_AREALIVRO")
    qAreaLivro.Active = True

    CurrentQuery.FieldByName("PUBLICARNOLIVRO").AsBoolean = True
    CurrentQuery.FieldByName("AREALIVRO").AsInteger = qAreaLivro.FieldByName ("HANDLE").AsInteger

  Else
    IsAreaLivroObrigatorio = True
  End If

  Set qAreaLivro = Nothing

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery

    qPermissao.Active = False
    qPermissao.Clear
    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("  FROM Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add(" WHERE A.FILIAL = :FILIAL")
    qPermissao.Add("   AND A.USUARIO = :USUARIO")

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
    qPermissao.Clear
    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("  FROM Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add(" WHERE A.FILIAL = :FILIAL")
    qPermissao.Add("   AND A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True

    If qPermissao.FieldByName("ALTERAR").AsString <> "S" Then
      bsShowmessage("Permissão negada! Usuário não pode alterar", "E")
      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If
    Set qPermissao = Nothing
  End If

  Dim qAreaLivro As Object
  Set qAreaLivro = NewQuery

  qAreaLivro.Active = False
  qAreaLivro.Clear
  qAreaLivro.Add ("SELECT COUNT(*) AS QTDE")
  qAreaLivro.Add ("FROM   SAM_AREALIVRO")
  qAreaLivro.Active = True

  IsAreaLivroObrigatorio = False

  If qAreaLivro.FieldByName ("QTDE").AsInteger = 0 Then
    PUBLICARNOLIVRO.Visible = False
    CurrentQuery.FieldByName("PUBLICARNOLIVRO").AsBoolean = False

  ElseIf qAreaLivro.FieldByName ("QTDE").AsInteger = 1 Then

    qAreaLivro.Active = False
    qAreaLivro.Clear
    qAreaLivro.Add ("SELECT HANDLE")
    qAreaLivro.Add ("FROM   SAM_AREALIVRO")
    qAreaLivro.Active = True
    CurrentQuery.FieldByName("PUBLICARNOLIVRO").AsBoolean = True
    CurrentQuery.FieldByName("AREALIVRO").AsInteger = qAreaLivro.FieldByName ("HANDLE").AsInteger

  Else
    IsAreaLivroObrigatorio = True
  End If

  Set qAreaLivro = Nothing

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery

    qPermissao.Active = False
    qPermissao.Clear
    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("  FROM Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add(" WHERE A.FILIAL = :FILIAL")
    qPermissao.Add("   AND A.USUARIO = :USUARIO")

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

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If IsAreaLivroObrigatorio And CurrentQuery.FieldByName ("TABIMPORTAR").AsInteger = 1 Then
    If CurrentQuery.FieldByName ("AREALIVRO").IsNull Then
      bsShowMessage("Campo Área do livro é obrigatório.", "E")
      AREALIVRO.SetFocus
      CanContinue = False
    End If
  End If
End Sub
