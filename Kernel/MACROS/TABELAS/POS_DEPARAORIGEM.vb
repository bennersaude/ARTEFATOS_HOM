'HASH: 7C46587610C5229356746C701F4F9FFB


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If TABORIGEM.PageIndex = 0 Then
    CurrentQuery.FieldByName("ORIGEMGUIA").Clear
    If CurrentQuery.FieldByName("ORIGEMAUTORIZACAO").IsNull Then
      MsgBox "Origem da autorização deve ser informado"
      CanContinue = False
      Exit Sub
    End If
  Else
    If CurrentQuery.FieldByName("ORIGEMGUIA").IsNull Then
      MsgBox "Origem da guia deve ser informado"
      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("ORIGEMAUTORIZACAO").IsNull Then
      MsgBox "Origem da autorização deve ser informado"
      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("SITUACAO").IsNull Then
      MsgBox "Situação que a guia será importada deve ser informado"
      CanContinue = False
      Exit Sub
    End If
  End If
End Sub

