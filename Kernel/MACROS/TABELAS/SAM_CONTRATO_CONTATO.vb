'HASH: 2E09728F74807EBDE8380B493154E70E
'#Uses "*bsShowMessage"

'Macro: SAM_CONTRATO_CONTATO

'#Uses "*VerificaEmail"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("EMAIL").IsNull Then
    If Not VerificaEmail(CurrentQuery.FieldByName("EMAIL").AsString) Then
      bsShowMessage("E-mail inválido", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
End Sub

