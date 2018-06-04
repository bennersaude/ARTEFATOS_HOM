'HASH: A4B9F0C02BE18A42FBFEC05919689C5E

'CLI_RECURSO_RECOMENDACAO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime < ServerNow Then
  '  MsgBox("A data inicial não pode ser anterior a data corrente!")
  '  CanContinue = False
  '  Exit Sub
  'End If

  If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime) And _
      (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    bsShowMessage("A data inicial não pode ser posterior a data final!", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

