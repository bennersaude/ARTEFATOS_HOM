'HASH: 9677A907F4DBFF884533FE0C1EB26CDB
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("IDADEMINIMA").AsInteger > CurrentQuery.FieldByName("IDADEMAXIMA").AsInteger Then
    CanContinue = False
    bsShowMessage("Idade mínima não pode ser superior à idade máxima!", "E")
    Exit Sub
  End If
End Sub
