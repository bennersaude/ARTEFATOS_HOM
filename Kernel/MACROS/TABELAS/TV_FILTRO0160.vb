'HASH: 04BF21DC97044ACDA221CD33AAE8D836
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
    CanContinue = False
    bsShowMessage("Data final não pode ser inferior à data inicial!", "E")
  End If
End Sub
