'HASH: A486B3B5E2FCBC23142CCE7D0D4BB385
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("MATMEDPRECOTAB1").IsNull And CurrentQuery.FieldByName("MATMEDPRECOTAB2").IsNull And CurrentQuery.FieldByName("MATMEDPRECOTAB3").IsNull Then
    bsShowMessage("Deve estar selecionada ao menos uma tabela", "E")
    CanContinue = False
  End If
End Sub

