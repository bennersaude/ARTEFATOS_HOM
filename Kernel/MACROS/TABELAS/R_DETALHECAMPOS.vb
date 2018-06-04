'HASH: B74DF350FEA8D9506764CF01922EA523


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("TIPO").AsInteger = 2 Then
    CurrentQuery.FieldByName("ORDENAR").Value = "N"
  End If

End Sub

