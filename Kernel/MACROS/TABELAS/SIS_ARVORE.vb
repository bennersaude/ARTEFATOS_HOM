'HASH: 07BBD12876E64391966C224D0D3BDFBA


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  CurrentQuery.FieldByName("SQL").Value = TiraAcento(CurrentQuery.FieldByName("SQL").AsString, True)
End Sub

