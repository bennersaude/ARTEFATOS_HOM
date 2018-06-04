'HASH: 179351BE1F4DEF2199B46010E8439C39


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("TABREGRAREDE").AsInteger = 2 Then
    CurrentQuery.FieldByName("QTDDIA").Clear
  End If
  If CurrentQuery.FieldByName("TABREGRAREDEPROPRIA").AsInteger = 2 Then
    CurrentQuery.FieldByName("QTDDIASREDEPROPRIA").Clear
  End If
End Sub

