'HASH: 3A00EFABA0559384DF4E3CC2EFD50FA5

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Texto, sVerba As String
  Dim Sql
  CanContinue = True
  CurrentQuery.UpdateRecord

  If (Not CurrentQuery.FieldByName("VERBASUBSTITUICAO").IsNull) And (Not CurrentQuery.FieldByName("VERBA").IsNull) Then
    If CurrentQuery.FieldByName("VERBA").Value = CurrentQuery.FieldByName("VERBASUBSTITUICAO").Value Then
      MsgBox("Verba nÒo pode ser igual Ó verba de substituiþÒo")
      CanContinue = False
    End If
  End If
End Sub

