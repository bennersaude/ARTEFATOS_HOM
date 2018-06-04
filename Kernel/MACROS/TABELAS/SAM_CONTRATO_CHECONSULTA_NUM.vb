'HASH: AD4DFAF6FCEE79AF2ABFD56C4419D9FA


Public Sub TABLE_AfterEdit()
  CurrentQuery.Edit
  If (CurrentQuery.FieldByName("CANCELADODATA").IsNull) And (CurrentQuery.FieldByName("GUIAEVENTO").IsNull) Then
    CurrentQuery.FieldByName("CANCELADODATA").AsDateTime = ServerNow
    CurrentQuery.FieldByName("CANCELADOUSUARIO").AsInteger = CurrentUser
  End If

End Sub

Public Sub TABLE_AfterPost()
  CANCELADOMOTIVO.ReadOnly = True
End Sub

Public Sub TABLE_AfterScroll()
  If (CurrentQuery.FieldByName("CANCELADODATA").IsNull) And (CurrentQuery.FieldByName("GUIAEVENTO").IsNull) Then
    CANCELADOMOTIVO.ReadOnly = False
  Else
    CANCELADOMOTIVO.ReadOnly = True
  End If

End Sub

