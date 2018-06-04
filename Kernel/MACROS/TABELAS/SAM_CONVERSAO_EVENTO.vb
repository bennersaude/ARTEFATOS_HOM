'HASH: 45D13E444DBD1662CDA9D589299D7BF8
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraGrauValido"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)

  ShowPopup = False
  If CurrentQuery.FieldByName("EVENTO").IsNull Then
    Exit Sub
  End If

  Dim vHandle As Long
  vHandle = ProcuraGrauValido(CurrentQuery.FieldByName("EVENTO").AsInteger, GRAU.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If

End Sub

