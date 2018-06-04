'HASH: 8AC8595195395813667E81EFD8F0FABE

'CLI_INSTRUTOR

Public Sub INSTRUTOR_OnPopup(ShowPopup As Boolean)
  '#Uses "*ProcuraPrestador"
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", INSTRUTOR.Text) ' pelo CPF e todos
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("INSTRUTOR").Value = vHandle
  End If
End Sub

