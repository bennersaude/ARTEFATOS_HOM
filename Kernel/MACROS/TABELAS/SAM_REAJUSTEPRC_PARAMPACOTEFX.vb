'HASH: FBCE4BBBE18AAED2A06A491DB1A05515
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(False, EVENTOINICIAL.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
  '  End If
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(False, EVENTOFINAL.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
  '  End If
End Sub

