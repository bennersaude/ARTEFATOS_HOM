'HASH: A2F2DC78345CDEB141F174C6681FC4EA
'#Uses "*ProcuraEvento"
'#uses "*ProcuraPrestador"
'#uses "*IsInt"
Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  ShowPopup = False
  Dim EventoSelecionado As Long
  EventoSelecionado = ProcuraEvento(True, EVENTO.Text)

  If EventoSelecionado > 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").AsInteger = EventoSelecionado
  End If
End Sub


Public Sub EVENTOGERADO_OnPopup(ShowPopup As Boolean)
  ShowPopup = False
  Dim EventoSelecionado As Long
  EventoSelecionado = ProcuraEvento(True, EVENTOGERADO.Text)

  If EventoSelecionado > 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOGERADO").AsInteger = EventoSelecionado
  End If
End Sub

Public Sub EXECUTOR_OnPopup(ShowPopup As Boolean)

  ShowPopup = False
  Dim vHandle As Long


  If IsInt(EXECUTOR.Text) Then
    vHandle = ProcuraPrestador("C", "T", EXECUTOR.Text)
  Else
    vHandle = ProcuraPrestador("N", "T", EXECUTOR.Text)
  End If

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EXECUTOR").AsInteger = vHandle
  End If



End Sub

