'HASH: 2A7EDFEA8019D57953BB4D7F23E3E68E
 
'Macro: SAM_TAXAADMPAGAMENTO_EVENTO
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"
Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

End Sub

Public Sub TABLE_AfterScroll()
  If (WebMode) Then
    EVENTO.WebLocalWhere = " A.ULTIMONIVEL = 'S' "
  End If
End Sub
