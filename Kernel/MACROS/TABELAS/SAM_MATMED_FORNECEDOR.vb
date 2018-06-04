'HASH: E5B6CCC484ED986D3B0D41A03AB6314D

'#Uses "*ProcuraEventoMatMed"

Public Sub MATMED_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEventoMatMed(True, MATMED.Text)

  If vHandle <> 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("MATMED").Value = vHandle
  End If
End Sub
