'HASH: C75CFF60130C1F608B173FB57D994A81
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
