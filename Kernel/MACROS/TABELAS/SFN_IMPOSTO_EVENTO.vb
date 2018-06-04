'HASH: 4CEF247FDF62BE2BCBA027438AADC951
 
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"

Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub
