'HASH: BB98AF6AE107255E53DD16A96E0F2FEA
'MACRO='Macro: TV_FILTRO_COT001
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
