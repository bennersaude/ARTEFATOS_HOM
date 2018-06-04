'HASH: 3176002FD3C46FC7FAC5073F67C44C3F
'#Uses "*ProcuraPrestador"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub RECEBEDOR_OnPopup(ShowPopup As Boolean)
	Dim vTipoBusca As String
  	Dim vHandle As Long

    ShowPopup = False

    If (IsNumeric(RECEBEDOR.Text)) Then
      vTipoBusca = "C"
    Else
      vTipoBusca = "N"
    End If

    vHandle = ProcuraPrestador(vTipoBusca, "R", RECEBEDOR.Text)


	If vHandle <> 0 Then
    	CurrentQuery.FieldByName("RECEBEDOR").Value = vHandle
  	End If

End Sub
