'HASH: 04A31DF1B241E165E36FCB2443D00B40
'#Uses "*ProcuraPrestador"

Option Explicit

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim vTipoBusca As String
  Dim vHandle As Long

  ShowPopup = False

  If (IsNumeric(PRESTADOR.Text)) Then
      vTipoBusca = "C"
  Else
      vTipoBusca = "N"
  End If

  vHandle = ProcuraPrestador(vTipoBusca, "R", PRESTADOR.Text)

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
End Sub
