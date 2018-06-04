'HASH: CCBE6187360DE2F2105C1E28C76BEDB7
'Macro: SAM_PLANO_PFGRAU
'#Uses "*ProcuraGrau"


Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  'If Len(GRAU.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraGrau(GRAU.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  ' End If
End Sub

