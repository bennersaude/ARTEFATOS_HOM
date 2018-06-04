'HASH: 18054810A420B594953435F0ED32CB1F
'#Uses "*ProcuraBeneficiarioAtivo"

Public Sub Beneficiario_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraBeneficiarioAtivo(False, ServerDate , BENEFICIARIO.Text)

  If vHandle <> 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
End Sub
