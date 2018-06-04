'HASH: 89420371ED8BC49C8C5BFB9C4FC7C4EB
 

 '#Uses "*ProcuraBeneficiarioAtivo"
 Option Explicit


Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)

  If VisibleMode Then
    Dim vHandle As Long

    ShowPopup = False
    vHandle = ProcuraBeneficiarioAtivo(False, ServerDate, BENEFICIARIO.Text)

    If vHandle <> 0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
    End If
  End If

End Sub
