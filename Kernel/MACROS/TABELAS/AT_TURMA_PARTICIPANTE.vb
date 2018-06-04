'HASH: 010EFD46DC643379DF5C29CA1B851EBB


Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "MATRICULAFUNCIONAL|NOME|BENEFICIARIO"

  vCriterio = ""
  vCampos = "Matrícula Funcional|Nome|Beneficiário"

  vHandle = interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 1, vCampos, vCriterio, "Beneficiário", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
  Set interface = Nothing

End Sub

