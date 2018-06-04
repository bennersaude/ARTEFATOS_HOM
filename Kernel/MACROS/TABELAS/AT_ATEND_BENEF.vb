'HASH: 7D05960806D1ACAD0E9F7FA50621D51C


Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "NOME|BENEFICIARIO"

  vCriterio = ""
  vCampos = "Nome|Beneficiário"

  vHandle = interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 1, vCampos, vCriterio, "Beneficiário", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
  Set INTERFACE = Nothing

End Sub

