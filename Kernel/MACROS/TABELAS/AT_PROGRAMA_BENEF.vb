'HASH: CA85D22B268A880691A193022EE497DF


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
  vCampos = "Matrícula Funcional|Nome|Código"

  vHandle = interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 2, vCampos, vCriterio, "Beneficiários", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
  Set interface = Nothing

End Sub

