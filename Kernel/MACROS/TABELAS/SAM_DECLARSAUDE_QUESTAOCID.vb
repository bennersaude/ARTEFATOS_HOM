'HASH: 732A3C24B5AA379AFE2E19B2E7058A0A

Public Sub CID_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|DESCRICAO"
  vCriterio = " ULTIMONIVEL = 'S'"
  vCampos = "Estrutura|Descricao"
  vHandle = interface.Exec(CurrentSystem, "SAM_CID", vColunas, 2, vCampos, vCriterio, "CID", False, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CID").Value = vHandle
  End If
  Set interface = Nothing
End Sub

