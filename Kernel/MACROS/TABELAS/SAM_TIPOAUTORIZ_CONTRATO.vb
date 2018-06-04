'HASH: E245DF849C1EA0AEA1DCBB38098DF987

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
  Dim VColunas As String
  Dim VCampos As String
  Dim Interface As Object
  Set Interface = CreateBennerObject("Procura.Procurar")
  ShowPopup = False
  VColunas = "SAM_CONTRATO.CONTRATO|SAM_CONTRATO.CONTRATANTE"
  VCampos = "Número contrato|Contratante"
  CurrentQuery.FieldByName("CONTRATO").Value = Interface.Exec(CurrentSystem, "SAM_CONTRATO", VColunas, 2, VCampos, "", "Selecione um contrato", False, "", "")
  Set Interface = Nothing
End Sub

