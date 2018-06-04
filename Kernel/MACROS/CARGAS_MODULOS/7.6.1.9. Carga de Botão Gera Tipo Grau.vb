'HASH: D79047DEB45EE0D7A3996669420158BC
 

Public Sub BOTAOGERATIPOGRAU_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO"

  vCampos ="Descrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_TIPOGRAU",vColunas,vCampos,"SFN_REGRAPAG_TIPOGRAU","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"TIPOGRAU","Seleciona Tipo de Grau")
  Set interface =Nothing
End Sub
