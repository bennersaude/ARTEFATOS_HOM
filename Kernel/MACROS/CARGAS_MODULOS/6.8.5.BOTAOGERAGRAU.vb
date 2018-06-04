'HASH: F31DF54E259DE2B7A6A34F29CAC59537
 

Public Sub GERAGRAU_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO|GRAU"

  vCampos ="Descrição|Grau"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_GRAU",vColunas,vCampos,"sam_TIPOGUIA_MDGUIA_GRAU","MODELOGUIA",RecordHandleOfTable("SAM_TIPOGUIA_MDGUIA"),"GRAU","Seleciona Grau")
  Set interface =Nothing
End Sub
