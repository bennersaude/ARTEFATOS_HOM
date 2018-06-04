'HASH: 5B98DF8E3B44156C7831C71F87CA8A6D
 

Public Sub GERACENTROCUSTO_OnClick()

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="ESTRUTURA|DESCRICAO"

  vCampos ="Estrutura|Desrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SFN_CENTROCUSTO",vColunas,vCampos,"SFN_ITEMNOTA_CLASSEGERENCIALCC","ITEMNOTACLASSEGERENCIAL",RecordHandleOfTable("SFN_ITEMNOTA_CLASSEGERENCIAL"),"CENTROCUSTO","Seleciona Centro de custo")
  Set interface =Nothing


End Sub
