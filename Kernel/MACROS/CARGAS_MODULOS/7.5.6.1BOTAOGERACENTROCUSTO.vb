'HASH: C9E883D5F8225897E4EB2F43BE87E8FE
 
Public Sub GERACENTROCUSTO_OnClick()

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="ESTRUTURA|DESCRICAO"

  vCampos ="Estrutura|Desrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SFN_CENTROCUSTO",vColunas,vCampos,"SFN_ITEMNOTA_TIPOLANCAMENTOCC","ITEMNOTATIPOLANCAMENTO",RecordHandleOfTable("SFN_ITEMNOTA_TIPOLANCAMENTO"),"CENTROCUSTO","Seleciona Centro de custo")
  Set interface =Nothing


End Sub
