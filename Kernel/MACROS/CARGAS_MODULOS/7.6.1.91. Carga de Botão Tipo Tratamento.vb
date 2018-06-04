'HASH: 2D618F7BD811BC236AA53CE9DE90AD0B
 

Public Sub BOTAOGERATIPOTRATAME_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="CODIGO|DESCRICAO"

  vCampos ="Codigo|Descrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_TIPOTRATAMENTO",vColunas,vCampos,"SFN_REGRAPAG_TIPOTRATAMENTO","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"TIPOTRATAMENTO","Seleciona Tipo de Tratamento")
  Set interface =Nothing
End Sub
