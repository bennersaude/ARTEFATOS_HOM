'HASH: 71964BD061DC8715F95F2929F128FDB5
 

Public Sub BOTAOGERACONDATENDIM_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO|URGENCIA"

  vCampos ="Descrição|Urgência"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_CONDATENDIMENTO",vColunas,vCampos,"SFN_REGRAPAG_CONDATENDIMENTO","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"CONDICAOATENDIMENTO","Seleciona Classe de Eventos")
  Set interface =Nothing
End Sub
