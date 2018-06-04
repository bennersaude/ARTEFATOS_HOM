'HASH: B346E7A9127D11BBD268E8619B5E8437
 
Public Sub BOTAOGERACLASSEEVENT_OnClick()

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO"

  vCampos ="Descrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_CLASSEEVENTO",vColunas,vCampos,"SFN_REGRAPAG_CLASSEEVENTO","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"CLASSEEVENTO ","Seleciona Classe de Eventos")
  Set interface =Nothing


End Sub
