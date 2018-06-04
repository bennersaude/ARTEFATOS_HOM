'HASH: B9EEBB2F70C71B2F6314826E3F08004A
 

Public Sub BOTAOGERAEVENTO_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_CONTRATO_EVENTOSEMAUTORIZ","CONTRATO",RecordHandleOfTable("SAM_CONTRATO"),"Gerando eventos sem autorização")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_CONTRATO_EVENTOSEMAUTORIZ"
End Sub
