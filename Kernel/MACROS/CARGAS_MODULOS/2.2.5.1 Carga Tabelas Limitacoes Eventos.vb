'HASH: 0E5A2A1BB2EAC12AF9DD37540766F45D
 

Public Sub BOTAOGERAEVENTO_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_LIMITACAO_EVENTO","LIMITACAO",RecordHandleOfTable("SAM_LIMITACAO"),"Gerando eventos para limites")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_LIMITACAO_EVENTO"
End Sub
