'HASH: D22F6FE8CA06979CEC6791BA4803E2F7
 

Public Sub BOTAOGERAEVENTOS_OnClick()
 
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SFN_REGRAPAG_EVENTO","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"Gerando eventos para Regra de Pagamento")
  Set Duplica =Nothing
  RefreshNodesWithTable "SFN_REGRAPAG_EVENTO"

End Sub
