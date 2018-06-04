'HASH: 569EF97F4A2B3E62850C354FC2FAB608
 

Public Sub BOTAOGERAEVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_TIPOGUIA_MDGUIA_EVENTOTGE","MODELOGUIA",RecordHandleOfTable("SAM_TIPOGUIA_MDGUIA"),"Gerando eventos para o modelo de guia")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_TIPOGUIA_MDGUIA_EVENTOTGE"
End Sub
