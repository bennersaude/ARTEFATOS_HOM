'HASH: 4B9B395BA0ECDC6AC3CFF5EA051ABCA8

Public Sub GERAREVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_GRPREFERENCIAMENTO_EVENTO","GRUPOREFERENCIAMENTO",RecordHandleOfTable("SAM_GRUPOREFERENCIAMENTO"),"Gerando eventos para referenciamento")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_GRPREFERENCIAMENTO_EVENTO"
End Sub
