'HASH: 19D14AB07303A50390A7ADEDE96A82CF

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_ESPECIALIDADESOLICITACAO","ESPECIALIDADE",RecordHandleOfTable("SAM_ESPECIALIDADE"),"Geração de eventos para especialidade")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_ESPECIALIDADESOLICITACAO"
End Sub
