'HASH: 544C6D3035C85850A98991478A14D294

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_TETOREEMBOLSO_EVENTO","TETOREEMBOLSO",RecordHandleOfTable("SAM_TETOREEMBOLSO"),"Duplicando eventos",1)
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_TETOREEMBOLSO_EVENTO"
  'SACI
End Sub
