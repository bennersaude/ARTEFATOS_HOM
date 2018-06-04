'HASH: D96658BCF5A9851CE5A4732BE512F7CF
 

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_ACOMODACAO_EVENTO","ACOMODACAO",RecordHandleOfTable("SAM_ACOMODACAO"),"Gerando eventos para acomodação")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_ACOMODACAO_EVENTO"
End Sub
