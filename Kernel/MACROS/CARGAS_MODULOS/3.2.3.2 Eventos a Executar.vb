'HASH: 421269F94C68E1A9D1F57E2D45A7A46B
 

Public Sub BOTAOEVENTOEXECUCAO_OnClick()
  Dim Interface As Object
  Set Interface =CreateBennerObject("DuplicaEvento.DupEvento")
  Interface.Especialidade(CurrentSystem,"SAM_ESPECIALIDADEEXECUCAO",RecordHandleOfTable("SAM_ESPECIALIDADE"))
End Sub
