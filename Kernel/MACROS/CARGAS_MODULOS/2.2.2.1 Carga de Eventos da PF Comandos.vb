'HASH: F7A6DB9D80D3EC996B7ED0B904F2B7BD
 

Public Sub BOTAOGERAEVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_TABPF_EVENTO","TABELAPFEVENTO",RecordHandleOfTable("SAM_TABPF"),"Gerando eventos para participações financeiras")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_TABPF_EVENTO"
End Sub
