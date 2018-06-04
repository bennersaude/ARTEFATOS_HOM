'HASH: 21F948307ED2F94579EA9539F68F46D6
 

Public Sub BOTAOGERAEVENTO_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_MODULO_EVENTO","MODULO",RecordHandleOfTable("SAM_MODULO"),"Gerando eventos para modulos")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_MODULO_EVENTO"
End Sub
