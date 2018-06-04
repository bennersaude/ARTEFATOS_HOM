'HASH: C3FB355ACA984EED6E8B325ADC56F66B
Option Explicit

Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_LIMINARBENEF_EVENTOS","LIMINAR",RecordHandleOfTable("SAM_LIMINARBENEF"),"Duplicando eventos para Liminar")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_LIMINARBENEF_EVENTOS"
End Sub
