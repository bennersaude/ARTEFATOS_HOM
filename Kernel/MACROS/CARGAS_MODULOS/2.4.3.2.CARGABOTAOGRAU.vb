'HASH: 64506500D65F332E6532ACE5C6DFEF0D
Option Explicit

Public Sub GERARGRAU_OnClick()
  Dim CaracteristicasTeto As Object

  Set CaracteristicasTeto = CreateBennerObject("SamDupEventos.Rotinas")
  CaracteristicasTeto.GerarCaracteristicasTeto(CurrentSystem, RecordHandleOfTable("SAM_TETOREEMBOLSO"),"SAM_GRAU", "SAM_TETOREEMBOLSO_GRAU", "GRAU")
  Set CaracteristicasTeto = Nothing

  RefreshNodesWithTable "SAM_TETOREEMBOLSO_GRAU"
End Sub
