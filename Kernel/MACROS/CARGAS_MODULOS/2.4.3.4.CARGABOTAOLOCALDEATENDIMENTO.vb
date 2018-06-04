'HASH: B3BF043DB026CB59BC313F01ACB7BA93
Option Explicit

Public Sub GERARLOCAIS_OnClick()
  Dim CaracteristicasTeto As Object

  Set CaracteristicasTeto = CreateBennerObject("SamDupEventos.Rotinas")
  CaracteristicasTeto.GerarCaracteristicasTeto(CurrentSystem, RecordHandleOfTable("SAM_TETOREEMBOLSO"),"SAM_LOCALATENDIMENTO", "SAM_TETOREEMBOLSO_LOCALATEND", "LOCALATENDIMENTO")
  Set CaracteristicasTeto = Nothing

  RefreshNodesWithTable("SAM_TETOREEMBOLSO_LOCALATEND")
End Sub
