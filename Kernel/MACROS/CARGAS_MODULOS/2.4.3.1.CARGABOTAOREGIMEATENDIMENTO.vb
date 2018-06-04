'HASH: 3BA993F78065115D75F3FAEEA1C41321
Option Explicit

Public Sub GERARREGIMEATEND_OnClick()
  Dim CaracteristicasTeto As Object

  Set CaracteristicasTeto = CreateBennerObject("SamDupEventos.Rotinas")
  CaracteristicasTeto.GerarCaracteristicasTeto(CurrentSystem, RecordHandleOfTable("SAM_TETOREEMBOLSO"),"SAM_REGIMEATENDIMENTO", "SAM_TETOREEMBOLSO_REGIMEATEND", "REGIMEATENDIMENTO")
  Set CaracteristicasTeto = Nothing

  RefreshNodesWithTable("SAM_TETOREEMBOLSO_REGIMEATEND")
End Sub
