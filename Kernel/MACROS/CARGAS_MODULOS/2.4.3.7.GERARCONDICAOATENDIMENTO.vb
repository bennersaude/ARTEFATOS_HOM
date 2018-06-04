'HASH: 650BF3A1F50EE3BBAFED6D95B4927BD6
Option Explicit

Public Sub GERARCONDICAOATEND_OnClick()
  Dim CaracteristicasTeto As Object

  Set CaracteristicasTeto = CreateBennerObject("SamDupEventos.Rotinas")
  CaracteristicasTeto.GerarCaracteristicasTeto(CurrentSystem, RecordHandleOfTable("SAM_TETOREEMBOLSO"),"SAM_CONDATENDIMENTO", "SAM_TETOREEMBOLSO_CONDATEND", "CONDICAOATENDIMENTO")
  Set CaracteristicasTeto = Nothing

  RefreshNodesWithTable("SAM_TETOREEMBOLSO_CONDATEND")
End Sub
