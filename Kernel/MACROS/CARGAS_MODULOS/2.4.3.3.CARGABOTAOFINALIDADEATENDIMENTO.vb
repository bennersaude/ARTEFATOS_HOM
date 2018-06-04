'HASH: C59D931F31A501FEB661469628380D29
Option Explicit

Public Sub GERARFINALIDADES_OnClick()

  Dim CaracteristicasTeto As Object

  Set CaracteristicasTeto = CreateBennerObject("SamDupEventos.Rotinas")
  CaracteristicasTeto.GerarCaracteristicasTeto(CurrentSystem, RecordHandleOfTable("SAM_TETOREEMBOLSO"),"SAM_FINALIDADEATENDIMENTO", "SAM_TETOREEMBOLSO_FINATEND", "FINALIDADEATENDIMENTO")
  Set CaracteristicasTeto = Nothing

  RefreshNodesWithTable "SAM_TETOREEMBOLSO_FINATEND"
End Sub
