'HASH: 0663BC6B9C1D479A2855E15DEFC79ECB
 
Option Explicit

Public Sub GERARTIPOTRATAMENTO_OnClick()
  Dim CaracteristicasTeto As Object

  Set CaracteristicasTeto = CreateBennerObject("SamDupEventos.Rotinas")
  CaracteristicasTeto.GerarCaracteristicasTeto(CurrentSystem, RecordHandleOfTable("SAM_TETOREEMBOLSO"),"SAM_TIPOTRATAMENTO", "SAM_TETOREEMBOLSO_TPTRAT", "TIPOTRATAMENTO")
  Set CaracteristicasTeto = Nothing

  RefreshNodesWithTable("SAM_TETOREEMBOLSO_TPTRAT")
End Sub
