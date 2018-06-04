'HASH: 9FDFBAC96ED9206FA1B5E9F5D32755EC
 
Option Explicit

Public Sub GERAROBJETIVOS_OnClick()
  Dim CaracteristicasTeto As Object

  Set CaracteristicasTeto = CreateBennerObject("SamDupEventos.Rotinas")
  CaracteristicasTeto.GerarCaracteristicasTeto(CurrentSystem, RecordHandleOfTable("SAM_TETOREEMBOLSO"),"SAM_OBJTRATAMENTO", "SAM_TETOREEMBOLSO_OBJTRAT", "OBJETIVOTRATAMENTO")
  Set CaracteristicasTeto = Nothing

  RefreshNodesWithTable("SAM_TETOREEMBOLSO_OBJTRAT")
End Sub
