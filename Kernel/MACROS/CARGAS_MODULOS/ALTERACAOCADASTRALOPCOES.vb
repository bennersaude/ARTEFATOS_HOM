'HASH: 84E841A2B82F74946F0EE5129B8A3772
 

Public Sub BOTAOMOVIMENTACAO_OnClick()

  Dim interface As Object

  Set interface = CreateBennerObject("BENNER.SAUDE.DESKTOP.BENEFICIARIOS.MONITORANALISEMOVIMENTACOES.Rotinas")
  interface.ExecutaMonitor(CurrentSystem)
  Set interface = Nothing
  RefreshNodesWithTable("WEB_SAM_BENEFICIARIO")

End Sub

