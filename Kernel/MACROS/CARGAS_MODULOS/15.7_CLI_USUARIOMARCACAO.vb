'HASH: 963549E6720FC5122943F9B7499BB026
 

Public Sub GERAR_OnClick()
  Dim AGENDA As Object
  Set AGENDA =CreateBennerObject("BSCli005.Rotinas")
  AGENDA.Gerar(CurrentSystem)
  Set AGENDA =Nothing
End Sub
