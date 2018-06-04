'HASH: 778745238FDA01D16BC07BF35A0D3B40


Public Sub REDECONTIDA_OnClick()

  Dim Interface As Object

  Set Interface =CreateBennerObject("BSPRE001.Rotinas")
  Interface.CadRedesContidas(CurrentSystem)
  Set Interface =Nothing

End Sub
