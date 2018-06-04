'HASH: 800D1A2A9B894105A6CD13B7739E7636
 

Public Sub BOTAOREAJUSTAR_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("SamReajuste.Modulo")
  interface.Exec(CurrentSystem)
  Set interface =Nothing
End Sub
