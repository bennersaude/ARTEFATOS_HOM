'HASH: 13F70B6F12FA956B2A38477D2311E9DC

Public Sub EDITAR_OnClick()
  Dim CSHelp As Object
  Set CSHelp = CreateBennerObject("CSHelp.Documento")
  CSHelp.Exec
  Set CSHelp = Nothing
End Sub


Public Sub PREVERAJUDA_OnClick()
  Set CSHelp = CreateBennerObject("CSHelp.Prever")
  CSHelp.Exec(CurrentSystem)
  Set CSHelp = Nothing
End Sub

