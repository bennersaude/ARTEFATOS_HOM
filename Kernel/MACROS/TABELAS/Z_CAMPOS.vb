'HASH: C52E9EF1934E2769A42FC15553FE6E59

Public Sub DOCUMENTACAO_OnClick()
  Dim CSHelp As Object
  Set CSHelp = CreateBennerObject("CSHelp.Campo")
  CSHelp.Exec
  Set CSHelp = Nothing
End Sub


Public Sub PREVERAJUDA_OnClick()
  Set CSHelp = CreateBennerObject("CSHelp.Prever")
  CSHelp.Exec(CurrentSystem)
  Set CSHelp = Nothing
End Sub

