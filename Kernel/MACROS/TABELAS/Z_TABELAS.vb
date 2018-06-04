'HASH: B0A715A2FB84B1FE03BA572CB6E21824

Public Sub DOCUMENTACAO_OnClick()
  Dim CSHelp As Object
  Set CSHelp = CreateBennerObject("CSHelp.Tabela")
  CSHelp.Exec
  Set CSHelp = Nothing
End Sub


Public Sub PREVERAJUDA_OnClick()
  Set CSHelp = CreateBennerObject("CSHelp.Prever")
  CSHelp.Exec(CurrentSystem)
  Set CSHelp = Nothing
End Sub

