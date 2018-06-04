'HASH: 13EB50E935DB06F0C4C824E734B9C71C

Public Sub GERAR_OnClick()
  Dim CSHelp As Object
  Set CSHelp = CreateBennerObject("CSHelp.Gerador")
  CSHelp.Exec
  Set CSHelp = Nothing
End Sub

Public Sub EDITAR_OnClick()
  Dim CSHelp As Object
  Set CSHelp = CreateBennerObject("CSHelp.Arvored")
  CSHelp.Exec
  Set CSHelp = Nothing
End Sub

