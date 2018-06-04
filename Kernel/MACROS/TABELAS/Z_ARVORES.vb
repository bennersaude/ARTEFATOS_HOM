'HASH: 893051192A59D95F02C2E56210AAE142

Public Sub DOCUMENTACAO_OnClick()
  Dim CSHelp As Object
  Dim aModule As Boolean
  If NodeInternalCode = 1 Then
    aModule = False
  Else
    aModule = True
  End If
  Set CSHelp = CreateBennerObject("CSHelp.Carga")
  CSHelp.Exec(aModule)
  Set CSHelp = Nothing
End Sub


Public Sub PREVERAJUDA_OnClick()
  Set CSHelp = CreateBennerObject("CSHelp.Prever")
  CSHelp.Exec(CurrentSystem)
  Set CSHelp = Nothing
End Sub

