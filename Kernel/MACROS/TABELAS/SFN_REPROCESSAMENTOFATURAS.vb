'HASH: 902E2B024B0F0B8C06BB2BABF4339014
'Macro: SFN_REPROCESSAMENTOFATURAS


Public Sub BOTAOEXECUTAR_OnClick()
  Dim Intergace As Object
  Set Interface = CreateBennerObject("FINANCEIRO.ReprocessaFatura")
  Interface.Exec(CurrentSystem)
  Set interface = Nothing
End Sub

