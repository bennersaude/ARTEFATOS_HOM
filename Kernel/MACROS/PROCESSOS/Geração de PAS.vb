'HASH: 6A5FED9ED4D17567A12E5829883DAFC2
Sub Main
  Dim dll As Object
  Set dll=CreateBennerObject("SAMSOLICITAUX.ROTINAS")
  dll.GeraSolicitacao(CurrentSystem)
  Set dll=Nothing
End sub
