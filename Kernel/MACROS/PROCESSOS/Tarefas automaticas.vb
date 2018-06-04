'HASH: 48B367C2E8B6CD447A73D6770684CD48
Sub Main()
   Dim Executar As Object
   Set Executar = CreateBennerObject("bsprocesso.TarefasAutomaticas")
   Executar.Exec(CurrentSystem)
   Set Executar = Nothing
End Sub

