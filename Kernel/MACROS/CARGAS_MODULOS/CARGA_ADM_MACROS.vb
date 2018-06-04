'HASH: A84CD777823FA9681C4019A94BD3CF0C

Public Sub AGENDAMENTOS_OnClick() 
Dim obj As Object 
  Set obj = CreateBennerObject("CS.AgendamentoProcesso") 
  obj.start(CurrentSystem) 
  Set obj = Nothing 
End Sub 
