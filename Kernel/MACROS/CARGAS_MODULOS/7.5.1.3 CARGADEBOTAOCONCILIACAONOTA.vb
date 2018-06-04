'HASH: D20762786F26CFAE2D11A58B84B644D1
 

Public Sub BOTAOCONCILIACAONOTA_OnClick()
Dim Interface As Object

Set Interface =CreateBennerObject("SFNNota.Rotinas")
Interface.Conciliar(CurrentSystem,0,"D")
Set Interface =Nothing
RefreshNodesWithTable("SFN_NOTA")
End Sub
