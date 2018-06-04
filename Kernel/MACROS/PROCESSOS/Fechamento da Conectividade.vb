'HASH: E14E47373022A2DEE910E3EE44DECAA9

Public Sub Main
	Dim Obj As Object
	Set Obj = CreateBennerObject("BSATE009.Rotinas")
	Obj.Fechamento(CurrentSystem,0)
	Set Obj = Nothing
End Sub
