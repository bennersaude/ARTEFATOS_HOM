'HASH: C1BEC1EE88C84A01BEA043F4F8479BCB

Public Sub Main

	Dim Obj As Object
	Set Obj = CreateBennerObject("ESPECIFICO.UESPECIFICO")

	ServiceVar("psResult") = CLng( Obj.MPU_PRO_GetNumeroGuia(CurrentSystem) )

	Set Obj = Nothing

End Sub
