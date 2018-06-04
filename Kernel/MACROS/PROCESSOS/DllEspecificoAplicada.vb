'HASH: 0FE2DFA1C4FDC0893190C6D043A85795

Public Sub Main

	Dim DLLEspecifico As Object
	Set DLLEspecifico = CreateBennerObject("ESPECIFICO.UESPECIFICO")
    MsgBox(Str( DLLEspecifico.Cliente(CurrentSystem)))
	Set DLLEspecifico = Nothing

End Sub
