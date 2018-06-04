'HASH: 6CDC047407C6F897DEAA0F9EE654B371
'#Uses "*getCliente"

Public Sub Main

	Dim pHandle As Long
	pHandle = CLng( ServiceVar("pHandle") )

	ServiceVar("piCliente") = CLng(getCliente)

End Sub
