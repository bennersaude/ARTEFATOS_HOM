'HASH: F7FAC28F9BDF66940E379299BC2624EA

Public Sub Main

	Dim pGuia As String

	pGuia = CStr( ServiceVar("pGuia") )

	Dim dll As Object
	Dim mensagem As String
	Dim retorno As Long

	Set dll=CreateBennerObject("BSWebService.ProcessamentoContas")
	retorno = dll.setGuia(CurrentSystem, pGuia, mensagem)

	Set dll=Nothing

	ServiceVar("psResult") = CStr(mensagem)

End Sub
