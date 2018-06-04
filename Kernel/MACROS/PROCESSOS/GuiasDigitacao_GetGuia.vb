'HASH: 302B7C61F0D4F4995CFFA9D6FC2E4204
'#Uses "*getCliente"

Public Sub Main

	Dim pHandle As Long

	pHandle = CLng( ServiceVar("pHandle") )

	Dim dll As Object
	Dim mensagem As String
	Dim retorno As Long
	Dim xml As String

	Set dll=CreateBennerObject("BSWebService.ProcessamentoContas")
	retorno = dll.getGuia(CurrentSystem, pHandle, xml, mensagem)
	Set dll=Nothing

	ServiceVar("piCliente") = CLng(getCliente)

	ServiceVar("pXml") = CStr( xml )

End Sub
