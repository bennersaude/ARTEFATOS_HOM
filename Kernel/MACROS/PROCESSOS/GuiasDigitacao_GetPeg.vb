'HASH: 5A359CB1C24D86CFAAD26BE97D7662EE
'#Uses "*getCliente"

Public Sub Main

	Dim pHandle As Long
	pHandle = CLng( ServiceVar("pHandle") )


	Dim dll As Object
	Dim mensagem As String
	Dim retorno As Long
	Dim xml As String

	Set dll=CreateBennerObject("BSWebService.ProcessamentoContas")
	retorno = dll.getPeg(CurrentSystem, pHandle, xml, mensagem)
	Set dll=Nothing

	ServiceVar("piCliente") = CLng(getCliente)

	ServiceVar("pXml") = CStr( xml )

	ServiceVar("psMensagem") = CStr(mensagem)

End Sub
