'HASH: BC782B3140CB21339B95BB8BD0DC1AF0

Public Sub Main

	Dim psDigitado As String
	Dim piPeg As Long

	psDigitado = CStr( ServiceVar("psDigitado") )
	piPeg = CLng( ServiceVar("piPeg") )

	Dim dll As Object
	Dim mensagem As String
	Dim retorno As Long
	Dim xml As String

	SessionVar("hPeg") = CStr(piPeg)

	Set dll=CreateBennerObject("BSWebService.ProcessamentoContas")
	retorno = dll.getDadosGuiaPrincipal(CurrentSystem, psDigitado, xml, mensagem)
	Set dll=Nothing

	ServiceVar("psResult") = CStr(xml)

End Sub
