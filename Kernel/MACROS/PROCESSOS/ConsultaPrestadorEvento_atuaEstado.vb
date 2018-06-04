'HASH: 78D12EBAE00D7F4F71DBBEE640EF38AA

Public Sub Main
 	' Acessada da interface Web
	Dim piPais As Long
	Dim psXmlEstados As String
	Dim psMensagem As String

	Dim SamConsultaDLL As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")

	piPais = CLng( ServiceVar("piPais") )
	psXmlEstados = CStr( ServiceVar("psXmlEstados") )
	psMensagem = CStr( ServiceVar("psMensagem") )

	SamConsultaDLL.AtualizaEstado(CurrentSystem, piPais, psXmlEstados)

	ServiceVar("psXmlEstados") = CStr( psXmlEstados )

	Set SamConsultaDLL = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	Set SamConsultaDLL = Nothing

End Sub
