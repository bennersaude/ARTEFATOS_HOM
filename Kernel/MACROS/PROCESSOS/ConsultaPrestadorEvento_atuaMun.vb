'HASH: A3EF1452C23B73C7ADF7399C85CE35A1

Public Sub Main
 	' Acessada da interface Web
	Dim piEstado As Long
	Dim psXmlMunicipios As String
	Dim psMensagem As String

	Dim SamConsultaDLL As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")

	piEstado = CLng( ServiceVar("piEstado") )
	psXmlMunicipios = CStr( ServiceVar("psXmlMunicipios") )
	psMensagem = CStr( ServiceVar("psMensagem") )

	SamConsultaDLL.AtualizaMunicipio(CurrentSystem, piEstado, psXmlMunicipios)

	ServiceVar("psXmlMunicipios") = CStr( psXmlMunicipios )

	Set SamConsultaDLL = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	Set SamConsultaDLL = Nothing

End Sub
