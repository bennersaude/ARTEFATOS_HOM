'HASH: E0869A9FC1297EBBB6DD02D5A97225B5

Public Sub Main

	Dim psCEP As String
	Dim piEstado As Long
	Dim piMunicipio As Long
	Dim psBairro As String
	Dim psXmlMunicipios As String
	Dim psMensagem As String

	Dim SamConsultaDLL As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")

	psCEP = CStr( ServiceVar("psCEP") )
	piEstado = CLng( ServiceVar("piEstado") )
	piMunicipio = CLng( ServiceVar("piMunicipio") )
	psBairro = CStr( ServiceVar("psBairro") )
	psXmlMunicipios = CStr( ServiceVar("psXmlMunicipios") )
	psMensagem = CStr( ServiceVar("psMensagem") )

	SamConsultaDLL.AtualizaLogradouro(CurrentSystem, _
																psCEP, _
																piEstado, _
																piMunicipio, _
																psBairro)

	SamConsultaDLL.AtualizaMunicipio(CurrentSystem, piEstado, psXmlMunicipios)

	ServiceVar("piEstado") = CLng( piEstado )
	ServiceVar("piMunicipio") = CLng( piMunicipio )
	ServiceVar("psBairro") = CStr( psBairro )
	ServiceVar("psXmlMunicipios") = CStr( psXmlMunicipios )

	Set SamConsultaDLL = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	Set SamConsultaDLL = Nothing

End Sub
