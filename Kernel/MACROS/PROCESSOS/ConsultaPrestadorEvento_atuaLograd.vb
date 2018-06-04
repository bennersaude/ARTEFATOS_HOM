'HASH: 1415EB3EE7279A898631F21407FE5270

Public Sub Main
 	' Acessada da interface Web
	Dim psCEP As String
	Dim piPais As Long
	Dim piEstado As Long
	Dim piMunicipio As Long
	Dim psBairro As String
	Dim psXmlMunicipios As String
	Dim psMensagem As String

	Dim SamConsultaDLL As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")

	psCEP = CStr( ServiceVar("psCEP") )
	piPais = CLng( ServiceVar("piPais") )
	piEstado = CLng( ServiceVar("piEstado") )
	piMunicipio = CLng( ServiceVar("piMunicipio") )
	psBairro = CStr( ServiceVar("psBairro") )
	psXmlMunicipios = CStr( ServiceVar("psXmlMunicipios") )
	psMensagem = CStr( ServiceVar("psMensagem") )

	SamConsultaDLL.AtualizaLogradouro(CurrentSystem, _
																psCEP, _
																piPais, _
																piEstado, _
																piMunicipio, _
																psBairro)

	SamConsultaDLL.AtualizaMunicipio(CurrentSystem, piEstado, psXmlMunicipios)

    ServiceVar("piPais") = CLng( piPais )
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
