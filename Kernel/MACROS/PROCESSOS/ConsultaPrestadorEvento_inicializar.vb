'HASH: 732846DC6E0188D980B8E7363DE35668

Public Sub Main
 	' Acessada da interface Web
 	Dim piPais As Long
	Dim piEstado As Long
	Dim piMunicipio As Long
	Dim piEspecialidade As Long
	Dim piPosicao As Long
	Dim psXmlRegimeAtendimento As String
	Dim psXmlEspecialidade As String
	Dim psXmlEstados As String
	Dim psXmlMunicipios As String
	Dim psXmlBeneficiarios As String
	Dim psMensagem As String

	Dim SamConsultaDLL As Object

	On Error GoTo erro

	piPais = CLng( ServiceVar("piPais") )
	piEstado = CLng( ServiceVar("piEstado") )
	piMunicipio = CLng( ServiceVar("piMunicipio") )
	piEspecialidade = CLng( ServiceVar("piEspecialidade") )
	piPosicao = CLng( ServiceVar("piPosicao") )
	psXmlRegimeAtendimento = CStr( ServiceVar("psXmlRegimeAtendimento") )
	psXmlEspecialidade = CStr( ServiceVar("psXmlEspecialidade") )
	psXmlEstados = CStr( ServiceVar("psXmlEstados") )
	psXmlMunicipios = CStr( ServiceVar("psXmlMunicipios") )
	psXmlBeneficiarios = CStr( ServiceVar("psXmlBeneficiarios") )


	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")
	SamConsultaDLL.InitPrestadorEvento(CurrentSystem,  _
																					 -1,  _
																					 "", _
																					 0, _
																					 piPais, _
																					 piEstado, _
																					 piMunicipio, _
																					 piEspecialidade, _
																					 piPosicao, _
																					 psXmlRegimeAtendimento, _
																					 psXmlEspecialidade, _
																					 psXmlEstados, _
																					 psXmlMunicipios, _
																					 psXmlBeneficiarios)

	ServiceVar("piPais") = CLng( piPais )
	ServiceVar("piEstado") = CLng( piEstado )
	ServiceVar("piMunicipio") = CLng( piMunicipio )
	ServiceVar("piEspecialidade") = CLng( piEspecialidade )
	ServiceVar("piPosicao") = CLng( piPosicao )
	ServiceVar("psXmlRegimeAtendimento") = CStr( psXmlRegimeAtendimento )
	ServiceVar("psXmlEspecialidade") = CStr( psXmlEspecialidade )
	ServiceVar("psXmlEstados") = CStr( psXmlEstados )
	ServiceVar("psXmlMunicipios") = CStr( psXmlMunicipios )
	ServiceVar("psXmlBeneficiarios") = CStr( psXmlBeneficiarios )

	Set SamConsultaDLL = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	Set SamConsultaDLL = Nothing


End Sub
