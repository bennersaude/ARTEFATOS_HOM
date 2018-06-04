'HASH: AC94E96CA3A48D4031DB4E41C64B3F54

Public Sub Main

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
