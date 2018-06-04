'HASH: A7F14E6B6C820A8146CEA18F42963EBF

Public Sub Main
 	' Acessada da interface Web
	Dim piEspecialidade As Long
	Dim psXmlSubEspec As String
	Dim psMensagem As String

	Dim SamConsultaDLL As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")

	piEspecialidade = CLng( ServiceVar("piEspecialidade") )
	psXmlSubEspec = CStr( ServiceVar("psXmlSubEspec") )
	psMensagem = CStr( ServiceVar("psMensagem") )

	SamConsultaDLL.EspecialidadeScroll(CurrentSystem, piEspecialidade, psXmlSubEspec)

	ServiceVar("psXmlSubEspec") = CStr( psXmlSubEspec )

	Set SamConsultaDLL = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	Set SamConsultaDLL = Nothing

End Sub
