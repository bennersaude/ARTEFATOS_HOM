'HASH: B97ED26866E52AC6F573BFF76B5E06A3

Public Sub Main
 	' Acessada da interface Web

	Dim piMestreHandle As Long
	Dim piChave As Long
	Dim psXmlResultMembros As String
	Dim psMensagem As String
	Dim pbTemResultado As Boolean

	Dim SamConsultaDLL As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")

	piMestreHandle = CLng( ServiceVar("piMestreHandle") )
	piChave = CLng( ServiceVar("piChave") )
	psXmlResultMembros = CStr( ServiceVar("psXmlResultMembros") )
	psMensagem = CStr( ServiceVar("psMensagem") )
'	pbTemResultado = CBool( ServiceVar("pbTemResultado") )

	SamConsultaDLL.InitPrestadorEventoMembrosFrm(CurrentSystem, _
																					 piMestreHandle, _
																					 piChave, _
																					 pbTemResultado, _
																					 psXmlResultMembros)

	ServiceVar("psXmlResultMembros") = CStr( psXmlResultMembros )
	ServiceVar("pbTemResultado")  = CBool( pbTemResultado )

	Set SamConsultaDLL = Nothing
	Set qTP = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	ServiceVar("pbTemResultado") = CBool( pbTemResultado )

	Set SamConsultaDLL = Nothing
End Sub
