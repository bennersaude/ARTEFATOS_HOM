'HASH: F65D0C49567C352E980288319483D89E

Public Sub Main
 	' Acessada da interface Web
	Dim psTab As String
	Dim psBairro As String
	Dim piPosicao As Long
	Dim piEvento As Long
	Dim piMunicipio As Long
	Dim piEstado As Long
	Dim piGrau As Long
	Dim piTipoPrestador As Long
	Dim piTipoPrestador2 As Long
	Dim piCategoriaPrestador As Long
	Dim piRegimeAtendimento As Long
	Dim piEspecialidade As Long
	Dim psSubEspecialidade As String
	Dim piRegiao As Long
	Dim psXmlResult As String
	Dim pbBuscaPorEvento As Boolean
	Dim piChave As Long
	Dim psMensagem As String

	Dim SamConsultaDLL As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")

	psTab = CStr( ServiceVar("psTab") )
	psBairro = CStr( ServiceVar("psBairro") )
	piPosicao = CLng( ServiceVar("piPosicao") )
	piEvento = CLng( ServiceVar("piEvento") )
	piMunicipio = CLng( ServiceVar("piMunicipio") )
	piEstado = CLng( ServiceVar("piEstado") )
	piGrau = CLng( ServiceVar("piGrau") )
	piTipoPrestador = CLng( ServiceVar("piTipoPrestador") )
	piTipoPrestador2 = CLng( ServiceVar("piTipoPrestador2") )
	piCategoriaPrestador = CLng( ServiceVar("piCategoriaPrestador") )
	piRegimeAtendimento = CLng( ServiceVar("piRegimeAtendimento") )
	piEspecialidade = CLng( ServiceVar("piEspecialidade") )
	psSubEspecialidade = CStr( ServiceVar("psSubEspecialidade") )
	piRegiao = CLng( ServiceVar("piRegiao") )
	psXmlResult = CStr( ServiceVar("psXmlResult") )
	'pbBuscaPorEvento = CBool( ServiceVar("pbBuscaPorEvento") )
	piChave = CLng( ServiceVar("piChave") )
	psMensagem = CStr( ServiceVar("psMensagem") )


	SamConsultaDLL.ConsultaPrestadorEvento(CurrentSystem, _
										   psTab, _
										   psBairro, _
										   psLogradouro, _
										   piEvento, _
										   piMunicipio, _
										   piEstado, _
										   piGrau, _
										   piTipoPrestador, _
										   piTipoPrestador2, _
										   piCategoriaPrestador, _
										   piRegimeAtendimento, _
										   piEspecialidade, _
										   psSubEspecialidade, _
										   -1, _
										   piRegiao, _
										   0, _
										   psXmlResult, _
										   pbBuscaPorEvento, _
										   piChave, "N")

	ServiceVar("psXmlResult") = CStr( psXmlResult )
	ServiceVar("pbBuscaPorEvento") = CBool( pbBuscaPorEvento )
	ServiceVar("piChave") = CLng( piChave )
    ServiceVar("psMensagem") = ""
	Set SamConsultaDLL = Nothing

   Exit Sub

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	Set SamConsultaDLL = Nothing


End Sub
