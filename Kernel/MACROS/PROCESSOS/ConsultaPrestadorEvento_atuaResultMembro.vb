'HASH: 063C5A14615AE3C06F775AC8EA086F82

Public Sub Main

 	' Acessada da interface Web
    Dim piEstado As Long
	Dim piMunicipio As Long
	Dim piRegiao As Long
	Dim psBairro As String
	Dim piExecutorHandle As Long
	Dim piSexo As Long
	Dim psXmlDadosPrestadorMembros As String
	Dim psXmlEndereco As String
	Dim psMensagem As String
	Dim psCategoria As String

	Dim SamConsultaDLL As Object
	Dim qCategoria As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")
	Set qCategoria = NewQuery

	piEstado = CLng( ServiceVar("piEstado") )
	piMunicipio = CLng( ServiceVar("piMunicipio") )
	piRegiao = CLng( ServiceVar("piRegiao") )
	psBairro = CStr( ServiceVar("psBairro") )
	piExecutorHandle = CLng( ServiceVar("piExecutorHandle") )
	piSexo = CLng( ServiceVar("piSexo") )
	psXmlDadosPrestadorMembros = CStr( ServiceVar("psXmlDadosPrestadorMembros") )
	psXmlEndereco = CStr( ServiceVar("psXmlEndereco") )
	psMensagem = CStr( ServiceVar("psMensagem") )
	psCategoria = CStr( ServiceVar("psCategoria") )

	qCategoria.Active = False
	qCategoria.Add("SELECT DESCRICAO")
	qCategoria.Add("  FROM SAM_CATEGORIA_PRESTADOR CP")
	qCategoria.Add("  JOIN SAM_PRESTADOR PR ON (CP.HANDLE = PR.CATEGORIA)")
	qCategoria.Add(" WHERE PR.HANDLE = :HANDLE")
	qCategoria.ParamByName("HANDLE").AsInteger = piExecutorHandle
	qCategoria.Active = True
	psCategoria = qCategoria.FieldByName("DESCRICAO").AsString

	SamConsultaDLL.MembrosFrmResultScroll(CurrentSystem, _
																		piEstado, _
																		piMunicipio, _
																		piRegiao, _
																		psBairro, _
																		piExecutorHandle, _
																		piSexo, _
																		psXmlDadosPrestadorMembros, _
																		psXmlEndereco)

	ServiceVar("piSexo") = CLng( piSexo )
	ServiceVar("psXmlDadosPrestadorMembros") = CStr( psXmlDadosPrestadorMembros )
	ServiceVar("psXmlEndereco") = CStr( psXmlEndereco )
	ServiceVar("psCategoria") = CStr( psCategoria )

	Set SamConsultaDLL = Nothing
	Set qCategoria = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	Set SamConsultaDLL = Nothing
	Set qCategoria = Nothing



End Sub
