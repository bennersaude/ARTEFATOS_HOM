'HASH: 8CFCEEEB42091C941F975AE695C8178A

Public Sub Main

 	' Acessada da interface Web
	Dim piMestreHandle As Long
	Dim piMunicipio As Long
	Dim piEstado As Long
	Dim piEspecialidade As Long
	Dim psBairro As String
	Dim pbBuscaPorEvento As Boolean
	Dim piSexo As Long
	Dim psStatusMensagem As String
	Dim psXmlEndereco As String
	Dim psXmlDadosPrestador As String
	Dim psMensagem As String
	Dim psCategoria As String

	Dim SamConsultaDLL As Object
	Dim qCategoria As Object

	On Error GoTo erro

	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")
	Set qCategoria = NewQuery

	piMestreHandle = CLng( ServiceVar("piMestreHandle") )
	piMunicipio = CLng( ServiceVar("piMunicipio") )
	piEstado = CLng( ServiceVar("piEstado") )
	piEspecialidade = CLng( ServiceVar("piEspecialidade") )
	psBairro = CStr( ServiceVar("psBairro") )
	pbBuscaPorEvento = CBool( ServiceVar("pbBuscaPorEvento") )
	piSexo = CLng( ServiceVar("piSexo") )
	psStatusMensagem = CStr( ServiceVar("psStatusMensagem") )
	psXmlEndereco = CStr( ServiceVar("psXmlEndereco") )
	psXmlDadosPrestador = CStr( ServiceVar("psXmlDadosPrestador") )
	psMensagem = CStr( ServiceVar("psMensagem") )
	psCategoria = CStr( ServiceVar("psCategoria") )

	psStatusMensagem = ""

	qCategoria.Active = False
	qCategoria.Add("SELECT DESCRICAO")
	qCategoria.Add("  FROM SAM_CATEGORIA_PRESTADOR CP")
	qCategoria.Add("  JOIN SAM_PRESTADOR PR ON (CP.HANDLE = PR.CATEGORIA)")
	qCategoria.Add(" WHERE PR.HANDLE = :HANDLE")
	qCategoria.ParamByName("HANDLE").AsInteger = piMestreHandle
	qCategoria.Active = True
	psCategoria = qCategoria.FieldByName("DESCRICAO").AsString

	SamConsultaDLL.ResultScroll(CurrentSystem, _
													piMestreHandle, _
													piMunicipio, _
													piEstado, _
													piEspecialidade, _
													psBairro, _
													pbBuscaPorEvento, _
													piSexo, _
													psStatusMensagem, _
													psXmlEndereco, _
													psXmlDadosPrestador)

	ServiceVar("piSexo") = CLng( piSexo )
	ServiceVar("psStatusMensagem") = CStr( psStatusMensagem )
	ServiceVar("psXmlEndereco") = CStr( psXmlEndereco )
	ServiceVar("psXmlDadosPrestador") = CStr( psXmlDadosPrestador )
	ServiceVar("psCategoria") = CStr( psCategoria )

	Set SamConsultaDLL = Nothing
	Set qCategoria = Nothing

	erro:
		psMensagem = Err.Description
		ServiceVar("psMensagem") = CStr( psMensagem )

	Set SamConsultaDLL = Nothing
	Set qCategoria = Nothing


End Sub
