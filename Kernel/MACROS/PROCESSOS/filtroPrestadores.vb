'HASH: 71407D933C75E6ECCFB8F0FFC61649A5
'#Uses "*addXML"

Sub Main
	Dim processo As Long
	Dim CA009 As Object
	Dim query As BPesquisa
	Dim xml As String
	Dim i As Integer
	Dim psTab As String
	Dim piRedeRestrita As Long

	Set query = NewQuery
	Set CA009 = CreateBennerObject("CA009.ConsultaPrestador")
	
	processo = CLng( ServiceVar( "processo" ) )
	
	Select Case processo
		Case 1
			query.Active = False
		
			query.Clear
		
			query.Add("SELECT PA.*,")
			query.Add("               PP.CONSULTAPRESTADORPELOHANDLE")
			query.Add("    FROM SAM_PARAMETROSATENDIMENTO PA,")
			query.Add("                SAM_PARAMETROSPRESTADOR PP")
		
			query.Active = True
		
			xml = ""
		
			xml = xml + "<parametros>"
		
			While Not query.EOF
				For i = 0 To query.FieldCount - 1
					xml = xml + addXML( query.Fields( i ).Name, query.Fields( i ).Name, query )
				Next i
		
				query.Next
			Wend
		
			xml = xml + "</parametros>"
		
			ServiceResult = CStr( xml )
		Case 2
			query.Active = False
		
			query.Clear
		
			query.Add("SELECT CORCADASTRADO,")
			query.Add("				  CORREFERENCIADO,")
			query.Add("				  CORNAOATIVOS,")
			query.Add("				  CORCOMVINCULO,")
			query.Add("				  CORREDEPROPRIA,")
			query.Add("				  CANCELADO")
			query.Add("	  FROM SAM_PARAMETROSATENDIMENTO")
		
			query.Active = True
		
			xml = ""
		
			xml = xml + "<legenda>"
		
			While Not query.EOF
				For i = 0 To query.FieldCount - 1
					xml = xml + addXML( LCase( query.Fields( i ).Name ), query.Fields( i ).Name, query )
				Next i
		
				query.Next
			Wend
		
			xml = xml + "</legenda>"
		
			ServiceResult = CStr( xml )
		Case 3
			Dim psFormacaoPrestador As String
			Dim pbAtend As Boolean
			Dim pbReferenciados As Boolean
			Dim psCodigoAntigo As String
			Dim psPrestador As String
			Dim psCPFCNPJ As String
			Dim psNome As String
			Dim piPosicaoNome As Long
			Dim pbFantasia As Boolean
			Dim psBairro As String
			Dim piPosicaoBairro As Long
			Dim piEstado As Long
			Dim piMunicipio As Long
			Dim piConselho As Long
			Dim piUfConselho As Long
			Dim psRegiao As String
			Dim psInscricao As String
			Dim pbNCredenc As Boolean
			Dim piEsp As Long
			Dim piEspSub As Long
			Dim piCategoriaPrestador As Long
			Dim piTipoPrestador As Long
			Dim piTipoServico As Long
			Dim pbRecebedor As Boolean
			Dim pbExecutor As Boolean
			Dim pbSolicitante As Boolean
			Dim pbLocalExecucao As Boolean
			Dim psPrestadores As String
			Dim psMensagem As String
			Dim result As Long
		
			psTab = CStr( ServiceVar("psTab") )
			psFormacaoPrestador = CStr( ServiceVar("psFormacaoPrestador") )
			pbAtend = CBool( ServiceVar("pbAtend") )
			pbReferenciados = CBool( ServiceVar("pbReferenciados") )
			psCodigoAntigo = CStr( ServiceVar("psCodigoAntigo") )
			psPrestador = CStr( ServiceVar("psPrestador") )
			psCPFCNPJ = CStr( ServiceVar("psCPFCNPJ") )
			psNome = CStr( ServiceVar("psNome") )
			piPosicaoNome = CLng( ServiceVar("piPosicaoNome") )
			pbFantasia = CBool( ServiceVar("pbFantasia") )
			psBairro = CStr( ServiceVar("psBairro") )
			piPosicaoBairro = CLng( ServiceVar("piPosicaoBairro") )
			piEstado = CLng( ServiceVar("piEstado") )
			piMunicipio = CLng( ServiceVar("piMunicipio") )
			piConselho = CLng( ServiceVar("piConselho") )
			piUfConselho = CLng( ServiceVar("piUfConselho") )
			psRegiao = CStr( ServiceVar("psRegiao") )
			psInscricao = CStr( ServiceVar("psInscricao") )
			pbNCredenc = CBool( ServiceVar("pbNCredenc") )
			piEsp = CLng( ServiceVar("piEsp") )
			piEspSub = CLng( ServiceVar("piEspSub") )
			piCategoriaPrestador = CLng( ServiceVar("piCategoriaPrestador") )
			piTipoPrestador = CLng( ServiceVar("piTipoPrestador") )
			piTipoServico = CLng( ServiceVar("piTipoServico") )
			pbRecebedor = CBool( ServiceVar("pbRecebedor") )
			pbExecutor = CBool( ServiceVar("pbExecutor") )
			pbSolicitante = CBool( ServiceVar("pbSolicitante") )
			pbLocalExecucao = CBool( ServiceVar("pbLocalExecucao") )
			psPrestadores = CStr( ServiceVar("psPrestadores") )
			psMensagem = CStr( ServiceVar("psMensagem") )
			piRedeRestrita = CLng(ServiceVar("piredeRestrita"))
		
			result = CA009.SelecionaPrestador( CurrentSystem, _
				psTab, _
				psFormacaoPrestador, _
				pbAtend, _
				pbReferenciados, _
				psCodigoAntigo, _
				psPrestador, _
				psCPFCNPJ, _
				psNome, _
				piPosicaoNome, _
				pbFantasia, _
				psBairro, _
				piPosicaoBairro, _
				EmptyStr, _
				0, _
				piEstado, _
				piMunicipio, _
				0, _
				piConselho, _
				piUfConselho, _
				psRegiao, _
				psInscricao, _
				pbNCredenc, _
				piEsp, _
				0, _
				piEspSub, _
				piCategoriaPrestador, _
				piTipoPrestador, _
				piTipoServico, _
				pbRecebedor, _
				pbExecutor, _
				pbSolicitante, _
				pbLocalExecucao, _
				piRedeRestrita, _
				psPrestadores, _
				psMensagem, _
				"", _
				False)
		
			ServiceVar("psPrestadores") = CStr( psPrestadores )
			ServiceVar("psMensagem") = CStr( psMensagem )
			ServiceResult = CLng( result )
		Case 4
			Dim piHPrestador As Long
			Dim piHEndereco As Long
			Dim piAux As Long
			Dim psRelacionado As String
			Dim psEspecialidades As String
			Dim psHorarios As String
			Dim psTodasEspec As String
			Dim psGeral As String
		
			piHPrestador = CLng( ServiceVar("piHPrestador") )
			piHEndereco = CLng( ServiceVar("piHEndereco") )
			psTab = CStr( ServiceVar("psTab") )
			piAux = CLng( ServiceVar("piAux") )
			psRelacionado = CStr( ServiceVar("psRelacionado") )
			psEspecialidades = CStr( ServiceVar("psEspecialidades") )
			psTodasEspec = CStr( ServiceVar("psTodasEspec") )
			psGeral = CStr( ServiceVar("psGeral") )
			piRedeRestrita = CLng(ServiceVar("piredeRestrita"))
		
			CA009.Scroll( CurrentSystem, _
				piHPrestador, _
				piHEndereco, _
				psTab, _
				piAux, _
				psRelacionado, _
				piRedeRestrita, _
				psEspecialidades, _
				psTodasEspec, _
				psGeral )

			ServiceVar("psEspecialidades") = CStr( psEspecialidades )
			ServiceVar("psHorarios") = CStr( psHorarios )
			ServiceVar("psTodasEspec") = CStr( psTodasEspec )
			ServiceVar("psGeral") = CStr( psGeral )
	End Select

	Set CA009 = Nothing
	Set query = Nothing
End Sub
