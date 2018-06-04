'HASH: 714D5F19EBEF3C287B9088D91B8FE8BB
'#Uses "*CriaTabelaTemporariaSqlServer"

Sub Main()
	Dim viPeg As Long
	Dim psXmlGlosas As String
	Dim psXmlNegacoes As String
	Dim psTipo As String

	Dim xmlGlosas As String
	Dim xmlNegacoes As String

	Dim vsMsgRetorno As String
	Dim vvPegConf As Object

	Dim vcContainer As CSDContainer
	Set vcContainer = NewContainer

	viPeg = CLng( SessionVar("hpeg") )
	psTipo = CStr( ServiceVar("psTipo") )
	psXmlGlosas = CStr( ServiceVar("psXmlGlosas") )
	psXmlNegacoes = CStr( ServiceVar("psXmlNegacoes") )

    vsMsgRetorno = ""

    psXmlGlosas = Replace( Replace( psXmlGlosas, "&lt", ">" ), "&gt", "<")
	psXmlNegacoes = Replace( Replace( psXmlNegacoes, "&lt", ">" ), "&gt", "<")

	On Error GoTo erro

		Dim SQLAgendado As Object
		Set SQLAgendado = NewQuery
		SQLAgendado.Add("SELECT EXECUTACONFERENCIAAGENDADA FROM SAM_PARAMETROSWEB")
		SQLAgendado.Active = True

		CriaTabelaTemporariaSqlServer

		If SQLAgendado.FieldByName("EXECUTACONFERENCIAAGENDADA").AsString = "S" Then

			Dim samPeg As BPesquisa
			Set samPeg = NewQuery

			samPeg.Clear
			samPeg.Add("SELECT P.SITUACAOPROCESSAMENTO FROM SAM_PEG P WHERE P.HANDLE = :PEG")
			samPeg.ParamByName("PEG").Value = viPeg
			samPeg.Active = True

			If samPeg.FieldByName("SITUACAOPROCESSAMENTO").AsString = "1" Or samPeg.FieldByName("SITUACAOPROCESSAMENTO").AsString = "5"Then

				samPeg.Clear
				samPeg.Add("UPDATE SAM_PEG ")
				samPeg.Add("   SET DATA = :DATA")
				samPeg.Add(" WHERE HANDLE = :HANDLE")
				samPeg.ParamByName("HANDLE").AsInteger = viPeg
				samPeg.ParamByName("DATA").AsDateTime = ServerNow

				StartTransaction
				samPeg.ExecSQL
				Commit


				Dim vsMensagemErro As String
			  	Dim viRet As Long
			  	Dim Obj As Object

				'Dim vcContainer As CSDContainer
			   	'Set vcContainer = NewContainer
			   	vcContainer.AddFields("HANDLEPEG:INTEGER")
			   	vcContainer.AddFields("XMLGLOSAS:STRING")
			   	vcContainer.AddFields("XMLNEGACOES:STRING")
			   	vcContainer.AddFields("TIPO:STRING")

				vcContainer.Insert
			 	vcContainer.Field("HANDLEPEG").AsInteger = viPeg
			 	vcContainer.Field("XMLGLOSAS").AsString = psXmlGlosas
			 	vcContainer.Field("XMLNEGACOES").AsString = psXmlNegacoes
			 	vcContainer.Field("TIPO").AsString = psTipo

			  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
				viRet = Obj.ExecucaoImediata(CurrentSystem, _
			                                	 "SAMPEGCONFERENCIA", _
			                                	 "Rotinas", _
			                                	 "Conferência de PEG", _
			                                	 viPeg, _
			                                	 "SAM_PEG", _
			                                	 "SITUACAOPROCESSAMENTO", _
			                                	 "", _
			                                	 "", _
			                                	 "P", _
			                                	 True, _
			                                	 vsMensagemErro, _
			                                	 vcContainer)

				If viRet = 0 Then
				 	vsMsgRetorno = "Processo enviado para execução no servidor!"
				Else
			     	vsMsgRetorno = "Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro
			   	End If

				Set Obj = Nothing

			Else

				vsMsgRetorno = "PEG em processamento, aguarde a conclusão para realizar uma nova solicitação!"

			End If

			Set samPeg = Nothing

			ServiceVar("vsMsgRetorno") = CStr(vsMsgRetorno)

		Else
			Set vvPegConf = CreateBennerObject("SAMPEGCONFERENCIA.Rotinas")
			vsMsgRetorno = vvPegConf.Exec(CurrentSystem, viPeg, psXmlGlosas, psXmlNegacoes, psTipo, -1, 0, 0,111111111)
			If Len(CStr(vsMsgRetorno)) > 0 Then
			  ServiceVar("vsMsgRetorno") = CStr(vsMsgRetorno)
			Else
              ServiceVar("vsMsgRetorno") = "Processo concluído!"
			End If
			Set vvPegConf = Nothing

		End If

		Set SQLAgendado = Nothing
		Set vcContainer = Nothing
		Exit Sub
	erro:
		ServiceVar("vsMsgRetorno") = CStr( Err.Description )
		Set vvPegConf = Nothing
		Set vcContainer = Nothing
End Sub
