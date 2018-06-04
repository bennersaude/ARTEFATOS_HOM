'HASH: 55F56BF03D2BCF747DCE0D048F88FF5D
'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"

Option Explicit

Public Sub BOTAOABRIR_OnClick()
  abrirProtocolo
End Sub

Public Sub BOTAODIGITAR_OnClick()
  SessionVar("HANDLEPROTOCOLOTRANSACAO") = CurrentQuery.FieldByName("HANDLE").AsString
  SessionVar("PROTOCOLOTRANSACAO") = CurrentQuery.FieldByName("HANDLE").AsString
  Dim Interface As Object
  Set Interface = CreateBennerObject("CA043.Autorizacao")
  Interface.Executar(CurrentSystem, 0, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, 0)
  Set Interface = Nothing
  SessionVar("HANDLEPROTOCOLOTRANSACAO") = ""
  SessionVar("PROTOCOLOTRANSACAO") = ""
End Sub

Public Sub BOTAOFECHAR_OnClick()
  fecharProtocolo
End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
	Dim handleRelatorio As String
	Dim nomeParametro As String

  	Dim relatorio As CSReportPrinter
	Dim qhRelatorio As Object
	Set qhRelatorio = NewQuery

	'Atenção à ordem de precedência dos flags, pois pode ter mais de um marcado simultaneamente.
	If (CurrentQuery.FieldByName("CONSULTA").AsString = "S") Then
        nomeParametro = "PRELATORIOCONSULTA"

    ElseIf (CurrentQuery.FieldByName("INTERNACAO").AsString = "S") Then
		nomeParametro = "PROTRELATORIOINTERNACAO"

    ElseIf (CurrentQuery.FieldByName("SPSADT").AsString = "S") Then
		nomeParametro = "PROTRELATORIOSPSADT"

	ElseIf (CurrentQuery.FieldByName("PRORROGACAO").AsString = "S") Then
		nomeParametro = "PROTRELATORIOPRORROGACAO"

	ElseIf (CurrentQuery.FieldByName("COMPLEMENTO").AsString = "S") Then
		nomeParametro = "PROTRELATORIOCOMPLEMENTO"

    ElseIf (CurrentQuery.FieldByName("ODONTOLOGICO").AsString = "S") Then
		nomeParametro = "PROTRELATORIOODONTO"

	ElseIf (CurrentQuery.FieldByName("ANEXOQUIMIO").AsString = "S") Then
		nomeParametro = "PROTRELATORIOQUIMIO"

    ElseIf (CurrentQuery.FieldByName("ANEXORADIO").AsString = "S") Then
		nomeParametro = "PROTRELATORIORADIO"

    ElseIf (CurrentQuery.FieldByName("ANEXOOPME").AsString = "S") Then
      Dim qAnexoOPME As BPesquisa
      Set qAnexoOPME = NewQuery

      qAnexoOPME.Add("SELECT SUM(CASE WHEN ANEX.INTERMEDIACAOCOMPRA = 'S' THEN 0 ELSE 1 END) QTDNAOINTERMEDIADOS,")
      qAnexoOPME.Add("       SUM(CASE WHEN ANEX.INTERMEDIACAOCOMPRA = 'N' THEN 0 ELSE 1 END) QTDINTERMEDIADOS")
      qAnexoOPME.Add("FROM SAM_AUTORIZ_ANEXOOPME ANEX")
      qAnexoOPME.Add("WHERE ANEX.PROTOCOLOTRANSACAO = :HPROTOCOLO")

      qAnexoOPME.ParamByName("HPROTOCOLO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qAnexoOPME.Active = True

      If (qAnexoOPME.FieldByName("QTDNAOINTERMEDIADOS").AsInteger = 0) And _
         (qAnexoOPME.FieldByName("QTDINTERMEDIADOS").AsInteger > 0) Then
        nomeParametro = "PROTRELATORIOINTERMEDIACAOOPME"
      Else
		nomeParametro = "PROTRELATORIOOPME"
	  End If

      Set qAnexoOPME = Nothing
	End If

	If nomeParametro <> "" Then
		qhRelatorio.Clear
		qhRelatorio.Add(" SELECT " + nomeParametro + " HANDLERELATORIO ")
		qhRelatorio.Add("   FROM SAM_PARAMETROSATENDIMENTO")
		qhRelatorio.Active = True

		If qhRelatorio.FieldByName("HANDLERELATORIO").AsInteger > 0 Then
			SessionVar("PROTOCOLOTRANSACAO") = CurrentQuery.FieldByName("HANDLE").AsString
			SessionVar("HANDLEAUTORIZ") = CurrentQuery.FieldByName("AUTORIZACAO").AsString

			Set relatorio = NewReport(qhRelatorio.FieldByName("HANDLERELATORIO").AsInteger)
			relatorio.CanFilter = False
			relatorio.Preview
		Else
			bsShowMessage("Parâmetro geral de atendimento " + nomeParametro + " não preenchido", "I")
		End If
	Else
		bsShowMessage("Protocolo inválido para impressão", "I")
	End If


	Set relatorio = Nothing
	Set qhRelatorio = Nothing
End Sub

Public Sub BOTAONEGAR_OnClick()
	Dim Interface As Object
    Dim vvContainer As CSDContainer
    Dim viRetorno As Integer
    Dim vsMensagemErro As String

   	Set vvContainer = NewContainer

	Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")


   	viRetorno = Interface.Exec(CurrentSystem, _
  						       1, _
                               "TV_MOTIVONEGACAO", _
           	                   "Motivo de Negação Manual", _
               	               0, _
                   	           200, _
                       	       350, _
                           	   False, _
                               vsMensagemErro, _
                               vvContainer)

   Select Case viRetorno
   	 Case -1
    	bsShowMessage("Operação cancelada pelo usuário!", "I")
  	 Case  0
   	 	'bsShowMessage("Opção selecionada" + vvContainer.Field("OPCAO").AsString , "I")
  	 Case  1
   	 	bsShowMessage(vsMensagemErro, "I")
	 End Select

	Set Interface = Nothing
End Sub

Public Sub BOTAOREVALIDAR_OnClick()

   Dim retorno As Integer
   Dim mensagem As String
   Dim alertas As String

   Dim dll As Object
   Set dll=CreateBennerObject("ca043.autorizacao")

   retorno = dll.revalidarAutorizacao(CurrentSystem, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, alertas, mensagem)

   If retorno > 0 Then
      InfoDescription = mensagem
   Else
      bsShowMessage("Revalidação concluída com sucesso.", "I")
   End If

   Set dll=Nothing

End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    SessionVar("HANDLEAUTORIZACAO") = CStr(RetornaNumeroAutorizacao)
  End If

  SessionVar("PROTOCTRANSAUTOR") = CurrentQuery.FieldByName("HANDLE").AsString
  SessionVar("HANDLEEVENTOSOLICIT") = "0"

  If Not WebMode Then
    ROTULORESUMO.Visible = False
  Else
    Dim dllCA043       As Object
    Dim vsTexto        As String
    Dim vsMensagemErro As String
    Set dllCA043 = CreateBennerObject("CA043.Resumo")

    If dllCA043.GerarResumoCriticas(CurrentSystem, _
                                    CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, _
                                    WebVisionCode, _
                                    vsTexto, _
                                    vsMensagemErro, _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger) = 0 Then

      Dim dllCripto As Object
      Set dllCripto = CreateBennerObject("Benner.Saude.Criptografia.Criptografia")
      vsTexto = dllCripto.CriptografaResumoAutoriz(CurrentSystem, vsTexto)
      Set dllCripto = Nothing

      ROTULORESUMO.Text = "@ " + vsTexto
    Else
      ROTULORESUMO.Text = "@ <p class=""frmerror""> <IMG SRC=""img/alert.gif"" />Erro na geração do resumo: " + vsMensagemErro + "</p>"
    End If

    Set dllCA043 = Nothing
  End If

End Sub


Public Sub fecharProtocolo
	Dim dll As Object
	Set dll = CreateBennerObject("SAMAUTO.Autorizador")
	Dim vMensagem As String
	Dim retorno As Integer

	retorno = dll.FecharAutorizacao(CurrentSystem, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, vMensagem)
	If retorno = 1 Then
		bsShowMessage("Erro: "+vMensagem, "I")
	Else
		bsShowMessage("Operação concluída com sucesso!", "I")
	End If
	Set dll = Nothing

End Sub


Public Sub abrirProtocolo
	Dim dll As Object
	Set dll = CreateBennerObject("SAMAUTO.Autorizador")
	Dim vMensagem As String
	Dim retorno As Integer
	retorno = dll.AbrirAutorizacao(CurrentSystem, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, vMensagem)
	If retorno = 1 Then
		bsShowMessage("Erro: "+vMensagem, "I")
	Else
		bsShowMessage("Operação concluída com sucesso!", "I")
	End If
	Set dll = Nothing

End Sub

Public Function RetornaNumeroAutorizacao As Long

  RetornaNumeroAutorizacao = RecordHandleOfTable("SAM_AUTORIZ")

  If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then
    Dim qBuscaHandleAutorizacao As Object
    Set qBuscaHandleAutorizacao  = NewQuery

    qBuscaHandleAutorizacao.Clear
    qBuscaHandleAutorizacao.Add("SELECT AUTORIZACAO FROM SAM_PROTOCOLOTRANSACAOAUTORIZ WHERE HANDLE = :HANDLE")
    qBuscaHandleAutorizacao.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
    qBuscaHandleAutorizacao.Active = True
    RetornaNumeroAutorizacao = qBuscaHandleAutorizacao.FieldByName("AUTORIZACAO").AsInteger

    Set qBuscaHandleAutorizacao  = Nothing
  End If

End Function

Public Sub validarFecharProtocolo
  	Dim mensagemRetorno As String
	mensagemRetorno = ValidarSituacaoAutorizacaoParaFechamento

	If mensagemRetorno <> "OK" Then
	  bsshowmessage(mensagemRetorno, "I")
      Exit Sub
    Else
      Dim SPPFECHARPROTOCOLOTRANSACAO As BStoredProc
	  Set SPPFECHARPROTOCOLOTRANSACAO = NewStoredProc

	  SPPFECHARPROTOCOLOTRANSACAO.AutoMode = True
	  SPPFECHARPROTOCOLOTRANSACAO.Name = "BSAUT_FECHARPROTOCOLOTRANSACAO"
	  SPPFECHARPROTOCOLOTRANSACAO.AddParam("P_AUTORIZACAO",ptInput, ftInteger)
	  SPPFECHARPROTOCOLOTRANSACAO.AddParam("P_ORIGEMPROCESSO",ptInput, ftString)
	  SPPFECHARPROTOCOLOTRANSACAO.AddParam("P_USUARIOFECHAMENTO",ptInput, ftInteger)
	  SPPFECHARPROTOCOLOTRANSACAO.AddParam("P_MENSAGEMRETORNO",ptOutput, ftString)

	  SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
	  SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_ORIGEMPROCESSO").AsString = "A"
	  SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_USUARIOFECHAMENTO").AsInteger = CurrentUser

	  SPPFECHARPROTOCOLOTRANSACAO.ExecProc

      Dim SPPGERARPROTOCOLOAUTORIZACAO As BStoredProc
	  Set SPPGERARPROTOCOLOAUTORIZACAO = NewStoredProc

      SPPGERARPROTOCOLOAUTORIZACAO.AutoMode = True
  	  SPPGERARPROTOCOLOAUTORIZACAO.Name = "BSAUT_GERARPROTOCOLOTRANSACAO"
      SPPGERARPROTOCOLOAUTORIZACAO.AddParam("P_ORIGEMPROCESSO",ptInput, ftString)
	  SPPGERARPROTOCOLOAUTORIZACAO.AddParam("P_AUTORIZACAO",ptInput, ftInteger)
	  SPPGERARPROTOCOLOAUTORIZACAO.AddParam("P_USUARIO",ptInput, ftInteger)
	  SPPGERARPROTOCOLOAUTORIZACAO.AddParam("P_NUMEROPROTOCOLOTRANSACAO",ptInputOutput, ftInteger)
	  SPPGERARPROTOCOLOAUTORIZACAO.AddParam("P_HANDLEATENDIMENTOCENTRAL",ptInputOutput, ftInteger)

      SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_ORIGEMPROCESSO").AsString = "A"
	  SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
	  SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_USUARIO").AsInteger = CurrentUser
	  SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsInteger = 0
	  SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_HANDLEATENDIMENTOCENTRAL").AsInteger = 0
	  SPPGERARPROTOCOLOAUTORIZACAO.ExecProc

	  SessionVar("HANDLEPROTOCOLOTRANSACAO") = SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsString

	  If SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsInteger > 0 Then
        If SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_MENSAGEMRETORNO").AsString = "" Then 'Por enquanto, esta mensagem só indica problema no processo de prorrogação
	      InfoDescription = "Transação finalizada com sucesso! Protocolo gerado: " + SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsString
	    Else
	      InfoDescription = SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_MENSAGEMRETORNO").AsString + " Protocolo gerado para os demais itens: " + SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsString
	    End If

	  Else
        InfoDescription = SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_MENSAGEMRETORNO").AsString
      End If

      Set SPPFECHARPROTOCOLOTRANSACAO = Nothing
	  Set SPPGERARPROTOCOLOAUTORIZACAO = Nothing

    End If
End Sub

Public Function ValidarSituacaoAutorizacaoParaFechamento
  If WebMode Then
    ValidarSituacaoAutorizacaoParaFechamento = "OK"

	Dim qVerificaExistenciaProcedimento As Object
	Set qVerificaExistenciaProcedimento = NewQuery

	qVerificaExistenciaProcedimento.Add("SELECT AO.HANDLE, AOP.EVENTO										  ")
	qVerificaExistenciaProcedimento.Add("  FROM SAM_AUTORIZ_ANEXOOPME      AO 	  							  ")
	qVerificaExistenciaProcedimento.Add("  LEFT JOIN SAM_AUTORIZ_ANEXOOPME_PROC AOP ON (AOP.ANEXOOPME = AO.HANDLE) ")
	qVerificaExistenciaProcedimento.Add("               WHERE AO.AUTORIZACAO = :HANDLE")
	qVerificaExistenciaProcedimento.Add("				  AND AO.PROTOCOLOTRANSACAO IS NULL			 		  ")
	qVerificaExistenciaProcedimento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
	qVerificaExistenciaProcedimento.Active = True

	If Not qVerificaExistenciaProcedimento.FieldByName("HANDLE").IsNull And qVerificaExistenciaProcedimento.FieldByName("EVENTO").IsNull Then
	  ValidarSituacaoAutorizacaoParaFechamento = "O protocolo anexo de OPME sem procedimento. Favor verificar"
	  Set qVerificaExistenciaProcedimento = Nothing
	  Exit Function
	End If

	qVerificaExistenciaProcedimento.Clear
	qVerificaExistenciaProcedimento.Add("SELECT AR.HANDLE, ARP.EVENTO													 ")
	qVerificaExistenciaProcedimento.Add("  FROM SAM_AUTORIZ_ANEXORADIO AR	  							  			     ")
	qVerificaExistenciaProcedimento.Add("  LEFT JOIN SAM_AUTORIZ_ANEXORADIO_PROC   ARP ON (ARP.ANEXORADIOTERAPIA = AR.HANDLE) ")
	qVerificaExistenciaProcedimento.Add(" WHERE AR.AUTORIZACAO = :HANDLE")
	qVerificaExistenciaProcedimento.Add("	AND AR.PROTOCOLOTRANSACAO IS NULL		 				      			   	 ")
	qVerificaExistenciaProcedimento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
	qVerificaExistenciaProcedimento.Active = True

	If Not qVerificaExistenciaProcedimento.FieldByName("HANDLE").IsNull And qVerificaExistenciaProcedimento.FieldByName("EVENTO").IsNull Then
	  ValidarSituacaoAutorizacaoParaFechamento = "O protocolo anexo de radioterapia sem procedimento. Favor verificar"
	  Set qVerificaExistenciaProcedimento = Nothing
	  Exit Function
	End If

	qVerificaExistenciaProcedimento.Clear
	qVerificaExistenciaProcedimento.Add("SELECT AQ.HANDLE, AQP.EVENTO											    ")
	qVerificaExistenciaProcedimento.Add("  FROM SAM_AUTORIZ_ANEXOQUIMIO AQ 							  			   	")
	qVerificaExistenciaProcedimento.Add("  LEFT JOIN SAM_AUTORIZ_ANEXOQUIMIO_PROC   AQP ON (AQP.ANEXOQUIMIO = AQ.HANDLE) ")
	qVerificaExistenciaProcedimento.Add("  WHERE AQ.AUTORIZACAO = :HANDLE")
	qVerificaExistenciaProcedimento.Add("	AND AQ.PROTOCOLOTRANSACAO IS NULL		 				      			")
	qVerificaExistenciaProcedimento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
	qVerificaExistenciaProcedimento.Active = True

	If Not qVerificaExistenciaProcedimento.FieldByName("HANDLE").IsNull And qVerificaExistenciaProcedimento.FieldByName("EVENTO").IsNull Then
	  ValidarSituacaoAutorizacaoParaFechamento = "O protocolo possui anexo de quimioterapia sem procedimento. Favor verificar"
	  Set qVerificaExistenciaProcedimento = Nothing
	  Exit Function
	End If

	Set qVerificaExistenciaProcedimento = Nothing
  End If
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
	  Case "BOTAOFECHAR"
		BOTAOFECHAR_OnClick
	  Case "BOTAOABRIR"
	    BOTAOABRIR_OnClick
	  Case "VALIDAR_FECHAR_PROTOCOLO"

	    If (InStr(SQLServer, "MSSQL") > 0) Then
          CriaTabelaTemporariaSqlServer
        End If

		validarFecharProtocolo
	  Case "BOTAOIMPRIMIR"
	    BOTAOIMPRIMIR_OnClick
	  Case "BOTAOIMPRIMIRANEXOS"
		SessionVar("PROTOCOLOTRANSACAO") = CurrentQuery.FieldByName("HANDLE").AsString
		SessionVar("HANDLEAUTORIZ") = CurrentQuery.FieldByName("AUTORIZACAO").AsString
		SessionVar("AUTORIZACAO") = CurrentQuery.FieldByName("AUTORIZACAO").AsString
	  Case "BOTAOREVISAR"
		Revisar
	  Case "BOTAOREVISARCOMANOTADM"
	  	RevisarComAnotacaoAdministrativa
	  Case "BOTAOREVALIDAR"
		BOTAOREVALIDAR_OnClick
	  Case "BOTAOFECHARAUTORIZACAO"
		fecharAutorizacao
	  Case "BOTAOABRIRATENDIMENTO"
	    abrirAutorizacao
	End Select
End Sub

Public Sub Revisar
	Dim protocoloTransacaoAutorizBLL As CSBusinessComponent

	Set protocoloTransacaoAutorizBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamProtocoloTransacaoAutorizBLL, Benner.Saude.Atendimento.Business")

	protocoloTransacaoAutorizBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger) 'protocoloTransacaoAutoriz

	protocoloTransacaoAutorizBLL.Execute("Revisar")

	Set protocoloTransacaoAutorizBLL = Nothing

	bsShowMessage("Revisão concluída com sucesso", "I")
End Sub

Public Sub RevisarComAnotacaoAdministrativa
	Dim protocoloTransacaoAutorizBLL As CSBusinessComponent

	Set protocoloTransacaoAutorizBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamProtocoloTransacaoAutorizBLL, Benner.Saude.Atendimento.Business")

	protocoloTransacaoAutorizBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger) 'autorizacao
	protocoloTransacaoAutorizBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger) 'protocoloTransacaoAutoriz
	protocoloTransacaoAutorizBLL.AddParameter(pdtInteger, CurrentVirtualQuery.FieldByName("ANOTACAOADMINISTRATIVA").AsInteger) 'anotacaoAdministrativa
	protocoloTransacaoAutorizBLL.AddParameter(pdtAutomatic, CurrentVirtualQuery.FieldByName("ENVIARRELATORIORESPOSTA").AsString = "S") 'enviarNoRelatorioResposta
	protocoloTransacaoAutorizBLL.AddParameter(pdtString, CurrentVirtualQuery.FieldByName("OBSERVACAO").AsString) 'observacao

	protocoloTransacaoAutorizBLL.Execute("Revisar")

	Set protocoloTransacaoAutorizBLL = Nothing

	bsShowMessage("Revisão concluída com sucesso", "I")
End Sub

Public Sub fecharAutorizacao
	Dim dll As Object
	Set dll = CreateBennerObject("SAMAUTO.Autorizador")
	Dim vMensagem As String
	Dim retorno As Integer

	retorno = dll.FecharAutorizacao(CurrentSystem, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, 0, vMensagem)
	If retorno = 1 Then
		bsShowMessage("Erro: "+vMensagem, "I")
	Else
		bsShowMessage("Operação concluída com sucesso!", "I")
	End If
	Set dll = Nothing

End Sub

Public Sub abrirAutorizacao
	Dim dll As Object
	Set dll = CreateBennerObject("SAMAUTO.Autorizador")
	Dim vMensagem As String
	Dim retorno As Integer
	retorno = dll.AbrirAutorizacao(CurrentSystem, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, 0, vMensagem)
	If retorno = 1 Then
		bsShowMessage("Erro: "+vMensagem, "I")
	Else
		bsShowMessage("Operação concluída com sucesso!", "I")
	End If
	Set dll = Nothing

End Sub
