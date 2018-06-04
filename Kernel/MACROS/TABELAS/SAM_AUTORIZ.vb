'HASH: EB3C903E706D8DD27C639C6305C86D1E
'Macro: SAM_AUTORIZ

'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"

Option Explicit

Public Sub BOTAOALTERARDATAATEND_OnClick()
  SessionVar("ALTERARDATAS") = "atendimento"
  MostrarTabelaVirtual
End Sub

Public Sub BOTAOALTERARDATAVALID_OnClick()
    SessionVar("ALTERARDATAS") = "validade"
	MostrarTabelaVirtual
End Sub

Public Sub BOTAODIGITAR_OnClick()
  Dim Interface As Object
  Set Interface = CreateBennerObject("CA043.Autorizacao")
  Interface.Executar(CurrentSystem, 0, CurrentQuery.FieldByName("HANDLE").AsInteger, 0)
  Set Interface = Nothing
End Sub

Public Sub BOTAOINTERCORRENCIA_OnClick()
  Dim CanContinue As Boolean
  If (CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "2") Or (CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "3") Then
    bsShowMessage("Autorização sendo reprocessada de forma agendada. Favor tentar a edição mais tarde", "E")
    CanContinue =False
  End If
End Sub

Public Sub TABLE_AfterScroll()
  SessionVar("HANDLEAUTORIZACAO") = CurrentQuery.FieldByName("HANDLE").AsString
  SessionVar("HANDLEAUTORIZACAOREL") = CurrentQuery.FieldByName("HANDLE").AsString
  SessionVar("PROTOCTRANSAUTOR") = "0"

  BOTAOABRIR.Enabled = False
  BOTAOACIDENTETRABALHO.Enabled = False
  BOTAOCANCELAR.Enabled = False
  BOTAODATAPRESTACAOCONTAS.Enabled = False
  BOTAOFATURARPF.Enabled = False
  BOTAOFECHAR.Enabled = False
  BOTAOGERARDIARIAS.Enabled = False
  BOTAOGERARFATURAADIANTAMENTO.Enabled = False
  BOTAOIMPRIMIR.Enabled = False
  BOTAOINTERCORRENCIA.Enabled = False
  BOTAOORCAMENTO.Enabled = False
  BOTAOPFINTEGRAL.Enabled = False
  BOTAOPRORROGARINTERNACAO.Enabled = False
  BOTAOREIMPRESSAO.Enabled = False
  BOTAOINSERIREVENTO.Enabled = False
  BOTAOINSERIREVENTOODONTO.Enabled = False

  If Not (CurrentQuery.FieldByName("TABRESPOSTA").AsInteger = 3) Then
    DATAHORARESPOSTA.Visible = False
    USUARIORESPOSTA.Visible = False
    SITUACAORESPOSTA.Visible = False
    STATUSFAX.Visible = False
  Else
    DATAHORARESPOSTA.Visible = True
    USUARIORESPOSTA.Visible = True
    SITUACAORESPOSTA.Visible = True
    STATUSFAX.Visible = True
  End If
  'Fim - SMS 47394

    If Not (CurrentQuery.FieldByName("TABRESPOSTAPREST").AsInteger = 3) Then ' SMS 75903 - Julio - 25/01/2007 - Inclusão do NOT
    DATAHORARESPOSTAPREST.Visible = False
    USUARIORESPOSTAPREST.Visible = False
    SITUACAORESPOSTAPREST.Visible = False
    STATUSFAXPREST.Visible = False
  Else
    DATAHORARESPOSTAPREST.Visible = True
    USUARIORESPOSTAPREST.Visible = True
    SITUACAORESPOSTAPREST.Visible = True
    STATUSFAXPREST.Visible = True
  End If

  If VisibleMode Then
    If Not CurrentQuery.IsVirtual Then
      TABLE.Pages("RESUMO").Visible = False
    End If
  Else
    Dim dllCA043       As Object
    Dim vsTexto        As String
    Dim vsMensagemErro As String
    Set dllCA043 = CreateBennerObject("CA043.Resumo")

    If dllCA043.GerarResumoCriticas(CurrentSystem, _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    WebVisionCode, _
                                    vsTexto, _
                                    vsMensagemErro, _
                                    0) = 0 Then

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


   'Atualizar rótulo com informações de diárias e prorrogações:
	Dim qTotalProrrogacoes As Object
	Set qTotalProrrogacoes = NewQuery

	qTotalProrrogacoes.Add("SELECT COALESCE(SUM(QTDAUTORIZADA),0) TOTALPRORROGACOES ")
	qTotalProrrogacoes.Add("  FROM SAM_AUTORIZ_EVENTOGERADO                         ")
	qTotalProrrogacoes.Add(" WHERE AUTORIZACAO = :AUTORIZACAO                       ")
	qTotalProrrogacoes.Add("   AND TIPOEVENTO = :TIPOEVENTO                         ")

	qTotalProrrogacoes.ParamByName("AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qTotalProrrogacoes.ParamByName("TIPOEVENTO").AsString = "P"
	qTotalProrrogacoes.Active = True

    ROTULODIARIAS.Text = "Diárias: " + CurrentQuery.FieldByName("DIARIASLIBERADAS").AsString
    ROTULOPRORROGACAO.Text = "Prorrogação: " + qTotalProrrogacoes.FieldByName("TOTALPRORROGACOES").AsString
    ROTULOTOTALDIARIAS.Text = "Total diárias: " + CStr(CurrentQuery.FieldByName("DIARIASLIBERADAS").AsInteger + qTotalProrrogacoes.FieldByName("TOTALPRORROGACOES").AsInteger)

    Set qTotalProrrogacoes = Nothing

End Sub



Public Sub gerarDiarias()
	Dim retorno As Integer
	Dim mensagem As String
	Dim dll As Object
	Set dll=CreateBennerObject("ca043.Autorizacao")
	retorno = dll.GerarDiarias(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, mensagem)
	Set dll=Nothing
	If retorno>0 Then
		InfoDescription = mensagem
	Else
		InfoDescription = "Geração de diárias concluída com sucesso"
	End If

End Sub

Public Sub prorrogarInternacao()
	Dim retorno As Integer
	Dim mensagem As String
	Dim dll As Object
	Set dll=CreateBennerObject("ca043.Autorizacao")
	retorno = dll.prorrogarInternacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, mensagem)
	Set dll=Nothing
	If retorno>0 Then
		InfoDescription = mensagem
	Else
		InfoDescription = "Prorrogação concluída com sucesso"
	End If

End Sub


Public Sub imprimir()

	If WebMode Then
		SessionVar("WEBHandleFiltro") = CurrentQuery.FieldByName("HANDLE").AsString
	End If

	Dim retorno As Integer
	Dim mensagem As String
	Dim dll As Object
	Set dll=CreateBennerObject("samauto.autorizador")
	retorno = dll.imprimir(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, mensagem)
	Set dll=Nothing
	If retorno>0 Then
		InfoDescription = mensagem
	Else
		InfoDescription = "Impressão concluída com sucesso"
	End If

End Sub

Public Sub reimprimir()
	Dim sql As BPesquisa
	Set sql=NewQuery
	sql.Add("DELETE FROM TMP_AUTORIZREIMPRESSAO WHERE USUARIO=:USUARIO")
	sql.ParamByName("USUARIO").AsInteger=CurrentUser
	sql.ExecSQL

	sql.Clear
	sql.Add("INSERT INTO TMP_AUTORIZREIMPRESSAO (HANDLE, USUARIO, AUTORIZACAO, NUMEROGUIA, EVENTO, GRAU) VALUES (:H, :U, :A, :N, :E, :G)")

	Dim guias As BPesquisa
	Set guias=NewQuery
	guias.Add("SELECT DISTINCT(EG.NUMEROGUIA) NUMEROGUIA,")
	guias.Add("       EG.EVENTOGERADO,")
	guias.Add("       ES.GRAU")
	guias.Add("  FROM SAM_AUTORIZ_EVENTOGERADO EG")
	guias.Add("  Join SAM_AUTORIZ_EVENTOSOLICIT ES On (ES.Handle=EG.EVENTOSOLICITADO)")
	guias.Add(" WHERE ES.AUTORIZACAO=:AUTORIZACAO")
	guias.Add("   And EG.NUMEROGUIA Is Not Null")
	guias.Add("   And (EG.SITUACAO = 'A' OR EG.SITUACAO='L')")
	guias.ParamByName("AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	guias.Active=True
	While Not guias.EOF
		sql.ParamByName("H").AsInteger = NewHandle("TMP_AUTORIZREIMPRESSAO")
		sql.ParamByName("U").AsInteger = CurrentUser
		sql.ParamByName("A").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		sql.ParamByName("N").AsInteger = guias.FieldByName("NUMEROGUIA").AsInteger
		sql.ParamByName("E").AsInteger = guias.FieldByName("EVENTOGERADO").AsInteger
		sql.ParamByName("G").AsInteger = guias.FieldByName("GRAU").AsInteger
		sql.ExecSQL
		guias.Next
	Wend


	Set guias=Nothing
	Set sql=Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

If (CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "2") Or (CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "3") Then
  bsShowMessage("Autorização sendo reprocessada de forma agendada. Favor tentar a edição mais tarde", "E")
  CanContinue =False
End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAODIGITAR") Then
		BOTAODIGITAR_OnClick
	ElseIf (CommandID = "BOTAOFECHAR") Then
		fecharAutorizacao
	ElseIf (CommandID = "BOTAOGERARDIARIAS") Then
		gerarDiarias
	ElseIf (CommandID = "BOTAOPRORROGARINTERNACAO") Then
		prorrogarInternacao
	ElseIf (CommandID = "BOTAOIMPRIMIR") Then
		imprimir
	ElseIf (CommandID = "BOTAOREIMPRESSAO") Then
		reimprimir
    ElseIf (CommandID = "BOTAOABRIR") Then
		abrirAutorizacao
	ElseIf (CommandID = "VALIDAR_FECHAR_PROTOCOLO") Then
		If (InStr(SQLServer, "MSSQL") > 0) Then
          CriaTabelaTemporariaSqlServer
        End If
		validarFecharProtocolo
	ElseIf (CommandID = "VALIDAR_FECHAR_SOLICITACAO") Then
		If (InStr(SQLServer, "MSSQL") > 0) Then
          CriaTabelaTemporariaSqlServer
        End If
		validarFecharSolicitacao
	ElseIf (CommandID = "BOTAOIMPRIMIRPRORROGACAO") Then
        ImprimirProrrogacao
	ElseIf (CommandID = "VALIDAR_ANEXOS") Then

		If (InStr(SQLServer, "MSSQL") > 0) Then
          CriaTabelaTemporariaSqlServer
        End If

		validarFecharProtocolo
	ElseIf CommandID = "CANCELARCOMUNICINTERNACAO" Then
 		CancelarComunicacaoInternacao
 	ElseIf CommandID = "CANCELARCOMUNICALTA" Then
 		CancelarComunicacaoAlta
 	ElseIf CommandID = "CANCELARCOMUNICFECHAMENTOPARCIAL" Then
 		CancelarComunicacaoFechamentoParcial
 	ElseIf CommandID = "IMPRIMIRRELATORIOCAPEANTE" Then
 		ImprimirRelatorioCapeante
	End If
End Sub

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

	  SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
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
	  SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_USUARIO").AsInteger = CurrentUser
	  SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsInteger = 0
	  SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_HANDLEATENDIMENTOCENTRAL").AsInteger = 0
	  SPPGERARPROTOCOLOAUTORIZACAO.ExecProc

	  SessionVar("HANDLEPROTOCOLOTRANSACAO") = SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsString

      Dim qParametrosAtendimento As Object
      Set qParametrosAtendimento = NewQuery

	  qParametrosAtendimento.Add("SELECT TABUTILIZACONCEITOPROTOCOLO,")
      qParametrosAtendimento.Add("       PROTRETORNOSOLICPROC        ")
	  qParametrosAtendimento.Add("  FROM SAM_PARAMETROSATENDIMENTO   ")
	  qParametrosAtendimento.Active = True

      If (qParametrosAtendimento.FieldByName("TABUTILIZACONCEITOPROTOCOLO").AsInteger = 2) And (qParametrosAtendimento.FieldByName("PROTRETORNOSOLICPROC").AsString = "P") Then

  	    If SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsInteger > 0 Then

	      If SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_MENSAGEMRETORNO").AsString = "" Then 'Por enquanto, esta mensagem só indica problema no processo de geração de prorrogação.
	        InfoDescription = "Transação finalizada com sucesso! Protocolo gerado: " + SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsString
	      Else
	        InfoDescription = SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_MENSAGEMRETORNO").AsString + " Protocolo gerado para os demais itens: " + SPPGERARPROTOCOLOAUTORIZACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsString
	      End If

  	    Else
	      InfoDescription = SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_MENSAGEMRETORNO").AsString
	    End If

      Else

        InfoDescription = RetornaSituacaoAutorizacao

      End If

      Set SPPFECHARPROTOCOLOTRANSACAO = Nothing
	  Set SPPGERARPROTOCOLOAUTORIZACAO = Nothing
	  Set qParametrosAtendimento = Nothing

    End If
End Sub
Public Sub validarFecharSolicitacao

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

	  SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_ORIGEMPROCESSO").AsString = "A"
	  SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_USUARIOFECHAMENTO").AsInteger = CurrentUser

	  SPPFECHARPROTOCOLOTRANSACAO.ExecProc


      Dim qParametrosAtendimento As Object
      Set qParametrosAtendimento = NewQuery

	  qParametrosAtendimento.Add("SELECT TABUTILIZACONCEITOPROTOCOLO,")
      qParametrosAtendimento.Add("       PROTRETORNOSOLICPROC        ")
	  qParametrosAtendimento.Add("  FROM SAM_PARAMETROSATENDIMENTO   ")
	  qParametrosAtendimento.Active = True

      If (qParametrosAtendimento.FieldByName("TABUTILIZACONCEITOPROTOCOLO").AsInteger = 2) And (qParametrosAtendimento.FieldByName("PROTRETORNOSOLICPROC").AsString = "P") Then

	      If SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_MENSAGEMRETORNO").AsString = "" Then 'Por enquanto, esta mensagem só indica problema no processo de geração de prorrogação.
	        InfoDescription = "Transação finalizada com sucesso!"
	      Else
	        InfoDescription = SPPFECHARPROTOCOLOTRANSACAO.ParamByName("P_MENSAGEMRETORNO").AsString
	      End If

      Else

        InfoDescription = RetornaSituacaoAutorizacao

      End If

      Set SPPFECHARPROTOCOLOTRANSACAO = Nothing
	  Set qParametrosAtendimento = Nothing

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
	qVerificaExistenciaProcedimento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qVerificaExistenciaProcedimento.Active = True

	If Not qVerificaExistenciaProcedimento.FieldByName("HANDLE").IsNull And qVerificaExistenciaProcedimento.FieldByName("EVENTO").IsNull Then
	  ValidarSituacaoAutorizacaoParaFechamento = "A autorização possui anexo de OPME sem procedimento. Favor verificar"
	  Set qVerificaExistenciaProcedimento = Nothing
	  Exit Function
	End If

	qVerificaExistenciaProcedimento.Clear
	qVerificaExistenciaProcedimento.Add("SELECT AR.HANDLE, ARP.EVENTO													 ")
	qVerificaExistenciaProcedimento.Add("  FROM SAM_AUTORIZ_ANEXORADIO AR	  							  			     ")
	qVerificaExistenciaProcedimento.Add("  LEFT JOIN SAM_AUTORIZ_ANEXORADIO_PROC   ARP ON (ARP.ANEXORADIOTERAPIA = AR.HANDLE) ")
	qVerificaExistenciaProcedimento.Add(" WHERE AR.AUTORIZACAO = :HANDLE")
	qVerificaExistenciaProcedimento.Add("	AND AR.PROTOCOLOTRANSACAO IS NULL		 				      			   	 ")
	qVerificaExistenciaProcedimento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qVerificaExistenciaProcedimento.Active = True

	If Not qVerificaExistenciaProcedimento.FieldByName("HANDLE").IsNull And qVerificaExistenciaProcedimento.FieldByName("EVENTO").IsNull Then
	  ValidarSituacaoAutorizacaoParaFechamento = "A autorização possui anexo de radioterapia sem procedimento. Favor verificar"
	  Set qVerificaExistenciaProcedimento = Nothing
	  Exit Function
	End If

	qVerificaExistenciaProcedimento.Clear
	qVerificaExistenciaProcedimento.Add("SELECT AQ.HANDLE, AQP.EVENTO											    ")
	qVerificaExistenciaProcedimento.Add("  FROM SAM_AUTORIZ_ANEXOQUIMIO AQ 							  			   	")
	qVerificaExistenciaProcedimento.Add("  LEFT JOIN SAM_AUTORIZ_ANEXOQUIMIO_PROC   AQP ON (AQP.ANEXOQUIMIO = AQ.HANDLE) ")
	qVerificaExistenciaProcedimento.Add("  WHERE AQ.AUTORIZACAO = :HANDLE")
	qVerificaExistenciaProcedimento.Add("	AND AQ.PROTOCOLOTRANSACAO IS NULL		 				      			")
	qVerificaExistenciaProcedimento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qVerificaExistenciaProcedimento.Active = True

	If Not qVerificaExistenciaProcedimento.FieldByName("HANDLE").IsNull And qVerificaExistenciaProcedimento.FieldByName("EVENTO").IsNull Then
	  ValidarSituacaoAutorizacaoParaFechamento = "A autorização possui anexo de quimioterapia sem procedimento. Favor verificar"
	  Set qVerificaExistenciaProcedimento = Nothing
	  Exit Function
	End If

	Set qVerificaExistenciaProcedimento = Nothing
  End If
End Function

Public Function RetornaSituacaoAutorizacao
  Dim qSituacaoAutorizacao As Object
  Set qSituacaoAutorizacao = NewQuery

  qSituacaoAutorizacao.Clear
  qSituacaoAutorizacao.Add(" SELECT N.HANDLE                                                    ")
  qSituacaoAutorizacao.Add("   FROM SAM_AUTORIZ_EVENTONEGACAO N                                 ")
  qSituacaoAutorizacao.Add("   JOIN SAM_AUTORIZ_EVENTOGERADO EG ON (N.EVENTOGERADO = EG.HANDLE) ")
  qSituacaoAutorizacao.Add("  WHERE EG.AUTORIZACAO = :AUTORIZACAO                               ")
  qSituacaoAutorizacao.Add("    and N.SITUACAO <> 'R'                                           ")
  qSituacaoAutorizacao.ParamByName("AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSituacaoAutorizacao.Active = True

  If Not qSituacaoAutorizacao.EOF Then
    RetornaSituacaoAutorizacao = "Operação concluída: Autorização NEGADA."
  Else
    qSituacaoAutorizacao.Clear
    qSituacaoAutorizacao.Add(" SELECT HANDLE FROM SAM_AUTORIZ_EVENTOGERADO WHERE AUTORIZACAO = :AUTORIZACAO ")
    qSituacaoAutorizacao.ParamByName("AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSituacaoAutorizacao.Active = True
    If qSituacaoAutorizacao.EOF Then
      RetornaSituacaoAutorizacao = "Autorização não gerou eventos. Verifique eventos complementares"
    Else
      qSituacaoAutorizacao.Clear
      qSituacaoAutorizacao.Add(" SELECT N.HANDLE                                                    ")
      qSituacaoAutorizacao.Add("   FROM SAM_AUTORIZ_EVENTONEGACAO N                                 ")
      qSituacaoAutorizacao.Add("   JOIN SAM_AUTORIZ_EVENTOGERADO EG ON (N.EVENTOGERADO = EG.HANDLE) ")
      qSituacaoAutorizacao.Add("  WHERE EG.AUTORIZACAO = :AUTORIZACAO                               ")
      qSituacaoAutorizacao.Add("    and ( N.SITUACAO not IN ('A', 'L'))                             ")
      qSituacaoAutorizacao.ParamByName("AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qSituacaoAutorizacao.Active = True

	  If qSituacaoAutorizacao.EOF Then
        RetornaSituacaoAutorizacao = "Operação concluída: AUTORIZADA."
      Else
        RetornaSituacaoAutorizacao = "Operação efetuada com sucesso"
      End If
    End If
  End If

  Set qSituacaoAutorizacao = Nothing

End Function

Public Sub fecharAutorizacao
	Dim dll As Object
	Set dll = CreateBennerObject("SAMAUTO.Autorizador")
	Dim vMensagem As String
	Dim retorno As Integer
	retorno = dll.FecharAutorizacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, vMensagem)
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
	retorno = dll.AbrirAutorizacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, vMensagem)
	If retorno = 1 Then
		bsShowMessage("Erro: "+vMensagem, "I")
	Else
		bsShowMessage("Operação concluída com sucesso!", "I")
	End If
	Set dll = Nothing

End Sub

Public Sub ImprimirProrrogacao
	If WebMode Then
	  SessionVar("HANDLEAUTORIZIMPRIMIRPRORROGACAOWEB") = CurrentQuery.FieldByName("HANDLE").AsString
	End If

	Dim retorno As Integer
	Dim mensagem As String
	Dim dll As Object
	Set dll = CreateBennerObject("samauto.autorizador")
	retorno = dll.imprimirProrrogacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, mensagem)
	Set dll=Nothing
	If retorno > 0 Then
		InfoDescription = mensagem
	Else
		InfoDescription = "Impressão concluída com sucesso"
	End If

	If WebMode Then
	  SessionVar("HANDLEAUTORIZIMPRIMIRPRORROGACAOWEB") = ""
	End If
End Sub

Public Function MostrarTabelaVirtual
	Dim INTERFACE0002 As Object
	Dim vsMensagem As String
	Dim vcContainer As CSDContainer


	Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

	INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_NOVADATAEVENTOS", _
					   "Alterar data dos eventos",  _
					   0, _
					   165, _
					   280, _
					   False, _
					   vsMensagem, _
					   vcContainer)

	Set INTERFACE0002 = Nothing
End Function

Public Sub CancelarComunicacaoInternacao
 	Dim samAutorizBLL As CSBusinessComponent
 	Dim resultado As String
 	Set samAutorizBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamAutorizBLL, Benner.Saude.Atendimento.Business")
 	samAutorizBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
 	resultado = CStr(samAutorizBLL.Execute("CancelarComunicacaoInternacao"))
 	If resultado <> "" Then
 		bsShowMessage(resultado, "I")
 	End If
End Sub

Public Sub CancelarComunicacaoAlta
 	Dim samAutorizBLL As CSBusinessComponent
 	Dim resultado As String
 	Set samAutorizBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamAutorizBLL, Benner.Saude.Atendimento.Business")
 	samAutorizBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
 	resultado = CStr(samAutorizBLL.Execute("CancelarComunicacaoAlta"))
 	If resultado <> "" Then
 		bsShowMessage(resultado, "I")
 	End If
End Sub

Public Sub CancelarComunicacaoFechamentoParcial
 	Dim samAutorizBLL As CSBusinessComponent
 	Dim resultado As String
 	Set samAutorizBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamAutorizBLL, Benner.Saude.Atendimento.Business")
 	samAutorizBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
 	resultado = CStr(samAutorizBLL.Execute("CancelarFechamentoParcial"))
 	If resultado <> "" Then
 		bsShowMessage(resultado, "I")
 	End If
End Sub

Public Sub ImprimirRelatorioCapeante
 	Dim qParametrosAtendimento As BPesquisa
 	Set qParametrosAtendimento = NewQuery

 	qParametrosAtendimento.Add("SELECT RELATORIOCAPEANTE")
 	qParametrosAtendimento.Add("FROM SAM_PARAMETROSATENDIMENTO")
 	qParametrosAtendimento.Active = True

 	If qParametrosAtendimento.FieldByName("RELATORIOCAPEANTE").IsNull Then
 		bsshowmessage("Relatório Capeante não foi configurado", "I")
 	Else
 		SessionVar("HAUTORIZCAPEANTE") = CurrentQuery.FieldByName("HANDLE").AsString
 		ReportPreview(qParametrosAtendimento.FieldByName("RELATORIOCAPEANTE").AsInteger,"", False, False)
 	End If

 	Set qParametrosAtendimento = Nothing
End Sub
