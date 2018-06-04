'HASH: 2D80BD61A452D454994365B7AF510E7F
'MACRO: TV_MOTIVONEGACAO'
'#Uses "*bsShowMessage"

Dim vHandleEventoSolicit As Long
Dim vHandleProtocolo As Long
Dim vHandleAutorizacao As Long

Option Explicit

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
    WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforeInsert - Início")
    Dim vMensagem As String
    Dim vMsgRetorno As String

	vHandleProtocolo   = CLng(SessionVar("PROTOCTRANSAUTOR"))
    vHandleAutorizacao = CLng(SessionVar("HANDLEAUTORIZACAO"))
	vMensagem      = ""
	vMsgRetorno    = ""

    If  (SessionVar("HANDLEEVENTOSOLICIT") <> "0" Or SessionVar("HANDLEEVENTOSOLICIT") <> Null Or SessionVar("HANDLEEVENTOSOLICIT") <> "")  Then
		vHandleEventoSolicit = CLng(SessionVar("HANDLEEVENTOSOLICIT"))
	Else
	    vHandleEventoSolicit = RecordHandleOfTable("SAM_AUTORIZ_EVENTOSOLICIT")
	End If

	Dim interface As Object
	Set interface = CreateBennerObject("ESPECIFICO.uEspecifico")

    If vHandleProtocolo > 0 And vHandleEventoSolicit = 0 Then
	   WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforeInsert - Por Protocolo")
       Dim selecionaEventoAutoriz As BPesquisa
	   Set selecionaEventoAutoriz = NewQuery

	   selecionaEventoAutoriz.Active = False
	   selecionaEventoAutoriz.Clear
	   selecionaEventoAutoriz.Add("SELECT AA.ESTRUTURA, AA.DESCRICAO, A.HANDLE, A.AUTORIZACAO ")
       selecionaEventoAutoriz.Add("  FROM SAM_AUTORIZ_EVENTOSOLICIT A ")
       selecionaEventoAutoriz.Add("  Left Join SAM_TGE AA On A.EVENTO = AA.Handle ")
       selecionaEventoAutoriz.Add(" WHERE A.AUTORIZACAO = :HANDLE     ")
       selecionaEventoAutoriz.Add("  AND A.HANDLE IN (SELECT OPME.EVENTOSOLICIT ")
       selecionaEventoAutoriz.Add("                     FROM SAM_AUTORIZ_ANEXOOPME A")
       selecionaEventoAutoriz.Add("                     JOIN SAM_AUTORIZ_ANEXOOPME_PROC OPME ON A.HANDLE = OPME.ANEXOOPME")
       selecionaEventoAutoriz.Add("                    WHERE A.PROTOCOLOTRANSACAO = :PROTOCOLO")
       selecionaEventoAutoriz.Add("                   UNION")
       selecionaEventoAutoriz.Add("                   SELECT QUIM.EVENTOSOLICIT")
       selecionaEventoAutoriz.Add("                     FROM SAM_AUTORIZ_ANEXOQUIMIO A")
       selecionaEventoAutoriz.Add("                     JOIN SAM_AUTORIZ_ANEXOQUIMIO_PROC QUIM ON A.HANDLE = QUIM.ANEXOQUIMIO")
       selecionaEventoAutoriz.Add("                    WHERE A.PROTOCOLOTRANSACAO = :PROTOCOLO")
       selecionaEventoAutoriz.Add("                   UNION")
       selecionaEventoAutoriz.Add("                   SELECT RADI.EVENTOSOLICIT")
       selecionaEventoAutoriz.Add("                     FROM SAM_AUTORIZ_ANEXORADIO A")
       selecionaEventoAutoriz.Add("                     JOIN SAM_AUTORIZ_ANEXORADIO_PROC RADI ON A.HANDLE = RADI.ANEXORADIOTERAPIA")
       selecionaEventoAutoriz.Add("                    WHERE A.PROTOCOLOTRANSACAO = :PROTOCOLO ")
       selecionaEventoAutoriz.Add("                   UNION")
       selecionaEventoAutoriz.Add("                   SELECT EVENTOSOLICIT")
       selecionaEventoAutoriz.Add("                     FROM SAM_AUTORIZ_COMPLEMENTO")
       selecionaEventoAutoriz.Add("                    WHERE PROTOCOLOTRANSACAO = :PROTOCOLO ")
       selecionaEventoAutoriz.Add("                   UNION ")
       selecionaEventoAutoriz.Add("                   SELECT EVENTOSOLICITADO EVENTOSOLICIT  ")
       selecionaEventoAutoriz.Add("                     FROM SAM_AUTORIZ_EVENTOGERADO        ")
       selecionaEventoAutoriz.Add("                    WHERE PROTOCOLOTRANSACAO = :PROTOCOLO ")
       selecionaEventoAutoriz.Add("                  )")
       selecionaEventoAutoriz.ParamByName("HANDLE").AsInteger    = vHandleAutorizacao
       selecionaEventoAutoriz.ParamByName("PROTOCOLO").AsInteger = vHandleProtocolo

	   selecionaEventoAutoriz.Active = True
	   selecionaEventoAutoriz.First

       While Not selecionaEventoAutoriz.EOF

	      vHandleEventoSolicit = selecionaEventoAutoriz.FieldByName("handle").AsInteger

	      If Not interface.BCB_ATE_AutorizacaoJaFinanciada(CurrentSystem, selecionaEventoAutoriz.FieldByName("AUTORIZACAO").AsInteger) Then

		     vMensagem = Negar(selecionaEventoAutoriz.FieldByName("handle").AsInteger, vHandleProtocolo)

		     If vMensagem <> "" Then
	            vMsgRetorno = vMsgRetorno + "***" + selecionaEventoAutoriz.FieldByName("ESTRUTURA").AsString+"-"+selecionaEventoAutoriz.FieldByName("DESCRICAO").AsString + "***" + Chr(13) + "*" + vMensagem + "*" + Chr(13)
	         End If

	      End If

	      selecionaEventoAutoriz.Next
	   Wend

	   Set selecionaEventoAutoriz = Nothing

	   If vMsgRetorno <> "" Then
	  	  bsShowMessage(vMsgRetorno, "E")
		  CanContinue = False
		  Set selecionaEventoAutoriz = Nothing
		  Exit Sub
	   End If


    Else
	   WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforeInsert - Por evento solicitado")
	   Dim selecionaAutorizEvento As BPesquisa
	   Set selecionaAutorizEvento = NewQuery

	   selecionaAutorizEvento.Active = False
	   selecionaAutorizEvento.Clear
	   selecionaAutorizEvento.Add("SELECT AUTORIZACAO               ")
	   selecionaAutorizEvento.Add("  FROM SAM_AUTORIZ_EVENTOSOLICIT ")
	   selecionaAutorizEvento.Add(" WHERE HANDLE = :HANDLE          ")
	   selecionaAutorizEvento.ParamByName("HANDLE").AsInteger = vHandleEventoSolicit
	   selecionaAutorizEvento.Active = True

	   If Not interface.BCB_ATE_AutorizacaoJaFinanciada(CurrentSystem, selecionaAutorizEvento.FieldByName("AUTORIZACAO").AsInteger) Then
          vMsgRetorno=Negar(vHandleEventoSolicit, vHandleProtocolo)
          If vMsgRetorno <> "" Then
	  	     bsShowMessage(vMsgRetorno, "E")
		     CanContinue = False
		     Set selecionaAutorizEvento = Nothing
		     Exit Sub
	      End If
	   End If

	   Set selecionaAutorizEvento = Nothing

	End If

	Set interface=Nothing
	WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforeInsert - Fim")
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Início")
	Dim vMsgRetorno As String
	Dim interface As Object

	Set interface = CreateBennerObject("CA043.Autorizacao")

    If vHandleProtocolo > 0 Then
		WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Por Protocolo")
		If vHandleEventoSolicit > 0 Then
			WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Somente de um evento solicitado")
			If bsShowMessage("Deseja negar todos os eventos gerados do evento solicitado?", "Q") = vbYes Then
				WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Negar eventos")
				If Not interface.ProcessaNegar(CurrentSystem, vHandleEventoSolicit, CurrentQuery.FieldByName("MOTIVONEGACAOMANUAL").AsInteger , vMsgRetorno, vHandleProtocolo) Then
					bsShowMessage(vMsgRetorno, "I")
				Else
					bsShowMessage("Todos os eventos gerados do evento solicitado foram negados.", "I")
				End If
    		End If
		Else
			WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Todos os eventos")
       		If bsShowMessage("Deseja negar todos os eventos gerados do protocolo?", "Q") = vbYes Then
    			WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Negar eventos")
          		vMsgRetorno = interface.NegarEventosDoProtocolo(CurrentSystem, vHandleAutorizacao, vHandleProtocolo, CurrentUser,CurrentQuery.FieldByName("MOTIVONEGACAOMANUAL").AsInteger)
	      		If vMsgRetorno <> "" Then
					bsShowMessage(vMsgRetorno, "I")
	      		Else
					bsShowMessage("Todos os eventos gerados do protocolo foram negados.", "I")
	      		End If
       		End If
    	End If
    Else
    	WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Por evento solicitado")
		If bsShowMessage("Deseja negar todos os eventos gerados do evento solicitado?", "Q") = vbYes Then
			WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Negar eventos")
			If Not interface.ProcessaNegar(CurrentSystem, vHandleEventoSolicit, CurrentQuery.FieldByName("MOTIVONEGACAOMANUAL").AsInteger , vMsgRetorno) Then
				bsShowMessage(vMsgRetorno, "I")
			Else
				bsShowMessage("Todos os eventos gerados do evento solicitado foram negados.", "I")
			End If
    	End If
    End If

	Set interface=Nothing
	WriteBDebugMessage("TV_MOTIVONEGACAO.TABLE_BeforePost - Fim")
End Sub

Public Function Negar(pHandleEvento As Long, pHandleProtocolo As Long) As String
	WriteBDebugMessage("TV_MOTIVONEGACAO.Negar - Início")

	Dim sqlSituacaoEventoSolicit As Object
	Dim sqlSituacaoEventoGerado As Object
	Dim sqlEventoPago As Object
	Dim sqlParametroAtendimento As Object
	Dim interface As Object
	Dim especifico As Object

	Set interface = CreateBennerObject("CA043.Autorizacao")
	Set especifico = CreateBennerObject("ESPECIFICO.uEspecifico")

	Set sqlSituacaoEventoSolicit = NewQuery
	Set sqlSituacaoEventoGerado = NewQuery
	Set sqlEventoPago = NewQuery
	Set sqlParametroAtendimento = NewQuery

	sqlSituacaoEventoSolicit.Active = False
	sqlSituacaoEventoSolicit.Clear
	sqlSituacaoEventoSolicit.Add("SELECT SITUACAO                  ")
	sqlSituacaoEventoSolicit.Add("  FROM SAM_AUTORIZ_EVENTOSOLICIT ")
	sqlSituacaoEventoSolicit.Add(" WHERE HANDLE = :HANDLE          ")
	sqlSituacaoEventoSolicit.ParamByName("HANDLE").AsInteger = pHandleEvento
  sqlSituacaoEventoSolicit.Active = True

  sqlSituacaoEventoGerado.Active = False
  sqlSituacaoEventoGerado.Add("SELECT COUNT(1) QTD                       ")
  sqlSituacaoEventoGerado.Add("  FROM SAM_AUTORIZ_EVENTOGERADO           ")
  sqlSituacaoEventoGerado.Add(" WHERE EVENTOSOLICITADO = :EVENTOSOLICIT  ")
  sqlSituacaoEventoGerado.ParamByName("EVENTOSOLICIT").AsInteger = pHandleEvento
  sqlSituacaoEventoGerado.Active = True

  sqlEventoPago.Active = False
  sqlEventoPago.Clear
  sqlEventoPago.Add("SELECT COUNT(1) QTD                      ")
  sqlEventoPago.Add("  FROM SAM_AUTORIZ_EVENTOGERADO          ")
  sqlEventoPago.Add(" WHERE EVENTOSOLICITADO = :EVENTOSOLICIT ")
  sqlEventoPago.Add("   AND QTDPAGA > 0                       ")
  sqlEventoPago.ParamByName("EVENTOSOLICIT").AsInteger = pHandleEvento
  sqlEventoPago.Active = True

  sqlParametroAtendimento.Active = False
  sqlParametroAtendimento.Clear
  sqlParametroAtendimento.Add("SELECT FORNECIMENTOMEDICAMENTO   ")
  sqlParametroAtendimento.Add("  FROM SAM_PARAMETROSATENDIMENTO ")
  sqlParametroAtendimento.Active = True

  If sqlSituacaoEventoSolicit.FieldByName("SITUACAO").AsString = "C" Then
  		Negar = "O evento já está cancelado!"
  ElseIf sqlSituacaoEventoGerado.FieldByName("QTD").AsInteger = 0 Then
		Negar = "Não existe evento autorizado ou liberado para o evento solicitado."
  ElseIf sqlEventoPago.FieldByName("QTD").AsInteger > 0 Then
		Negar = "Existem eventos pagos para essa solicitação."
  ElseIf (sqlParametroAtendimento.FieldByName("FORNECIMENTOMEDICAMENTO").AsString <> "N" And Not ValidaNegacaoEventoFornecMEDE(pHandleEvento)) Then
		Exit Function
  ElseIf (sqlParametroAtendimento.FieldByName("FORNECIMENTOMEDICAMENTO").AsString <> "N" And Not interface.SituacaoEventoFornecimentoPermite(CurrentSystem, 0, pHandleEvento, False)) Then
		Negar = "Não é possível Negar o evento!" + Chr(13) + "Situação da Liberação de Análise Técnica de Material deste evento está 'Concluída'."
  Else
		If (interface.PericiaEventosOdonto(CurrentSystem, "S", pHandleEvento, pHandleProtocolo) = 0 And Not especifico.ATE_IgnorarPericia(CurrentSystem)) Then
			Negar = "Não é possível Negar o evento! Evento odontológico já analisado pelo perito!"
		End If
  End If


  Set interface =Nothing
  Set especifico =Nothing

  Set sqlSituacaoEventoSolicit = Nothing
  Set sqlSituacaoEventoGerado = Nothing
  Set sqlEventoPago = Nothing
  Set sqlParametroAtendimento = Nothing
  WriteBDebugMessage("TV_MOTIVONEGACAO.Negar: " + Negar)
End Function

Public Function ValidaNegacaoEventoFornecMEDE(pEventoSolicitado As Long) As Boolean
	Dim sqlSituacaoCompras As Object
	Set sqlSituacaoCompras = NewQuery

	sqlSituacaoCompras.Active = False
	sqlSituacaoCompras.Clear
	sqlSituacaoCompras.Add("SELECT F.SITUACAO, F.SITUACAOANALISETECNICA, F.SITUACAOCOMPRAS, F.HANDLE HFORNECIMENTO ")
	sqlSituacaoCompras.Add("  FROM SAM_AUTORIZ_EVENTOSOLICIT S                                                     ")
	sqlSituacaoCompras.Add("  JOIN SAM_TGE T ON T.HANDLE = S.EVENTO                                                ")
	sqlSituacaoCompras.Add("  JOIN SAM_AUTORIZ_FORNEC_EVENTO E ON E.EVENTOSOLICITADO = S.HANDLE                    ")
	sqlSituacaoCompras.Add("  JOIN SAM_AUTORIZ_FORNECIMENTO F ON F.HANDLE = E.FORNECIMENTO                         ")
	sqlSituacaoCompras.Add(" WHERE E.EVENTOSOLICITADO = :EVENTOSOLICITADO                                          ")
	sqlSituacaoCompras.Add("   AND T.TABTIPOEVENTO = 4                                                             ")
	sqlSituacaoCompras.ParamByName("EVENTOSOLICITADO").AsInteger = pEventoSolicitado
	sqlSituacaoCompras.Active = True

	If sqlSituacaoCompras.FieldByName("SITUACAOCOMPRAS").AsString = "A" Or _
	  sqlSituacaoCompras.FieldByName("SITUACAOCOMPRAS").AsString = "5" Or _
	  sqlSituacaoCompras.FieldByName("SITUACAOCOMPRAS").AsString = "6" Or _
	  sqlSituacaoCompras.FieldByName("SITUACAOCOMPRAS").AsString = "7" Or _
	  sqlSituacaoCompras.FieldByName("SITUACAOCOMPRAS").AsString = "8" Or _
	  ValidaRotinaIntegracaoCompras(sqlSituacaoCompras.FieldByName("HFORNECIMENTO").AsInteger) Then

		bsShowMessage("Evento não pode ser negado devido à compra em andamento.", "I")
		ValidaNegacaoEventoFornecMEDE = False
		Exit Function
	End If

	If sqlSituacaoCompras.FieldByName("SITUACAO").AsString = "3" Or _
		sqlSituacaoCompras.FieldByName("SITUACAO").AsString = "4" Or _
		sqlSituacaoCompras.FieldByName("SITUACAOANALISETECNICA").AsString = "3" Or _
		sqlSituacaoCompras.FieldByName("SITUACAOANALISETECNICA").AsString = "4" Then

		If bsShowMessage("Evento Possui Fornecimento de Medicamento Especial. Continuar?", "Q") = vbNo Then
			ValidaNegacaoEventoFornecMEDE = False
	  	End If
	End If

	Set sqlSituacaoCompras = Nothing
End Function

Public Function ValidaRotinaIntegracaoCompras(pHFornecimento As Long) As Boolean
	Dim sqlRotinaArquivo As Object
	Dim sqlExigeTipoAquisicao As Object

	Set sqlRotinaArquivo = NewQuery
	Set sqlExigeTipoAquisicao = NewQuery

	ValidaRotinaIntegracaoCompras = False

	sqlRotinaArquivo.Active = False
	sqlRotinaArquivo.Clear
	sqlRotinaArquivo.Add("SELECT A.HANDLE, A.SITUACAO, F.TIPOFORNECIMENTOPAI             ")
	sqlRotinaArquivo.Add("  FROM SFN_ROTINAARQUIVO A                                     ")
	sqlRotinaArquivo.Add("  JOIN SAM_AUTORIZ_FORNECIMENTO F ON F.HANDLE = A.FORNECIMENTO ")
	sqlRotinaArquivo.Add(" WHERE A.FORNECIMENTO = :HFORNEC                               ")
	sqlRotinaArquivo.ParamByName("HFORNEC").AsInteger = pHFornecimento
	sqlRotinaArquivo.Active = True

	sqlExigeTipoAquisicao.Active = False
	sqlExigeTipoAquisicao.Clear
	sqlExigeTipoAquisicao.Add("SELECT HANDLE                     ")
  	sqlExigeTipoAquisicao.Add("  FROM SAM_TIPOFORNECIMENTO       ")
  	sqlExigeTipoAquisicao.Add(" WHERE EXIGETIPOAQUISICAO = 'S'   ")
  	sqlExigeTipoAquisicao.Add("   And INATIVO = 'N'              ")
  	sqlExigeTipoAquisicao.Add("   And HANDLE = :TIPOFORNEC       ")
  	sqlExigeTipoAquisicao.ParamByName("TIPOFORNEC").AsInteger = sqlRotinaArquivo.FieldByName("TIPOFORNECIMENTOPAI").AsInteger
  	sqlExigeTipoAquisicao.Active = True

	If sqlRotinaArquivo.FieldByName("HANDLE").AsInteger > 0 And _
		sqlRotinaArquivo.FieldByName("SITUACAO").AsString <> "1" And _
		sqlRotinaArquivo.FieldByName("SITUACAO").AsString <> "8" And _
		sqlExigeTipoAquisicao.FieldByName("HANDLE").AsInteger > 0 Then

		ValidaRotinaIntegracaoCompras = True

	End If

	Set sqlRotinaArquivo = Nothing
	Set sqlExigeTipoAquisicao = Nothing
End Function
