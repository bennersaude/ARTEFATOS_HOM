'HASH: 4FF9A74DF3A7900814D077FFA18A465F
'#uses "*CriaTabelaTemporariaSqlServer"

Option Explicit

'--------------------------------------------------------------------------------------------------------------------------
'  SOMENTE USAR A PARTIR DA SAM_AUTORIZ_EVENTOSOLICIT E SAM_PROTOCOLOTRANSACAOAUTORIZ--------------------------------------
'--------------------------------------------------------------------------------------------------------------------------

Public Sub reverter(pUsuario As Long)
	WriteBDebugMessage("TV_FORM0022.reverter - pUsuario [" + CStr(pUsuario) + "]")
	CriaTabelaTemporariaSqlServer
	Dim mensagem As String
	Dim vHandleProtocolo As Long
	Dim vHandleAutorizacao As Long
	Dim vHandleEventoSolicit As Long
	Dim dll As Object

	Set dll=CreateBennerObject("samauto.autorizador")
	dll.inicializar(CurrentSystem, "A")

	vHandleProtocolo   = CLng(SessionVar("PROTOCTRANSAUTOR"))
	vHandleAutorizacao = CLng(SessionVar("HANDLEAUTORIZACAO"))

	If  (SessionVar("HANDLEEVENTOSOLICIT") <> "0" Or SessionVar("HANDLEEVENTOSOLICIT") <> Null Or SessionVar("HANDLEEVENTOSOLICIT") <> "")  Then
	   vHandleEventoSolicit = CLng(SessionVar("HANDLEEVENTOSOLICIT"))
	Else
	   vHandleEventoSolicit = RecordHandleOfTable("SAM_AUTORIZ_EVENTOSOLICIT")
	End If

	If vHandleProtocolo > 0 Then

	   mensagem = dll.ReverterEventosDoProtocoloComMotivo( _
		          CurrentSystem, _
				  vHandleAutorizacao, _
		          vHandleProtocolo, _
		          vHandleEventoSolicit, _
				  pUsuario, _
				  CurrentQuery.FieldByName("MOTIVOREVERSAO").AsInteger)
	Else

	   mensagem = dll.reverterSAM_AUTORIZ_EVENTOGERADO_comMotivo( _
				  CurrentSystem, _
				  vHandleEventoSolicit, _
				  pUsuario, _
				  CurrentQuery.FieldByName("MOTIVOREVERSAO").AsInteger, _
				  "S", _
				   0)

    End If

	dll.finalizar
	Set dll=Nothing

	If mensagem<>"" Then
		InfoDescription = mensagem
	Else
		InfoDescription = "Reversão concluída com sucesso"
	End If
	WriteBDebugMessage("TV_FORM0022.reverter - InfoDescription [" + InfoDescription + "]")
End Sub


Public Sub TABLE_AfterPost()
	WriteBDebugMessage("TV_FORM0022.TABLE_AfterPost - Início")
	Dim vUsuario As Long
	vUsuario = verificaUsuario
	If vUsuario > 0 Then
		reverter(vUsuario)
	Else
		InfoDescription = "Usuário e senha não conferem"
	End If
	WriteBDebugMessage("TV_FORM0022.TABLE_AfterPost - InfoDescription [" + InfoDescription + "]")
End Sub

Public Function verificaUsuario As Long
	' se não digitar o usuário/SENHA, assume o USUARIO corrente
	If (CurrentQuery.FieldByName("USUARIO").AsString <> "") Then
		Dim dll As Object
		Set dll = CreateBennerObject("samauto.rotinas")
		dll.inicializar(CurrentSystem)
		verificaUsuario = dll.VerificaUsuario(CurrentSystem, CurrentQuery.FieldByName("USUARIO").AsString, CurrentQuery.FieldByName("SENHA").AsString)
		dll.Finalizar
		Set dll = Nothing
	Else
		verificaUsuario = CurrentUser
	End If
End Function


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	WriteBDebugMessage("TV_FORM0022.TABLE_BeforeInsert - Início")
	Dim dll As Object
  	Set dll=CreateBennerObject("samauto.autorizador")

  	Dim vHandleProtocolo As Long
  	vHandleProtocolo   = CLng(SessionVar("PROTOCTRANSAUTOR"))

	If vHandleProtocolo > 0 Then

	   Dim qReverteProtocoloTransacao As BPesquisa
	   Set qReverteProtocoloTransacao = NewQuery

	   qReverteProtocoloTransacao.Add("SELECT D.*")
  	   qReverteProtocoloTransacao.Add("FROM SAM_AUTORIZ_EVENTOGERADO D")
	   qReverteProtocoloTransacao.Add("WHERE D.PROTOCOLOTRANSACAO = :PROTOCOLOTRANSACAO")
  	   qReverteProtocoloTransacao.Add("And D.SITUACAO    = 'N' ")
	   qReverteProtocoloTransacao.ParamByName("PROTOCOLOTRANSACAO").Value = vHandleProtocolo
	   qReverteProtocoloTransacao.Active = True
	   qReverteProtocoloTransacao.First

	   If qReverteProtocoloTransacao.EOF Then
          CanContinue=False
  	      CancelDescription = "Não existem eventos negados para o protocolo ou há evento cancelado."
	   Else

	      While Not qReverteProtocoloTransacao.EOF

             If Not dll.verificaNecessidadeReverter(CurrentSystem, qReverteProtocoloTransacao.FieldByName("EVENTOSOLICITADO").AsInteger, "S") Then
                CanContinue=False
  	            CancelDescription = "Não existem eventos negados para o protocolo ou há evento cancelado."
             End If

             qReverteProtocoloTransacao.Next
	      Wend

	   End If

	   Set qReverteProtocoloTransacao=Nothing

	Else

       If Not dll.verificaNecessidadeReverter(CurrentSystem, RecordHandleOfTable("SAM_AUTORIZ_EVENTOSOLICIT"), "S") Then
  	      CanContinue=False
  	      CancelDescription = "Não existem eventos negados para o evento solicitado ou o evento solicitado está cancelado."
       End If

	End If

	Set dll=Nothing
	WriteBDebugMessage("TV_FORM0022.TABLE_BeforeInsert - Fim")
End Sub
