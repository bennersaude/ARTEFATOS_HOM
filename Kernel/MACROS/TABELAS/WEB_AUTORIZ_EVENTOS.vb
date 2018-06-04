'HASH: 98BF91749BFB81162016D243E1D23E97
'#uses "*CriaTabelaTemporariaSqlServer"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterCommitted()
If WebMode Then

  If Not InTransaction Then
    StartTransaction
  End If


  If InStr(SQLServer, "SQL") > 0 Then
    Dim SQLX As Object
    Set SQLX = NewQuery

    On Error GoTo TabelasTemporarias

     SQLX.Clear
     SQLX.Add("SELECT 1 FROM #TMP_EVENTOGERADO")
     SQLX.Active = True

     Set SQLX = Nothing

     GoTo Executa

     TabelasTemporarias:
       CriaTabelaTemporariaSqlServer
  End If

  Executa:
  On Error GoTo Erro
  If WebVisionCode = "W_WEB_SADT_EVENTOS_MATMED" Then
		Dim SPP As Object
		Dim vSQL As Object


		Set SPP = NewStoredProc
		SPP.Name = "BSAUT_AUTORIZWEB"
		SPP.AutoMode = True
		SPP.AddParam("P_WEBAUTORIZ",ptInput)
		SPP.ParamByName("P_WEBAUTORIZ").DataType   = ftInteger

		SPP.AddParam("P_VERSAOTISS",ptInput)		'Int
		SPP.ParamByName("P_VERSAOTISS").DataType   = ftInteger 'Gabriel

		SPP.AddParam("P_TIPOOPERACAO",ptInput)
		SPP.ParamByName("P_TIPOOPERACAO").DataType = ftInteger

		SPP.AddParam("P_AUTORIZACAO",ptInput)
		SPP.ParamByName("P_AUTORIZACAO").DataType  = ftInteger

		SPP.AddParam("P_TIPOTISS",ptInput)
		SPP.ParamByName("P_TIPOTISS").DataType     = ftString

		SPP.AddParam("P_ORIGEM",ptInput)
		SPP.ParamByName("P_ORIGEM").DataType       = ftString

		SPP.AddParam("P_USUARIO",ptInput)
		SPP.ParamByName("P_USUARIO").DataType      = ftInteger

		SPP.AddParam("P_EHREEMBOLSO",ptInput)
		SPP.ParamByName("P_ORIGEM").DataType       = ftString

		SPP.AddParam("P_RETORNO",ptOutput)
		SPP.ParamByName("P_RETORNO").DataType      = ftString

		Set vSQL = NewQuery

		vSQL.Clear
		vSQL.Add("SELECT NUMEROAUTORIZACAO FROM WEB_AUTORIZ WHERE HANDLE = :HANDLE")
		vSQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger
		vSQL.Active = True

		SPP.ParamByName("P_WEBAUTORIZ").Value      = CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger
		SPP.ParamByName("P_TIPOOPERACAO").Value    = 110
		If vSQL.FieldByName("NUMEROAUTORIZACAO").AsInteger > 0 Then
			SPP.ParamByName("P_AUTORIZACAO").Value = vSQL.FieldByName("NUMEROAUTORIZACAO").AsInteger
		Else
			SPP.ParamByName("P_AUTORIZACAO").Value = Null
		End If
		SPP.ParamByName("P_TIPOTISS").Value        = "M"
		SPP.ParamByName("P_ORIGEM").Value          = "2"
		SPP.ParamByName("P_USUARIO").Value         = CurrentUser

	    SPP.ParamByName("P_VERSAOTISS").AsInteger = PegarHandleVersaoTISS(CLng(vSQL.FieldByName("NUMEROAUTORIZACAO").AsString))

		SPP.ExecProc

		InfoDescription = SPP.ParamByName("P_RETORNO").AsString

		Set vSQL = Nothing
		Set SPP = Nothing

  ElseIf (WebVisionCode = "W_WEB_AUTORIZ_EVENTOS_AUTORIZ") Or (WebVisionCode = "V_WEB_AUTORIZ_EVENTOS_ODONTO") Or (WebVisionCode = "V_WEB_AUTORIZ_EVENTOS") Or (WebVisionCode = "W_WEB_AUTORIZ_EVENTOS_ODONTO") Then
    Dim sql As Object
    Set sql = NewQuery

    sql.Active = False
	sql.Clear
	sql.Add("SELECT NUMEROAUTORIZACAO FROM WEB_AUTORIZ WHERE HANDLE = :HANDLE ")
	sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger
	sql.Active = True

	If sql.FieldByName("NUMEROAUTORIZACAO").AsInteger > 0 Then

      Dim SP2 As Object
	  Set SP2 = NewStoredProc
	  SP2.Name = "BSAUT_AUTORIZINSEREEVENTOSWEB"
	  SP2.AddParam("P_WEBAUTORIZ", ptInput)
	  SP2.ParamByName("P_WEBAUTORIZ").DataType = ftInteger
	  SP2.AddParam("P_WEBAUTORIZEVENTO", ptInput)
	  SP2.ParamByName("P_WEBAUTORIZEVENTO").DataType = ftInteger
	  SP2.AddParam("P_HANDLEAUTORIZ", ptInput)
	  SP2.ParamByName("P_HANDLEAUTORIZ").DataType = ftInteger
	  SP2.AddParam("P_USUARIO", ptInput)
	  SP2.ParamByName("P_USUARIO").DataType = ftInteger
	  SP2.AddParam("P_TIPOOPERACAOTISS", ptInput)
	  SP2.ParamByName("P_TIPOOPERACAOTISS").DataType = ftString

	  SP2.ParamByName("P_WEBAUTORIZ").AsInteger = CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger
	  SP2.ParamByName("P_WEBAUTORIZEVENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  SP2.ParamByName("P_HANDLEAUTORIZ").AsInteger = sql.FieldByName("NUMEROAUTORIZACAO").AsInteger
	  SP2.ParamByName("P_USUARIO").AsInteger = CurrentUser
	  SP2.ParamByName("P_TIPOOPERACAOTISS").AsString = "S"
	  SP2.ExecProc

      If WebMode Then
		Dim spBSAUT_GERARPROTOCOLOTRANSACAO As BStoredProc
		Set spBSAUT_GERARPROTOCOLOTRANSACAO = NewStoredProc
		spBSAUT_GERARPROTOCOLOTRANSACAO.AutoMode = True
		spBSAUT_GERARPROTOCOLOTRANSACAO.Name = "BSAUT_GERARPROTOCOLOTRANSACAO"
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_ORIGEMPROCESSO",ptInput, ftString)
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_AUTORIZACAO",ptInput, ftInteger)
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_USUARIO",ptInput, ftInteger)
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_NUMEROPROTOCOLOTRANSACAO",ptInputOutput, ftInteger)
		spBSAUT_GERARPROTOCOLOTRANSACAO.AddParam("P_HANDLEATENDIMENTOCENTRAL",ptInputOutput, ftInteger)
		spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_ORIGEMPROCESSO").AsString = "D" ' simulando como desktop, mesmo sendo no ambiente web, para localizar o tipo dele (complemento, prorrogação, e outros)
		spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_AUTORIZACAO").AsInteger = sql.FieldByName("NUMEROAUTORIZACAO").AsInteger
		spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_USUARIO").AsInteger = CurrentUser

		If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") <= 0 Then
		  spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsInteger = 0 ' Não veio do protocolo de transação portanto criar um protocolo para ele
		Else
		  spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_NUMEROPROTOCOLOTRANSACAO").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
		End If

        spBSAUT_GERARPROTOCOLOTRANSACAO.ParamByName("P_HANDLEATENDIMENTOCENTRAL").AsInteger = 0

		spBSAUT_GERARPROTOCOLOTRANSACAO.ExecProc
		Set spBSAUT_GERARPROTOCOLOTRANSACAO = Nothing
      End If

      Set SP2 = Nothing
	  End If
  End If

    If InTransaction Then
      Commit
    End If
End If
Exit Sub

Erro:
  InfoDescription = Err.Description
  CancelDescription = Err.Description
  If InTransaction Then
    Rollback
  End If
End Sub

Public Sub TABLE_AfterInsert()
	If CurrentQuery.FieldByName("WEBAUTORIZ").IsNull Then
		assumirWebAutoriz
	End If
End Sub

Public Sub TABLE_AfterScroll()
    If WebMode Then
        Dim vSQL As Object
        Dim numAutoriz As Long
        Set vSQL = NewQuery

		vSQL.Clear
		vSQL.Add("SELECT NUMEROAUTORIZACAO,")
		vSQL.Add("       CBOSSOLICITANTE,")
		vSQL.Add("       CARATERATENDIMENTO,")
		vSQL.Add("       REGIMEINTERNACAO,")
		vSQL.Add("       TIPOINTERNACAO,")
		vSQL.Add("       TIPOATENDIMENTO")
		vSQL.Add("FROM WEB_AUTORIZ")
		vSQL.Add("WHERE HANDLE = :HANDLE")
		vSQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger
		vSQL.Active = True

		If vSQL.FieldByName("NUMEROAUTORIZACAO").IsNull Or vSQL.FieldByName("NUMEROAUTORIZACAO").AsString = "" Then
		   numAutoriz = CLng("0")
		Else
		   numAutoriz =  CLng(vSQL.FieldByName("NUMEROAUTORIZACAO").AsString)
		End If

		EVENTO.WebLocalWhere = "( A.HANDLE IN (SELECT EVENTO FROM SAM_TGE_TABELATISS WHERE TABELATISS = @~CAMPO(CODIGOTABELA)) OR (@~CAMPO(CODIGOTABELA) = -1) "+ _
		                       " AND A.ULTIMONIVEL = 'S') AND A.INATIVO = 'N'AND A.INCLUINOTIPOANEXO = 'N' "

		If Not vSQL.FieldByName("CARATERATENDIMENTO").IsNull Then
			EVENTO.WebLocalWhere = EVENTO.WebLocalWhere + " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
			                                              "                  FROM SAM_CLASSEEVENTO_CARATERATEND CCA" + Chr(13) + _
			                                              "                  WHERE CCA.CLASSEEVENTO = A.CLASSEEVENTO) OR " + Chr(13) + _
			                                              "      EXISTS (SELECT 1" + Chr(13) + _
			                                              "              FROM SAM_CLASSEEVENTO_CARATERATEND CCA" + Chr(13) + _
			                                              "              JOIN TIS_CARATERATENDIMENTO CAR ON CAR.HANDLE = CCA.CARATERATENDIMENTO" + Chr(13) + _
			                                              "              WHERE CCA.CLASSEEVENTO = A.CLASSEEVENTO" + Chr(13) + _
			                                              "                AND CAR.CODIGO = " + vSQL.FieldByName("CARATERATENDIMENTO").AsString + Chr(13) + _
			                                              "                AND CAR.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')))"
		End If

		If Not vSQL.FieldByName("REGIMEINTERNACAO").IsNull Then
			EVENTO.WebLocalWhere = EVENTO.WebLocalWhere + Chr(13) + " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
			                                                        "                  FROM SAM_CLASSEEVENTO_REGIMEINTER CRI" + Chr(13) + _
			                                                        "                  WHERE CRI.CLASSEEVENTO = A.CLASSEEVENTO) OR" + Chr(13) + _
			                                                        "      EXISTS (SELECT 1" + Chr(13) + _
			                                                        "              FROM SAM_CLASSEEVENTO_REGIMEINTER CRI" + Chr(13) + _
			                                                        "              WHERE CRI.CLASSEEVENTO = A.CLASSEEVENTO" + Chr(13) + _
			                                                        "                AND CRI.REGIMEINTERNACAO = " + vSQL.FieldByName("REGIMEINTERNACAO").AsString +"))"
		End If

		If Not vSQL.FieldByName("TIPOINTERNACAO").IsNull Then
			EVENTO.WebLocalWhere = EVENTO.WebLocalWhere + Chr(13) + " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
			                                                        "                  FROM SAM_CLASSEEVENTO_TIPOINTER CTI" + Chr(13) + _
			                                                        "                  WHERE CTI.CLASSEEVENTO = A.CLASSEEVENTO) OR " + Chr(13) + _
			                                                        "      EXISTS (SELECT 1" + Chr(13) + _
			                                                        "              FROM SAM_CLASSEEVENTO_TIPOINTER CTI" + Chr(13) + _
			                                                        "              WHERE CTI.CLASSEEVENTO = A.CLASSEEVENTO" + Chr(13) + _
			                                                        "                AND CTI.TIPOINTERNACAO = " + vSQL.FieldByName("TIPOINTERNACAO").AsString +"))"
		End If

		If Not vSQL.FieldByName("TIPOATENDIMENTO").IsNull Then
			EVENTO.WebLocalWhere = EVENTO.WebLocalWhere + Chr(13) + " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
			                                                        "                  FROM SAM_CLASSEEVENTO_TIPOATEND CTA" + Chr(13) + _
			                                                        "                  WHERE CTA.CLASSEEVENTO = A.CLASSEEVENTO) OR " + Chr(13) + _
			                                                        "      EXISTS (SELECT 1" + Chr(13) + _
			                                                        "              FROM SAM_CLASSEEVENTO_TIPOATEND CTA" + Chr(13) + _
			                                                        "              WHERE CTA.CLASSEEVENTO = A.CLASSEEVENTO" + Chr(13) + _
			                                                        "                AND CTA.TIPOATENDIMENTO = " + vSQL.FieldByName("TIPOATENDIMENTO").AsString +"))"
		End If

		CODIGOTABELA.WebLocalWhere = " A.VERSAOTISS = "+ CStr(PegarHandleVersaoTISS(numAutoriz))
		GRAUPARTICIPACAO.WebLocalWhere = " A.VERSAOTISS = "+ CStr(PegarHandleVersaoTISS(numAutoriz))

		Set vSQL = Nothing
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If WebMode Then

	  If CurrentQuery.FieldByName("CODIGOTABELA").AsInteger > 0 Then
	  	Dim qTabTiss As Object
	  	Set qTabTiss = NewQuery
	  	qTabTiss.Add("SELECT * FROM SAM_TGE_TABELATISS WHERE EVENTO = :EVENTO AND TABELATISS = :TABELATISS")
	  	qTabTiss.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
		qTabTiss.ParamByName("TABELATISS").AsInteger = CurrentQuery.FieldByName("CODIGOTABELA").AsInteger
		qTabTiss.Active = True

		If qTabTiss.EOF Then
			bsShowMessage("Evento incompatível com o código tabela selecionado.", "E")
			CanContinue = False
			Set qTabTiss = Nothing
			Exit Sub
		End If

		Set qTabTiss = Nothing

	  End If


		If CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger = 0 Then
			CanContinue=False
			CancelDescription = "Autorização não foi incluída pela WEB, só será permitido a inclusão de eventos pelo sistema Desktop"
			Exit Sub
		End If
		If WebVisionCode = "W_WEB_SADT_EVENTOS_MATMED" Then
          CurrentQuery.FieldByName("INSERIUNAGUIA").AsString = "N" ' AINDA NÃO INSERIU
		End If

        'Validação da autorização dentro dos prazos permitidos em parametros
        'Lembrando que no Sistema Desktop essa validação é feita na própria DLL - CA043
        Dim vDataInferior As Date
        Dim vDataPosterior As Date
		Dim qParametrosAtendimento As Object
		Set qParametrosAtendimento = NewQuery

		qParametrosAtendimento.Clear
		qParametrosAtendimento.Add("SELECT QTIDADEDIASANTERIORES,   ")
		qParametrosAtendimento.Add("       QTIDADEDIASPOSTERIORES   ")
		qParametrosAtendimento.Add("  FROM SAM_PARAMETROSATENDIMENTO")
		qParametrosAtendimento.Active = True

		vDataInferior = DateAdd("d",-qParametrosAtendimento.FieldByName("QTIDADEDIASANTERIORES").Value,ServerDate)
		vDataPosterior = DateAdd("d",qParametrosAtendimento.FieldByName("QTIDADEDIASPOSTERIORES").Value,ServerDate)

		If (vDataInferior > CurrentQuery.FieldByName("DATA").AsDateTime) Then
		    bsShowMessage("A data de atendimento não deve ser inferior a " + qParametrosAtendimento.FieldByName("QTIDADEDIASANTERIORES").AsInteger + " dias em relação a data atual.","E")
		    CanContinue = False
		End If

		If (vDataPosterior < CurrentQuery.FieldByName("DATA").AsDateTime) Then
		    bsShowMessage("A data de atendimento não deve ser superior a " + qParametrosAtendimento.FieldByName("QTIDADEDIASPOSTERIORES").AsInteger + " dias em relação a data atual.","E")
		    CanContinue = False
		End If

		Set qParametrosAtendimento = Nothing


		If WebVisionCode = "V_WEB_AUTORIZ_EVENTOS_ODONTO" Or WebVisionCode = "W_WEB_AUTORIZ_EVENTOS_ODONTO" Then 'SMS 90455 - Ricardo Rocha
			Dim qOdonto As Object
			Set qOdonto = NewQuery

			qOdonto.Clear
			qOdonto.Add("SELECT HANDLE")
			qOdonto.Add("  FROM TIS_DENTEFACE")
			qOdonto.Add(" WHERE FACEOCLUSAL = :OCLUSAL and FACELINGUAL = :LINGUAL and FACEMESIAL = :MESIAL")
			qOdonto.Add("   and FACEVESTIBULAR = :VESTIBULAR and FACEDISTAL = :DISTAL and FACEINCISAL = :INCISAL")
			qOdonto.Add("   and PALATINA = :PALATINA")

            If CurrentQuery.FieldByName("DENTE").AsInteger > 0 Then
			  qOdonto.Add("   and DENTE2 = :DENTE2")
			  qOdonto.ParamByName("DENTE2").AsInteger = CurrentQuery.FieldByName("DENTE").AsInteger
			Else
			  qOdonto.Add("   and DENTE2 IS NULL")
			End If
			If CurrentQuery.FieldByName("REGIAODENTARIA").AsInteger > 0 Then
			  qOdonto.Add("   And REGIAO = :REGIAO")
			  qOdonto.ParamByName("REGIAO").AsInteger = CurrentQuery.FieldByName("REGIAODENTARIA").AsInteger
			Else
			  qOdonto.Add("   And REGIAO IS NULL ")
			End If

			qOdonto.ParamByName("OCLUSAL").AsString = CurrentQuery.FieldByName("OCLUSAL").AsString
			qOdonto.ParamByName("LINGUAL").AsString = CurrentQuery.FieldByName("LINGUAL").AsString
			qOdonto.ParamByName("MESIAL").AsString = CurrentQuery.FieldByName("MESIAL").AsString
			qOdonto.ParamByName("VESTIBULAR").AsString = CurrentQuery.FieldByName("VESTIBULAR").AsString
			qOdonto.ParamByName("DISTAL").AsString = CurrentQuery.FieldByName("DISTAL").AsString
			qOdonto.ParamByName("INCISAL").AsString = CurrentQuery.FieldByName("INCISAL").AsString
			qOdonto.ParamByName("PALATINA").AsString = CurrentQuery.FieldByName("PALATINA").AsString
			qOdonto.Active = True

			If qOdonto.FieldByName("HANDLE").IsNull Then
				bsShowMessage("Para as informações dos campos Dente, Região e Face não existe grau informado nos parâmetros do TISS", "I")
			End If

			If Not (CurrentQuery.FieldByName("CODIGOTABELA").AsInteger > 0) Then
			  Select Case WebVisionCode
		  		Case "W_WEB_AUTORIZ_EVENTOS_AUTORIZ", "W_WEB_AUTORIZ_EVENTOS_ODONTO"

	        	  Dim vObrigarCamposTissWeb As Boolean
		          Dim vDllEspec As Object
	 	   		  Set vDllEspec = CreateBennerObject("Especifico.UEspecifico")
				  vObrigarCamposTissWeb = vDllEspec.AUT_ExigeCamposTISSWeb(CurrentSystem)
				  Set vDllEspec = Nothing

				If vObrigarCamposTissWeb Then
	           	  bsShowMessage("Campo Código Tabela é obrigatório", "E")
	           	  CanContinue = False
				End If

	    	  End Select
	  		End If
		End If

		Dim qAux As Object
		Set qAux = NewQuery
		qAux.Clear
        qAux.Add("SELECT HANDLE, INCLUINOTIPOANEXO FROM SAM_TGE WHERE HANDLE = :HANDLE AND INCLUINOTIPOANEXO <> 'N'")
        qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
        qAux.Active = True
        If qAux.FieldByName("HANDLE").AsInteger > 0 Then
          bsShowMessage("Procedimento solicitado está parametrizado como " + GetTipoAnexo(qAux.FieldByName("INCLUINOTIPOANEXO").AsString) + " portanto sua inclusão deve ser realizado através da solicitação de anexo","E")
          CanContinue = False
          Set qAux = Nothing
          Exit Sub
        End If

        Set qAux = Nothing

        Dim vSQL As Object
		Set vSQL = NewQuery

		vSQL.Clear
		vSQL.Add("SELECT NUMEROAUTORIZACAO FROM WEB_AUTORIZ WHERE HANDLE = :HANDLE")
		vSQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger
		vSQL.Active = True

		If Not vSQL.FieldByName("NUMEROAUTORIZACAO").IsNull Then
		    bsShowMessage("Solicitação já validada! Não é possível inserir novos procedimentos.", "E")
		    CanContinue = False
		    Set vSQL = Nothing
		    Exit Sub
		End If

		Set vSQL = Nothing
	End If
End Sub

Public Function GetTipoAnexo(pIncluiNoTipoAnexo As String)
  If pIncluiNoTipoAnexo = "O" Then
    GetTipoAnexo = "OPME"
  ElseIf pIncluiNoTipoAnexo = "R" Then
    GetTipoAnexo = "Radioterapia"
  ElseIf pIncluiNoTipoAnexo = "Q" Then
    GetTipoAnexo = "Quimioterapia"
  End If
End Function

Public Sub assumirWebAutoriz
	Dim sql As BPesquisa
	Set sql=NewQuery

	If RecordHandleOfTable("SAM_AUTORIZ") > 0 Then
	  sql.Clear
  	  sql.Add("SELECT HANDLE FROM WEB_AUTORIZ WHERE NUMEROAUTORIZACAO = :AUTORIZ")
	  sql.ParamByName("AUTORIZ").AsInteger = RecordHandleOfTable("SAM_AUTORIZ")
	  sql.Active=True
	Else
	  sql.Clear
      sql.Add("SELECT HANDLE                                                                                           ")
      sql.Add("  FROM WEB_AUTORIZ                                                                                      ")
      sql.Add(" WHERE PROTOCOLOTRANSACAO IN (SELECT MIN(P1.HANDLE)                                                     ")
      sql.Add("                                FROM SAM_PROTOCOLOTRANSACAOAUTORIZ P1                                   ")
      sql.Add("                              	JOIN SAM_PROTOCOLOTRANSACAOAUTORIZ P2 ON P1.AUTORIZACAO = P2.AUTORIZACAO ")
      sql.Add("                               WHERE P2.HANDLE = :PROTOCOLOTRANSACAO                                    ")
      sql.Add("                              )                                                                         ")
	  sql.ParamByName("PROTOCOLOTRANSACAO").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
	  sql.Active=True
	End If

	CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger=sql.FieldByName("HANDLE").AsInteger

	Set sql=Nothing
End Sub

Public Function PegarHandleVersaoTISS(Optional piHandleAutorizacao As Long = 0) As Integer
    Dim viVersaoTISS As Integer
    Dim achouVersaoTiss As Boolean
    Dim sql As Object
    Set sql = NewQuery
 	achouVersaoTiss =False

	If piHandleAutorizacao <> 0 Then
      sql.Active = False
	  sql.Clear
      sql.Add("SELECT A.VERSAOTISS VERSAOTISS")
      sql.Add("FROM SAM_AUTORIZ A ")
      sql.Add("WHERE A.HANDLE = :HANDLE")
      sql.ParamByName("HANDLE").Value = piHandleAutorizacao
      sql.Active = True
      achouVersaoTiss = Not sql.FieldByName("VERSAOTISS").IsNull
	End If

	If Not achouVersaoTiss Then
        sql.Active = False
	    sql.Clear
        sql.Add("SELECT MAX(A.HANDLE) VERSAOTISS ")
        sql.Add("FROM TIS_VERSAO A ")
        sql.Add("WHERE A.ATIVODESKTOP = 'S' ")
        sql.Active = True
    End If

    viVersaoTISS = sql.FieldByName("VERSAOTISS").AsInteger
	Set sql = Nothing

	PegarHandleVersaoTISS = viVersaoTISS
End Function
