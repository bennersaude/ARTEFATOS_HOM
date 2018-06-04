'HASH: D5D9A991204B0A1C5A378CA546224B30
'#uses "*bsShowMessage"
'#uses "*Modulo11"
'#uses "*CriaTabelaTemporariaSqlServer"
Option Explicit

Public Sub ChamarRelatorioAutorizacao
  Dim viHandleRelatorio As Integer
  Dim qRelatorio        As Object

  'Alterado para buscar o relatório configurado no tipo de autorização - SMS - 107122
  Dim qRel As Object
  Set qRel = NewQuery

  qRel.Add("SELECT TA.RELATORIOAUTORIZACAO                           ")
  qRel.Add("  FROM SAM_TIPOAUTORIZ TA                                ")
  qRel.Add("  JOIN SAM_AUTORIZ A  ON (TA.HANDLE = A.TIPOAUTORIZACAO) ")
  qRel.Add("WHERE A.HANDLE = :HANDLE")

  qRel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger

  qRel.Active = True

  If qRel.FieldByName("RELATORIOAUTORIZACAO").IsNull Then
    bsShowMessage("Não encontrado o relatório para guias de consulta com o tipo de autorização selecionado", "I")
    Set qRel = Nothing
  Else
    viHandleRelatorio = qRel.FieldByName("RELATORIOAUTORIZACAO").AsInteger
    SessionVar("WEBHandleFiltro") = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsString

    Set qRel = Nothing

    If InTransaction Then
		Commit
	End If

    ReportPreview(viHandleRelatorio, "", False, True)
  End If
End Sub


Public Sub TABLE_AfterInsert()
  If WebVisionCode = "V_WEB_ODONTO_AUTORIZ" Or WebVisionCode = "V_WEB_ODONTO_AUT_REEMBOLSO" Or _
     WebVisionCode = "W_WEB_ODONTO_AUTORIZ" Or WebVisionCode = "K_WEB_ODONTO_AUTORIZ" Then
	 CurrentQuery.FieldByName("TIPOATENDIMENTOODONTOLOGICO").AsString = "1"
  Else
    CurrentQuery.FieldByName("TIPOATENDIMENTOODONTOLOGICO").AsString = "0"
  End If

  If Not CurrentQuery.FieldByName("NRSOLICITACAOPRINC").IsNull Then
	Dim bs As CSBusinessComponent
	Dim viAutorizacao As Long
	Dim vsTipoSolicitacao As String

	If WebVisionCode = "V_WEB_SPSADT_AUTORIZ" Or _
	   WebVisionCode = "W_WEB_SADTAUTORIZ" Then
		vsTipoSolicitacao = "1"
	ElseIf WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ" Or _
	   	   WebVisionCode = "W_WEB_INTERNACAO_AUTORIZ" Then
		vsTipoSolicitacao = "2"
  	ElseIf WebVisionCode = "V_WEB_ODONTO_AUTORIZ" Or _
	   	   WebVisionCode = "W_WEB_ODONTO_AUTORIZ" Then
    	vsTipoSolicitacao = "3"
    Else
		vsTipoSolicitacao = "9"
  	End If

	Set bs = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamAutorizBLL, Benner.Saude.Atendimento.Business") ' formato: [namespace.classe], [assembly]
	bs.AddParameter(pdtString, CurrentQuery.FieldByName("NRSOLICITACAOPRINC").AsString) 'numeroGuiaPrincipal
	bs.AddParameter(pdtString, vsTipoSolicitacao) 'tipoSolicitacao
	viAutorizacao = CLng(bs.Execute("LocalizarAutorizacaoPrincipalValida"))
	If viAutorizacao > 0 Then
		Dim sqlAutorizacao As BPesquisa
		Set sqlAutorizacao = NewQuery

		sqlAutorizacao.Add("SELECT BENEFICIARIO")
		sqlAutorizacao.Add("FROM SAM_AUTORIZ")
		sqlAutorizacao.Add("WHERE HANDLE = :HAUTORIZACAO")
		sqlAutorizacao.ParamByName("HAUTORIZACAO").AsInteger = viAutorizacao
		sqlAutorizacao.Active = True

		CurrentQuery.FieldByName("AUTORIZACAOPRINCIPAL").AsInteger = viAutorizacao
		CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = sqlAutorizacao.FieldByName("BENEFICIARIO").AsInteger
		NRSOLICITACAOPRINC.ReadOnly = True
		BENEFICIARIO.ReadOnly = True

		Set sqlAutorizacao = Nothing
	End If
  End If

End Sub

Public Sub TABLE_AfterPost()
  WriteBDebugMessage("WEB_AUTORIZ.TABLE_AfterPost - Início")
  Dim viNumeroAutorizacao As Long
  Dim Obj As Object
  Dim UsuarioAutoriz As BPesquisa

  Set UsuarioAutoriz = NewQuery
  UsuarioAutoriz.Clear
  UsuarioAutoriz.Add("UPDATE WEB_AUTORIZ")
  UsuarioAutoriz.Add(" SET USUARIO = :USUARIO")
  UsuarioAutoriz.Add("WHERE HANDLE = :HANDLE")
  UsuarioAutoriz.ParamByName("USUARIO").Value = CurrentUser
  UsuarioAutoriz.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  UsuarioAutoriz.ExecSQL

  Set UsuarioAutoriz = Nothing

	If WebMode Then
		'W_WEB_SADTPAGAMENTO
		If (WebVisionCode = "W_WEB_SADTAUTORIZPAGAMAENTO")  Or (WebVisionCode = "W_WEB_SADTELEGIBILIDADE") Or (WebVisionCode = "W_WEB_SADTAUTORIZ") Or _
			   (WebVisionCode = "V_WEB_SPSADT_AUTORIZ") Or (WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ") Or (WebVisionCode = "V_WEB_ODONTO_AUTORIZ") Or _
			   (WebVisionCode = "W_WEB_ODONTO_AUTORIZ") Or (WebVisionCode = "V_WEB_INTERNACAO_AUT_REEMBOLSO") Or (WebVisionCode = "V_WEB_SPSADT_AUT_REEMBOLSO") Or _
			   (WebVisionCode = "V_WEB_ODONTO_AUT_REEMBOLSO") Then

			Set Obj = CreateBennerObject("Especifico.UEspecifico")
			bsShowMessage(Obj.ATE_MensagemAutorizadorExterno(CurrentSystem, WebVisionCode), "I")


			Set Obj = Nothing
			Exit Sub

		End If


		' --- Inícion da transação ---------------------------------
		' ----------------------------------------------------------
		If Not InTransaction Then
			StartTransaction
		End If

		If InStr(SQLServer, "SQL") > 0 Then
			Dim SQLx As Object
			Set SQLx = NewQuery

			On Error GoTo TabelasTemporarias

			SQLx.Clear
			SQLx.Add("SELECT 1 FROM #TMP_LIMITE")
			SQLx.Active = True

			Set SQLx = Nothing

			GoTo Procedure

			TabelasTemporarias:
			CriaTabelaTemporariaSqlServer

			Set SQLx = Nothing
		End If

		Procedure:
		On Error GoTo Erro

		Dim SPP As BStoredProc
		Set SPP = NewStoredProc
		SPP.AutoMode = True
		SPP.Name = "BSAUT_AUTORIZWEB"
		SPP.AddParam("P_VERSAOTISS",ptInput, ftInteger)		   'Int ' SMS 104421 - TISS 2.2.1 - Danilo Raisi
		SPP.AddParam("P_WEBAUTORIZ",ptInput, ftInteger)        'Int
		SPP.AddParam("P_TIPOOPERACAO",ptInput, ftInteger)      'Int
		SPP.AddParam("P_AUTORIZACAO",ptInput, ftInteger)       'Int
		SPP.AddParam("P_TIPOTISS",ptInput, ftString)          'Varchar(1)
		SPP.AddParam("P_ORIGEM",ptInput, ftString)            'Varchar(1)
		SPP.AddParam("P_USUARIO",ptInput, ftInteger)           'Int
		SPP.AddParam("P_NUMEROAUTORIZACAO", ptInput, ftFloat)  ' Paulo Melo
		SPP.AddParam("P_EHREEMBOLSO",ptInput, ftString)
		SPP.AddParam("P_RETORNO",ptOutput, ftString)          'Varchar(100)

		Select Case WebVisionCode
			' somente autorização
			Case "V_WEB_CONSULTA_AUTORIZ"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "C"
				SPP.ParamByName("P_ORIGEM").AsString        = "1"

			Case "V_WEB_CONSULTA_AUT_REEMBOLSO"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "C"
				SPP.ParamByName("P_ORIGEM").AsString        = "1"
				SPP.ParamByName("P_EHREEMBOLSO").AsString   = "S"

			Case "V_WEB_SPSADT_AUTORIZ"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "S"
				SPP.ParamByName("P_ORIGEM").AsString        = "1"

			Case "V_WEB_SPSADT_AUT_REEMBOLSO"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "S"
				SPP.ParamByName("P_ORIGEM").AsString        = "1"
				SPP.ParamByName("P_EHREEMBOLSO").AsString   = "S"

			Case "V_WEB_INTERNACAO_AUTORIZ"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 120
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "I"
				SPP.ParamByName("P_ORIGEM").AsString        = "1"

            Case "V_WEB_INTERNACAO_AUT_REEMBOLSO"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "I"
				SPP.ParamByName("P_ORIGEM").AsString        = "1"
				SPP.ParamByName("P_EHREEMBOLSO").AsString   = "S"

			Case "V_WEB_ODONTO_AUTORIZ" 'SMS 90455 - Ricardo Rocha - 04/06/2008
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "O"
				SPP.ParamByName("P_ORIGEM").AsString		= "1"

			Case "V_WEB_ODONTO_AUTORIZ" 'SMS 90455 - Ricardo Rocha - 04/06/2008
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "O"
				SPP.ParamByName("P_ORIGEM").AsString		= "2"

			Case "V_WEB_ODONTO_AUT_REEMBOLSO"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "O"
				SPP.ParamByName("P_ORIGEM").AsString		= "1"
				SPP.ParamByName("P_EHREEMBOLSO").AsString	= "S"

			'Autorizador Externo - Gabriel
			Case "W_WEB_CONSULTA_AUTORIZ"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 120
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "C"
				SPP.ParamByName("P_ORIGEM").AsString		= "2"
			'Autorizador Externo - Fim.


			'outros processos
			Case "W_WEB_CONSULTA_EXECUCAO"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 110
				SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
				SPP.ParamByName("P_TIPOTISS").AsString      = "C"
				SPP.ParamByName("P_ORIGEM").AsString        = "2"
			Case "W_WEB_CONSULTA_GUIA"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 110
				SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
				SPP.ParamByName("P_TIPOTISS").AsString      = "C"
				SPP.ParamByName("P_ORIGEM").AsString        = "2"
			Case "W_WEB_CONSULTA_ELEGIBILIDADE"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 160
				SPP.ParamByName("P_AUTORIZACAO").Value      = Null
				SPP.ParamByName("P_TIPOTISS").AsString      = "C"
				SPP.ParamByName("P_ORIGEM").AsString        = "2"
			Case "W_WEB_SADTPAGAMENTO"
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 110
				SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
				SPP.ParamByName("P_TIPOTISS").AsString      = "S"
				SPP.ParamByName("P_ORIGEM").AsString        = "2"

			Case Else
			    SPP.ParamByName("P_ORIGEM").AsString        = "2"
		End Select

		' SMS 104421 - TISS 2.2.1 - Danilo Raisi

		Dim sql As BPesquisa
		Set sql = NewQuery

		sql.Add("SELECT MAX(HANDLE) HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'")
		sql.Active = True

		NewCounter2("SAM_AUTORIZ", 0, 1, viNumeroAutorizacao)

		Dim vsDigito As String
		vsDigito = Modulo11(CStr(viNumeroAutorizacao))

		viNumeroAutorizacao = (viNumeroAutorizacao * 10) + CInt(vsDigito)

		SPP.ParamByName("P_VERSAOTISS").AsInteger = sql.FieldByName("HANDLE").AsInteger
		' SMS 104421 - TISS 2.2.1 - Danilo Raisi
		SPP.ParamByName("P_USUARIO").AsInteger    = CurrentUser
		SPP.ParamByName("P_WEBAUTORIZ").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		SPP.ParamByName("P_NUMEROAUTORIZACAO").AsFloat = viNumeroAutorizacao

		sql.Active = False

		Set sql = Nothing

		SPP.ExecProc

		' autorizador externo
		If SPP.ParamByName("P_RETORNO").AsString <> "" Then
			bsShowMessage(SPP.ParamByName("P_RETORNO").AsString, "I")
		End If

		Set SPP = Nothing

		' --- Fim da transação ----------------------
		' -------------------------------------------
		If InTransaction Then
			Commit
		End If
		' ===========================================
	End If
	WriteBDebugMessage("WEB_AUTORIZ.TABLE_AfterPost - Fim")
  Exit Sub
  Erro:
    WriteBDebugMessage("WEB_AUTORIZ.TABLE_AfterPost - Erro: " + Err.Description)
    InfoDescription = Err.Description
    CancelDescription = Err.Description
    If InTransaction Then
		Rollback
	End If
End Sub

Public Sub TABLE_AfterScroll()
  If WebVisionCode = "W_WEB_SADTAUTORIZPAGAMAENTO" Then
    If CurrentQuery.FieldByName("SITUACAOCONSULTA").Value = "C" Then
      ROTULOSITUACAO.Text = "@<TABLE BORDER='0' WIDTH='100%' CELLPADDING='0' CELLSPACING='0'><TR><TD ALIGN='CENTER'><FONT FACE='' SIZE='4' COLOR='#FF0000'>Operação Cancelada</FONT></TD></TR></TABLE>"
    Else
      ROTULOSITUACAO.Text = ""
    End If
  End If

  INDICADORDEACIDENTE.WebLocalWhere = " A.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
  TIPOINTERNACAO.WebLocalWhere = " A.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
  REGIMEINTERNACAO.WebLocalWhere = " A.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"

  ACOMODACAOEVENTO.WebLocalWhere = " A.EVENTO IN (SELECT MODEV.EVENTO" + Chr(13) + _
                                   "              FROM SAM_MODULO_EVENTO MODEV" + Chr(13) + _
                                   "              JOIN SAM_CONTRATO_MOD CM ON CM.MODULO = MODEV.MODULO" + Chr(13) + _
                                   "              JOIN SAM_BENEFICIARIO_MOD BM ON CM.MODULO = CM.MODULO" + Chr(13) + _
                                   "              WHERE BM.BENEFICIARIO = @CAMPO(BENEFICIARIO)" + Chr(13) + _
                                   "                AND BM.DATAADESAO <= @CAMPO(DATA)" + Chr(13) + _
                                   "                AND (BM.DATACANCELAMENTO IS NULL OR" + Chr(13) + _
                                   "                     (BM.DATACANCELAMENTO IS NOT NULL AND BM.DATACANCELAMENTO >= @CAMPO(DATA))))" + Chr(13) + _
                                   " AND A.ACOMODACAO IN (SELECT TIS.ACOMODACAO" + Chr(13) + _
                                   "                      FROM TIS_TIPOACOMODACAO TIS" + Chr(13) + _
                                   "                      WHERE TIS.VERSAOTISS = (SELECT MAX(HANDLE) VERSAOTISS FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'))" + Chr(13) + _
                                   " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
                                   "                  FROM SAM_TGE TGE" + Chr(13) + _
                                   "                  JOIN SAM_CLASSEEVENTO_CARATERATEND CCA ON CCA.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                   "                  WHERE TGE.HANDLE = A.EVENTO) OR" + Chr(13) + _
                                   "      EXISTS (SELECT 1" + Chr(13) + _
                                   "              FROM SAM_TGE TGE" + Chr(13) + _
                                   "              JOIN SAM_CLASSEEVENTO_CARATERATEND CCA ON CCA.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                   "              JOIN TIS_CARATERATENDIMENTO CAR ON CAR.HANDLE = CCA.CARATERATENDIMENTO" + Chr(13) + _
                                   "              WHERE TGE.HANDLE = A.EVENTO" + Chr(13) + _
                                   "                AND CAR.CODIGO = @CAMPO(CARATERATENDIMENTO)" + Chr(13) + _
                                   "                AND CAR.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')))" + Chr(13) + _
                                   " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
                                   "                  FROM SAM_TGE TGE" + Chr(13) + _
                                   "                  JOIN SAM_CLASSEEVENTO_REGIMEINTER CRI ON CRI.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                   "                  WHERE TGE.HANDLE = A.EVENTO) OR" + Chr(13) + _
                                   "      EXISTS (SELECT 1" + Chr(13) + _
                                   "              FROM SAM_TGE TGE" + Chr(13) + _
                                   "              JOIN SAM_CLASSEEVENTO_REGIMEINTER CRI ON CRI.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                   "              WHERE TGE.HANDLE = A.EVENTO" + Chr(13) + _
                                   "                AND CRI.REGIMEINTERNACAO = @CAMPO(REGIMEINTERNACAO)))" + Chr(13) + _
                                   " AND (NOT EXISTS (SELECT 1" + Chr(13) + _
                                   "                  FROM SAM_TGE TGE" + Chr(13) + _
                                   "                  JOIN SAM_CLASSEEVENTO_TIPOINTER CTI ON CTI.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                   "                  WHERE TGE.HANDLE = A.EVENTO) OR" + Chr(13) + _
                                   "      EXISTS (SELECT 1" + Chr(13) + _
                                   "              FROM SAM_TGE TGE" + Chr(13) + _
                                   "              JOIN SAM_CLASSEEVENTO_TIPOINTER CTI ON CTI.CLASSEEVENTO = TGE.CLASSEEVENTO" + Chr(13) + _
                                   "              WHERE TGE.HANDLE = A.EVENTO" + Chr(13) + _
                                   "                AND CTI.TIPOINTERNACAO = @CAMPO(TIPOINTERNACAO)))"
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		If (WebVisionCode = "V_WEB_CONSULTA_AUTORIZ") Or (WebVisionCode = "W_WEB_CONSULTA_AUTORIZ") Then
			If CurrentQuery.FieldByName("VALIDADO").AsString = "S" Then
				CancelDescription = "Não é possível editar o registro."
				CanContinue = False
			End If
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
    WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Início")
	Dim qSQL       As Object
	Set qSQL = NewQuery
	qSQL.Add("SELECT EXIGESOLICITANTE,")
	qSQL.Add("       EXIGERECEBEDOR,")
	qSQL.Add("       EXIGEEXECUTOR,")
	qSQL.Add("       EXIGELOCALEXECUCAO,")
	qSQL.Add("       TISSTIPOSOLICITACAO")
	qSQL.Add("FROM SAM_TIPOAUTORIZ")
	qSQL.Add("WHERE HANDLE = :HTIPOAUTORIZ")
	qSQL.ParamByName("HTIPOAUTORIZ").AsInteger = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
	qSQL.Active = True

	If WebMode Then
       If qSQL.FieldByName("TISSTIPOSOLICITACAO").AsString = "2" Then 'Somente se for internação
         WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Solicitação de Internação")
         Dim qParametrosAtendimento As BPesquisa
         Set qParametrosAtendimento = NewQuery

         qParametrosAtendimento.Add("SELECT TABCONTROLEDATASINTERNACAOALTA,")
         qParametrosAtendimento.Add("       DIASLIMITEINTERNACAORETROATIVA")
         qParametrosAtendimento.Add("FROM SAM_PARAMETROSATENDIMENTO")
         qParametrosAtendimento.Active = True

 		If qParametrosAtendimento.FieldByName("TABCONTROLEDATASINTERNACAOALTA").AsInteger = 2 Then
 		  WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Configurado para controlar datas de Internação/Alta")
 		  If CurrentQuery.FieldByName("DATAADMISSAOHOSPITALAR").IsNull Then
 			bsShowMessage("É obrigatório informar a data provável da admissão hospitalar!", "E")
 			CanContinue = False
 		  ElseIf WebVisionCode <> "V_WEB_INTERNACAO_AUT_REEMBOLSO" And (CurrentQuery.FieldByName("CONDICAOATENDIMENTO").IsNull Or Trim(CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsString) = "") Then
 			bsShowMessage("É obrigatório informar a condição do atendimento!", "E")
 			CanContinue = False
 		  Else
 		    If CurrentQuery.FieldByName("CARATERATENDIMENTO").AsInteger = 2 Then
 		      WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Caráter de atendimento de Urgência/Emergência")
 		      If CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsString = "1" Then
 			    bsShowMessage("Para Urgência/Emergência a condição não pode ser 'Paciente NÃO no local'!", "E")
 			    CanContinue = False
 		      Else
 			    CurrentQuery.FieldByName("DATAADMISSAOHOSPITALAR").AsDateTime = ServerDate
 			  End If
 		    Else
 		      WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Caráter de atendimento Eletivo")
 			  If CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsString = "1" Then
 			    WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Paciente NÃO no local")
 			    If CurrentQuery.FieldByName("DATAADMISSAOHOSPITALAR").AsDateTime < ServerDate Then
 			      WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Previsão de admissão hospitalar não pode ser retroativa")
 			      bsShowMessage("Para a condição 'Paciente NÃO no local' a data provável da admissão hospitalar não pode ser retroativa!", "E")
 			      CanContinue = False
 			    Else
 				  bsShowMessage("Será necessária a Confirmação da Internação na data efetiva da admissão hospitalar do beneficiário!", "I")
 			    End If
 			  ElseIf CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsString = "2" Then
 			    WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Paciente no local")
 			    If CurrentQuery.FieldByName("DATAADMISSAOHOSPITALAR").AsDateTime < ServerDate Then
 			      WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Previsão de admissão hospitalar não pode ser retroativa")
 			      bsShowMessage("Para a condição 'Paciente no local' a data provável da admissão hospitalar não pode ser retroativa!", "E")
 			      CanContinue = False
 			    ElseIf CurrentQuery.FieldByName("DATAADMISSAOHOSPITALAR").AsDateTime > ServerDate Then
 				  bsShowMessage("Será necessária a Confirmação da Internação na data efetiva da admissão hospitalar do beneficiário!", "I")
 			    End If
 			  ElseIf CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsString = "3" Then
 			    WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Paciente internado")
 			    If CurrentQuery.FieldByName("DATAADMISSAOHOSPITALAR").AsDateTime < ServerDate Then
 			      WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Previsão de admissão hospitalar retroativa")
 				  If CurrentQuery.FieldByName("DATAADMISSAOHOSPITALAR").AsDateTime < DateAdd("d", -qParametrosAtendimento.FieldByName("DIASLIMITEINTERNACAORETROATIVA").AsInteger, ServerDate) Then
 				    WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Previsão de admissão hospitalar anterior ao limite")
 			  	    Dim qUsuario As BPesquisa
 			  	    Set qUsuario = NewQuery

 			  	    qUsuario.Add("SELECT PERMITEINFORMARINTERRETROATIVA")
 			  	    qUsuario.Add("FROM Z_GRUPOUSUARIOS U")
 				    qUsuario.Add("JOIN Z_GRUPOS G ON G.HANDLE = U.GRUPO")
 			  	    qUsuario.Add("WHERE U.HANDLE = :HUSUARIO")
 			  	    qUsuario.ParamByName("HUSUARIO").AsInteger = CurrentUser
 			  	    qUsuario.Active = True

 			  	    If qUsuario.FieldByName("PERMITEINFORMARINTERRETROATIVA").AsString <> "S" Then
 			  	      WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Usuário sem permissão para data retroativa além do limite")
 			  		  bsShowMessage("Não é permitido informar previsão de admissão hospitalar retroativa!", "E")
 			  		  CanContinue = False
 			  	    End If
 			  	    Set qUsuario = Nothing
 				  End If
 			    ElseIf CurrentQuery.FieldByName("DATAADMISSAOHOSPITALAR").AsDateTime > ServerDate Then
 			      WriteBDebugMessage("WEB_AUTORIZ.TABLE_BeforePost - Previsão de admissão hospitalar não pode ser futura")
 				  bsShowMessage("Para a condição 'Paciente internado' a data provável da admissão hospitalar não pode ser futura!", "E")
 			      CanContinue = False
 			    End If
 			  End If
 		    End If
 		  End If
 		End If
         Set qParametrosAtendimento = Nothing
      End If

      If WebVisionCode = "W_WEB_CONSULTA_AUTORIZ" Then
		If Not CurrentQuery.FieldByName("RECEBEDOR").AsInteger > 0 Then
		  bsShowMessage("Campo Recebedor é obrigatório para esta autorização!", "E")

		  CanContinue = False
		End If

		If Not CurrentQuery.FieldByName("EXECUTOR").AsInteger > 0 Then
		  bsShowMessage("Campo Executor é obrigatório para esta autorização!", "E")

		  CanContinue = False
		End If
      ElseIf WebVisionCode = "V_WEB_CONSULTA_AUTORIZ"   Or _
             WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ" Or _
             WebVisionCode = "V_WEB_SPSADT_AUTORIZ"     Or _
             WebVisionCode = "V_WEB_ODONTO_AUTORIZ" Then
        'Verificar a obrigatoriedade dos campos conforme a configuração do tipo de autorização

		If qSQL.FieldByName("EXIGERECEBEDOR").AsString = "S" And _
		   CurrentQuery.FieldByName("RECEBEDOR").IsNull Then
 		  bsShowMessage("Recebedor é obrigatório para este Tipo de Autorização!", "E")

		  CanContinue = False
		End If

		If qSQL.FieldByName("EXIGEEXECUTOR").AsString = "S" And _
		   CurrentQuery.FieldByName("EXECUTOR").IsNull Then
		  bsShowMessage("Executor é obrigatório para este Tipo de Autorização!", "E")

		  CanContinue = False
		End If

		If qSQL.FieldByName("EXIGESOLICITANTE").AsString = "S" And _
		   CurrentQuery.FieldByName("SOLICITANTE").IsNull Then
		  bsShowMessage("Solicitante é obrigatório para este Tipo de Autorização!", "E")

		  CanContinue = False
		End If

		If qSQL.FieldByName("EXIGELOCALEXECUCAO").AsString = "S" And _
		   CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull Then
		  bsShowMessage("Local de Execução é obrigatório para este Tipo de Autorização!", "E")

		  CanContinue = False
		End If

        Set qSQL = Nothing

	  ElseIf (WebVisionCode = "V_WEB_INTERNACAO_AUT_REEMBOLSO" Or _
             WebVisionCode = "V_WEB_SPSADT_AUT_REEMBOLSO" Or _
             WebVisionCode = "V_WEB_ODONTO_AUT_REEMBOLSO") And _
             CurrentQuery.FieldByName("EXECUTOR").IsNull Then

        bsShowMessage("Executor é obrigatório para este Tipo de Autorização!", "E")

	  	CanContinue = False

      End If

	  If WebVisionCode = "W_WEB_CONSULTA_AUTORIZ" Then
		If Not (CurrentQuery.FieldByName("EVENTO").AsInteger > 0) Then
		  bsShowMessage("Evento obrigatório para esta autorização!", "E")
		  CanContinue = False
		End If
	  End If

      If (CurrentQuery.FieldByName("EVENTO").AsInteger <= 0) Then
  	    CurrentQuery.FieldByName("EVENTO").Clear
	  End If

	 If Not CurrentQuery.FieldByName("CARATERATENDIMENTO").IsNull And Not CurrentQuery.FieldByName("TIPOATENDIMENTO").IsNull Then
	     Dim qCarater As BPesquisa
	     Set qCarater = NewQuery

	     qCarater.Add("SELECT COUNT(1) QTDE ")
	     qCarater.Add("  FROM TIS_TIPOATENDIMENTO A ")
	     qCarater.Add("  JOIN TIS_CARATERATENDIMENTO B ON (B.HANDLE = A.CARATERSOLICITACAO) ")
	     qCarater.Add(" WHERE A.HANDLE = :TIPOATENDIMENTO ")
	     If CurrentQuery.FieldByName("CARATERATENDIMENTO").AsString = "1" Then
	     	qCarater.Add("   AND B.CODIGO IN ('1', 'E') ")
	     ElseIf CurrentQuery.FieldByName("CARATERATENDIMENTO").AsString = "2" Then
			qCarater.Add("   AND B.CODIGO IN ('2', 'U') ")
	     End If
	     qCarater.ParamByName("TIPOATENDIMENTO").AsInteger = CurrentQuery.FieldByName("TIPOATENDIMENTO").AsInteger
	     qCarater.Active = True

	     If Not (qCarater.FieldByName("QTDE").AsInteger >0) Then
	 	     bsShowMessage("Tipo de Atendimento não se refere ao 'Caráter de Atendimento' selecionado!", "E")
	         CanContinue = False

		     Set qCarater= Nothing
			 Exit Sub
	     End If
	     Set qCarater = Nothing
 	 End If


	If Not (CurrentQuery.FieldByName("CBOSSOLICITANTE").AsInteger > 0) Then
		Select Case WebVisionCode
		  Case "W_WEB_SADTAUTORIZ", "W_WEB_SADTAUTORIZPAGAMAENTO", "W_WEB_CONSULTA_CANCELAMENTO", _
	           "W_WEB_CONSULTA_GUIA", "W_WEB_CONSULTA_GUIA_LEITURA", "W_WEB_SADTCANCELAMENTO"

	        Dim vObrigarCamposTissWeb As Boolean
	        Dim vDllEspec As Object
			Set vDllEspec = CreateBennerObject("Especifico.UEspecifico")
			vObrigarCamposTissWeb = vDllEspec.AUT_ExigeCamposTISSWeb(CurrentSystem)
			Set vDllEspec = Nothing

			If vObrigarCamposTissWeb Then
	           bsShowMessage("Campo CBOS Solicitante é obrigatório", "E")
	           CanContinue = False
			End If

	    End Select
	End If

      '---------------------------------------------------------------------
      If WebVisionCode = "V_WEB_CONSULTA_AUTORIZ" Or _
         WebVisionCode = "W_WEB_CONSULTA_ELEGIBILIDADE" Or (WebVisionCode = "W_WEB_CONSULTA_AUTORIZ") Then

        CurrentQuery.FieldByName("TIPOOPERACAOTISS").Value = "C"
      ElseIf WebVisionCode = "W_WEB_SADTAUTORIZPAGAMAENTO" Or _
             WebVisionCode = "W_WEB_SADTAUTORIZ" Or _
             WebVisionCode = "W_WEB_SADTELEGIBILIDADE"  Then

         CurrentQuery.FieldByName("TIPOOPERACAOTISS").Value = "S"
      ElseIf WebVisionCode = "V_WEB_ODONTO_AUTORIZ" Or WebVisionCode = "V_WEB_ODONTO_AUT_REEMBOLSO" Or _
             WebVisionCode = "W_WEB_ODONTO_AUTORIZ" Then
      	 CurrentQuery.FieldByName("TIPOOPERACAOTISS").Value = "O"
      End If
      '---------------------------------------------------------------------
	End If
End Sub

Public Sub TABLE_NewRecord()
	If WebMode Then
		Dim vSQL As Object
		Set vSQL = NewQuery

		vSQL.Clear
		vSQL.Add("SELECT HANDLE                                    ")
		vSQL.Add("  FROM TIS_TIPOATENDIMENTO                       ")
		vSQL.Add(" WHERE CODIGO = :COD                             ")
		vSQL.Add("   And VERSAOTISS In (Select MAX (HANDLE)        ")
		vSQL.Add("                        FROM TIS_VERSAO          ")
		vSQL.Add("                       WHERE ATIVODESKTOP = 'S') ")

		vSQL.ParamByName("COD").AsInteger = 5
		vSQL.Active = True

		If WebVisionCode = "W_WEB_SADTAUTORIZPAGAMAENTO" Then
			CurrentQuery.FieldByName("TIPOATENDIMENTO").AsInteger = vSQL.FieldByName("HANDLE").AsInteger
		ElseIf WebVisionCode = "W_WEB_SADTAUTORIZ" Then
			CurrentQuery.FieldByName("TIPOATENDIMENTO").AsInteger = vSQL.FieldByName("HANDLE").AsInteger
		End If

		vSQL.Active = False
		vSQL.Clear
		vSQL.Add("SELECT COUNT(1) QTD FROM Z_GRUPOUSUARIOS_PRESTADORAEX")
		vSQL.Add(" WHERE EXECUTOR = 'S' AND USUARIO = :USER")
		vSQL.ParamByName("USER").AsInteger = CurrentUser
		vSQL.Active = True

		If vSQL.FieldByName("QTD").AsInteger = 1 Then
			vSQL.Active = False
			vSQL.Clear
			vSQL.Add("SELECT PRESTADOR, EVENTOPADRAO FROM Z_GRUPOUSUARIOS_PRESTADORAEX")
			vSQL.Add(" WHERE EXECUTOR = 'S' AND USUARIO = :USER")
			vSQL.ParamByName("USER").AsInteger = CurrentUser
			vSQL.Active = True
			CurrentQuery.FieldByName("EXECUTOR").Value = vSQL.FieldByName("PRESTADOR").AsInteger

			If (WebVisionCode <> "W_WEB_SADTELEGIBILIDADE") And (WebVisionCode <> "W_WEB_SADTPAGAMENTO") Then
			  CurrentQuery.FieldByName("EVENTO").Value = vSQL.FieldByName("EVENTOPADRAO").AsInteger
			End If

			Dim SPROC As Object

			Set SPROC = NewStoredProc
			SPROC.AutoMode = True
			SPROC.Name = "BSAut_AutorizVerificaEndereco"
			SPROC.AddParam("P_EXECUTOR",ptInput)
			SPROC.ParamByName("P_EXECUTOR").DataType = ftInteger 'SMS 95930 - Marcelo Barbosa - 14/04/2008
			SPROC.AddParam("P_ENDERECO",ptOutput)
			SPROC.ParamByName("P_ENDERECO").DataType = ftInteger 'SMS 95930 - Marcelo Barbosa - 14/04/2008
			SPROC.AddParam("P_TEMMAISDEUMENDERECO",ptOutput)
			SPROC.ParamByName("P_TEMMAISDEUMENDERECO").DataType = ftString 'SMS 95930 - Marcelo Barbosa - 14/04/2008

			SPROC.ParamByName("P_EXECUTOR").AsInteger = vSQL.FieldByName("PRESTADOR").AsInteger

			SPROC.ExecProc

			If SPROC.ParamByName("P_TEMMAISDEUMENDERECO").AsString = "N" And SPROC.ParamByName("P_ENDERECO").AsInteger > 0 Then
				CurrentQuery.FieldByName("ENDERECO").Value = SPROC.ParamByName("P_ENDERECO").AsInteger
			ElseIf SPROC.ParamByName("P_TEMMAISDEUMENDERECO").AsString = "N" And SPROC.ParamByName("P_ENDERECO").AsInteger = -1 Then
				bsShowMessage("Não foram encontrados endereços para o prestador executor.", "I")
			ElseIf SPROC.ParamByName("P_TEMMAISDEUMENDERECO").AsString = "S"  Then
				bsShowMessage("O Executor possui mais de um endereço válido, escolha um.", "I")
			End If
			Set SPROC = Nothing
		End If

		vSQL.Active = False
		vSQL.Clear
		vSQL.Add("SELECT COUNT(1) QTD FROM Z_GRUPOUSUARIOS_PRESTADORAEX")
		vSQL.Add(" WHERE RECEBEDOR = 'S' AND USUARIO = :USER")
		vSQL.ParamByName("USER").AsInteger = CurrentUser
		vSQL.Active = True

		If vSQL.FieldByName("QTD").AsInteger = 1 Then
			vSQL.Active = False
			vSQL.Clear
			vSQL.Add("SELECT PRESTADOR, EVENTOPADRAO FROM Z_GRUPOUSUARIOS_PRESTADORAEX")
			vSQL.Add(" WHERE RECEBEDOR = 'S' AND USUARIO = :USER")
			vSQL.ParamByName("USER").AsInteger = CurrentUser
			vSQL.Active = True
			CurrentQuery.FieldByName("RECEBEDOR").Value = vSQL.FieldByName("PRESTADOR").AsInteger
		End If

		vSQL.Active = False
		vSQL.Clear
		vSQL.Add("SELECT COUNT(1) QTD FROM Z_GRUPOUSUARIOS_PRESTADORAEX")
		vSQL.Add(" WHERE LOCALEXECUCAO = 'S' AND USUARIO = :USER")
		vSQL.ParamByName("USER").AsInteger = CurrentUser
		vSQL.Active = True

		If vSQL.FieldByName("QTD").AsInteger = 1 Then
			vSQL.Active = False
			vSQL.Clear
			vSQL.Add("SELECT PRESTADOR, EVENTOPADRAO FROM Z_GRUPOUSUARIOS_PRESTADORAEX")
			vSQL.Add(" WHERE LOCALEXECUCAO = 'S' AND USUARIO = :USER")
			vSQL.ParamByName("USER").AsInteger = CurrentUser
			vSQL.Active = True
			CurrentQuery.FieldByName("LOCALEXECUCAO").Value = vSQL.FieldByName("PRESTADOR").AsInteger
		End If

		vSQL.Active = False
		vSQL.Clear
		vSQL.Add("SELECT COUNT(1) QTD FROM Z_GRUPOUSUARIOS_PRESTADORAEX")
		vSQL.Add(" WHERE SOLICITANTE = 'S' AND USUARIO = :USER")
		vSQL.ParamByName("USER").AsInteger = CurrentUser
		vSQL.Active = True

		If vSQL.FieldByName("QTD").AsInteger = 1 Then
			vSQL.Active = False
			vSQL.Clear
			vSQL.Add("SELECT PRESTADOR, EVENTOPADRAO FROM Z_GRUPOUSUARIOS_PRESTADORAEX")
			vSQL.Add(" WHERE SOLICITANTE = 'S' AND USUARIO = :USER")
			vSQL.ParamByName("USER").AsInteger = CurrentUser
			vSQL.Active = True
			CurrentQuery.FieldByName("SOLICITANTE").Value = vSQL.FieldByName("PRESTADOR").AsInteger
		End If

		Set vSQL = Nothing
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  WriteBDebugMessage("WEB_AUTORIZ - OnCommnandClick")
  Dim SQLx As Object
  Dim viNumeroAutorizacao As Long
  Dim SQLAgendado As Object


  If WebMode Then

	Set SQLAgendado = NewQuery

	SQLAgendado.Add("SELECT EXECUTAAUTORIZACAOAGENDADA FROM SAM_PARAMETROSWEB")
	SQLAgendado.Active = True

	' Caso o parametro geral de web "Executa Autorização Agendada" esteja marcado
	If SQLAgendado.FieldByName("EXECUTAAUTORIZACAOAGENDADA").AsString = "S" Then

		Dim vcContainer As CSDContainer
		Dim vsMensagemErro As String
		Dim Obj As Object
        Dim viRet As Long
		Set vcContainer = NewContainer
       	vcContainer.AddFields("HANDLE:INTEGER")
       	vcContainer.AddFields("P_TIPOOPERACAO:INTEGER")
       	vcContainer.AddFields("P_AUTORIZACAO:INTEGER")
       	vcContainer.AddFields("P_TIPOTISS:STRING")
       	vcContainer.AddFields("P_ORIGEM:STRING")
       	vcContainer.AddFields("P_USUARIO:INTEGER")
       	vcContainer.AddFields("P_WEBAUTORIZ:INTEGER")
       	vcContainer.AddFields("P_VERSAOTISS:INTEGER")
       	vcContainer.AddFields("P_NUMEROAUTORIZACAO:DOUBLE")
       	vcContainer.AddFields("P_EHREEMBOLSO:STRING")
       	vcContainer.AddFields("P_RETORNO:INTEGER")


		vcContainer.Insert
'		vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

		If WebVisionCode = "W_WEB_CONSULTA_CONSULTAS" And CommandID = "ATUALIZAR" Then
			vcContainer.Field("P_TIPOOPERACAO").AsInteger = 110
			vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
			vcContainer.Field("P_TIPOTISS").AsString = "S"
			vcContainer.Field("P_ORIGEM").AsString = "W"
			vcContainer.Field("P_USUARIO").AsInteger = CurrentUser
			vcContainer.Field("P_WEBAUTORIZ").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

			If vcContainer.Field("P_RETORNO").AsString <> "" Then
				bsShowMessage(vcContainer.Field("P_RETORNO").AsString, "I")
			End If

		ElseIf CommandID = "IMPRIMIR" Then
			'Alterado para permitir reimpressão de autorizações SP/SADT  SMS - 111851
			Call ChamarRelatorioAutorizacao
		Else
			vcContainer.Field("P_TIPOOPERACAO").AsInteger = 0 ' iniciar com zero, se mudar a procedure deve ser executada
			If ((WebVisionCode = "V_WEB_CONSULTA_AUTORIZ" And CommandID = "CANCELAAUTORIZ") Or _
	                (WebVisionCode = "V_WEB_CONSULTA_AUTORIZ" And CommandID = "CANCELAGUIA") Or _
			        (WebVisionCode = "W_WEB_CONSULTA_CANCELAMENTO" And CommandID = "CANCELARGUIA") Or _
			        (WebVisionCode = "W_WEB_CONSULTA_GUIA" And (CommandID = "CANCELARGUIA" Or CommandID = "CANCELARAUTO"))) Then
				vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
				vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
				vcContainer.Field("P_TIPOTISS").AsString = "C"
			ElseIf (WebVisionCode = "V_WEB_CONSULTA_AUTORIZ") Or (WebVisionCode = "W_WEB_CONSULTA_GUIA") Or (WebVisionCode = "W_WEB_CONSULTA_AUTORIZ") Then
				Select Case CommandID
					Case "ATUALIZAR"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 120
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = CurrentQuery.FieldByName("TIPOOPERACAOTISS").AsString
	                Case "GERARGUIA"
	     			    vcContainer.Field("P_TIPOOPERACAO").AsInteger = 110
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
					    vcContainer.Field("P_TIPOTISS").AsString = CurrentQuery.FieldByName("TIPOOPERACAOTISS").AsString
					'Luciano T. Alberti - SMS 94564 - 13/03/2008 - Início
					Case "IMPRIMIRGUIACONSULTA"
						Call ChamarRelatorioAutorizacao
					'Luciano T. Alberti - SMS 94564 - 13/03/2008 - Fim
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTAUTORIZPAGAMAENTO" Then '====================================================================================================
				Select Case CommandID
					Case "VALIDARSADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 100
						vcContainer.Field("P_AUTORIZACAO").AsInteger = -1
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "ATUALIZARSADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 110
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "CANCELARSADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 140
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "SADTPAGCANCELAR"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					'Luciano T. Alberti - SMS 94808 - 09/09/2008 - Início
					Case "IMPRIMIRGUIASPSADT"
						Call ChamarRelatorioAutorizacao
					'Luciano T. Alberti - SMS 94808 - 09/09/2008 - Fim
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTELEGIBILIDADE" Then '====================================================================================================
				Select Case CommandID
					Case "VALIDARSADTELEGIBILIDADE"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 160
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_CONSULTA_EXECUCAO" Then '====================================================================================================
				Select Case CommandID
					Case "GERARGUIA"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 110
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "CANCELARGUIA"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTAUTORIZ" Then '====================================================================================================
				Select Case CommandID
					Case "CANCELARSADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 140
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "ATUALIZASADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 120
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "VALIDARSADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 120
						vcContainer.Field("P_AUTORIZACAO").AsInteger = -1
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					'Luciano T. Alberti - SMS 94808 - 09/09/2008 - Início
					Case "IMPRIMIRGUIASPSADT"
						Call ChamarRelatorioAutorizacao
					'Luciano T. Alberti - SMS 94808 - 09/09/2008 - Fim
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTPAGAMENTO" Then '====================================================================================================
				Select Case CommandID
		            Case "SADTPAGCANCAUTORIZ"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "SADTPAGCANCGUIA"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "SADTPAGVALIDAR"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 100
						vcContainer.Field("P_AUTORIZACAO").AsInteger = -1
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "ATUALIZAR"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 110
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTCANCELAMENTO" Then '====================================================================================================
				Select Case CommandID
					Case "CANCELARSADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
					Case "CANCELARGUIA"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_CONSULTA_CANCELAMENTO" Then '====================================================================================================
				Select Case CommandID
					Case "CANCELARAUTORIZACAO"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "C"
				End Select
			ElseIf WebVisionCode = "V_WEB_SPSADT_AUTORIZ" Then '====================================================================================================
				Select Case CommandID
					Case "VALIDARSPSADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 121
						vcContainer.Field("P_AUTORIZACAO").AsInteger = 0
						vcContainer.Field("P_TIPOTISS").AsString = "S"
				End Select
			ElseIf WebVisionCode = "V_WEB_SPSADT_AUT_REEMBOLSO" Then
				Select Case CommandID
					Case "VALIDARSPSADT"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 121
						vcContainer.Field("P_AUTORIZACAO").AsInteger = 0
						vcContainer.Field("P_TIPOTISS").AsString = "S"
						vcContainer.Field("P_EHREEMBOLSO").AsString = "S"
				End Select
			ElseIf WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ" Then '====================================================================================================
				Select Case CommandID
					Case "VALIDARINTERNACAO"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 120
						vcContainer.Field("P_AUTORIZACAO").AsInteger = -1
						vcContainer.Field("P_TIPOTISS").AsString = "I"
					Case "CANCELARSOLICITINTERNACAO"
                        Call CancelarInternacao
				End Select
			ElseIf WebVisionCode = "V_WEB_INTERNACAO_AUT_REEMBOLSO" Then
				Select Case CommandID
					Case "VALIDARINTERNACAO"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 121
						vcContainer.Field("P_AUTORIZACAO").AsInteger = -1
						vcContainer.Field("P_TIPOTISS").AsString = "I"
						vcContainer.Field("P_EHREEMBOLSO").AsString = "S"
				End Select
			ElseIf WebVisionCode = "V_WEB_ODONTO_AUTORIZ" Or WebVisionCode = "W_WEB_ODONTO_AUTORIZ" Then 'SMS 90455 - Ricardo Rocha - 04/06/2008
				Select Case CommandID
					Case "VALIDARODONTO"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 120
						vcContainer.Field("P_AUTORIZACAO").AsInteger = -1
						vcContainer.Field("P_TIPOTISS").AsString = "O"
				End Select
			ElseIf WebVisionCode = "V_WEB_ODONTO_AUT_REEMBOLSO" Then
				Select Case CommandID
					Case "VALIDARODONTO"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 121
						vcContainer.Field("P_AUTORIZACAO").AsInteger = -1
						vcContainer.Field("P_TIPOTISS").AsString = "O"
						vcContainer.Field("P_EHREEMBOLSO").AsString = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_ODONTO_CANCELAMENTO" Then
				Select Case CommandID
					Case "CANCELARODONTO"
						vcContainer.Field("P_TIPOOPERACAO").AsInteger = 130
						vcContainer.Field("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						vcContainer.Field("P_TIPOTISS").AsString = "O"
				End Select
			End If

            If WebVisionCode = "V_WEB_CONSULTA_AUTORIZ" Or _
               WebVisionCode = "V_WEB_SPSADT_AUTORIZ"   Or _
               WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ" Or _
               WebVisionCode = "W_WEB_SADTAUTORIZ" Or _
               WebVisionCode = "V_WEB_ODONTO_AUTORIZ" Or _
               WebVisionCode = "V_WEB_INTERNACAO_AUT_REEMBOLSO" Or _
               WebVisionCode = "V_WEB_SPSADT_AUT_REEMBOLSO"   Or _
               WebVisionCode = "V_WEB_ODONTO_AUT_REEMBOLSO" Then
              vcContainer.Field("P_ORIGEM").AsString = "1"
            Else
              vcContainer.Field("P_ORIGEM").AsString = "2"
            End If

			If vcContainer.Field("P_TIPOOPERACAO").AsInteger > 0 Then
			    If Not InTransaction Then
					StartTransaction
				End If

				Dim sql As BPesquisa
				Set sql = NewQuery

				sql.Add("SELECT MAX(HANDLE) HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'")
				sql.Active = True

				NewCounter2("SAM_AUTORIZ", 0, 1, viNumeroAutorizacao)

				Dim vsDigito As String
				vsDigito = Modulo11(CStr(viNumeroAutorizacao))

				viNumeroAutorizacao = (viNumeroAutorizacao * 10) + CInt(vsDigito)

				vcContainer.Field("P_VERSAOTISS").AsInteger = sql.FieldByName("HANDLE").AsInteger
				vcContainer.Field("P_USUARIO").AsInteger = CurrentUser
				vcContainer.Field("P_WEBAUTORIZ").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
				vcContainer.Field("P_NUMEROAUTORIZACAO").AsDouble = viNumeroAutorizacao


				sql.Active = False

				Set sql = Nothing

				vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

				Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
		  		viRet = Obj.ExecucaoImediata(CurrentSystem, _
                                             "SAMVALIDACAO", _
                                             "GerarAutorizacaoWeb", _
                                             "Processamento de Autorização", _
                                             CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                             "WEB_AUTORIZ", _
                                             "SITUACAOPROCESSAMENTO", _
                                             "", _
                                             "", _
                                             "P", _
                                             True, _
                                             vsMensagemErro, _
                                             vcContainer)
 'CurrentQuery.FieldByName("HANDLE").AsInteger,
				If viRet = 0 Then
				 	bsShowMessage("Processo enviado para execução no servidor!", "I")
				 	bsShowMessage("Autorização sendo processada!", "I")

				Else
			     	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
			   	End If

				If vcContainer.Field("P_RETORNO").AsString <> "" Then
				End If

				If InTransaction Then
					Commit
				End If
			End If
			Set Obj = Nothing
		End If
	' Caso o parametro geral de web "Executa Autorização Agendada" não esteja marcado
	Else
		Dim SPP As BStoredProc

		If InStr(SQLServer, "SQL") > 0 Then
			Set SQLx = NewQuery
			On Error GoTo TabelasTemporarias
			SQLx.Clear
			SQLx.Add("SELECT 1 FROM #TMP_ORIGEMCALCULO")
			SQLx.Active = True

			GoTo Procedure
			TabelasTemporarias:
			CriaTabelaTemporariaSqlServer
			Set SQLx = Nothing
		End If
		Procedure:
		On Error GoTo Erro
		Set SPP = NewStoredProc
		SPP.AutoMode = True
		WriteBDebugMessage("Preparando chamada da BSAUT_AUTORIZWEB")
		SPP.Name = "BSAUT_AUTORIZWEB"
		' SMS 104421 - TISS 2.2.1 - Danilo Raisi
		SPP.AddParam("P_VERSAOTISS",ptInput)			'Int
		SPP.ParamByName("P_VERSAOTISS").DataType   		= ftInteger
		' SMS 104421 - TISS 2.2.1 - Danilo Raisi
		SPP.AddParam("P_WEBAUTORIZ",ptInput,ftInteger)        	'Int
		SPP.AddParam("P_TIPOOPERACAO",ptInput,ftInteger)      	'Int
		SPP.AddParam("P_AUTORIZACAO",ptInput,ftInteger)       	'Int
		SPP.AddParam("P_TIPOTISS",ptInput,ftString)          	'Varchar(1)
		SPP.AddParam("P_ORIGEM",ptInput,ftString)            	'Varchar(1)
		SPP.AddParam("P_USUARIO",ptInput,ftInteger)           	'Int
		SPP.AddParam("P_NUMEROAUTORIZACAO", ptInput,ftFloat)
		SPP.AddParam("P_EHREEMBOLSO", ptInput, ftString)
		SPP.AddParam("P_RETORNO",ptOutput,ftString)          	'Varchar(100)


		If WebVisionCode = "W_WEB_CONSULTA_CONSULTAS" And CommandID = "ATUALIZAR" Then
            If Not InTransaction Then
				StartTransaction
			End If

			SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 110
			SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
			SPP.ParamByName("P_TIPOTISS").AsString      = "S"
			SPP.ParamByName("P_ORIGEM").AsString        = "W"
			SPP.ParamByName("P_USUARIO").AsInteger      = CurrentUser
			SPP.ParamByName("P_WEBAUTORIZ").AsInteger   = CurrentQuery.FieldByName("HANDLE").AsInteger
			SPP.ExecProc
			If SPP.ParamByName("P_RETORNO").AsString <> "" Then
				bsShowMessage(SPP.ParamByName("P_RETORNO").AsString, "I")
			End If

			If InTransaction Then
				Commit
			End If
		ElseIf CommandID = "IMPRIMIR" Then
			'Alterado para permitir reimpressão de autorizações SP/SADT  SMS - 111851
			Call ChamarRelatorioAutorizacao
		Else
			SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 0 ' iniciar com zero, se mudar a procedure deve ser executada
			If ((WebVisionCode = "V_WEB_CONSULTA_AUTORIZ" And CommandID = "CANCELAAUTORIZ") Or _
	                (WebVisionCode = "V_WEB_CONSULTA_AUTORIZ" And CommandID = "CANCELAGUIA") Or _
			        (WebVisionCode = "W_WEB_CONSULTA_CANCELAMENTO" And CommandID = "CANCELARGUIA") Or _
			        (WebVisionCode = "W_WEB_CONSULTA_GUIA" And (CommandID = "CANCELARGUIA" Or CommandID = "CANCELARAUTO"))) Then
				SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
				SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
				SPP.ParamByName("P_TIPOTISS").AsString      = "C"
			ElseIf (WebVisionCode = "V_WEB_CONSULTA_AUTORIZ") Or (WebVisionCode = "W_WEB_CONSULTA_GUIA") Or (WebVisionCode = "W_WEB_CONSULTA_AUTORIZ") Then
				Select Case CommandID
					Case "ATUALIZAR"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 120
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = CurrentQuery.FieldByName("TIPOOPERACAOTISS").AsString
	                Case "GERARGUIA"
	     			    SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 110
					    SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
					    SPP.ParamByName("P_TIPOTISS").AsString      = CurrentQuery.FieldByName("TIPOOPERACAOTISS").AsString
					'Luciano T. Alberti - SMS 94564 - 13/03/2008 - Início
					Case "IMPRIMIRGUIACONSULTA"
						Call ChamarRelatorioAutorizacao
					'Luciano T. Alberti - SMS 94564 - 13/03/2008 - Fim
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTAUTORIZPAGAMAENTO" Then '====================================================================================================
				Select Case CommandID
					Case "VALIDARSADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 100
						SPP.ParamByName("P_AUTORIZACAO").Value      = Null
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "ATUALIZARSADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 110
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "CANCELARSADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 140
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "SADTPAGCANCELAR"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					'Luciano T. Alberti - SMS 94808 - 09/09/2008 - Início
					Case "IMPRIMIRGUIASPSADT"
						Call ChamarRelatorioAutorizacao
					'Luciano T. Alberti - SMS 94808 - 09/09/2008 - Fim
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTELEGIBILIDADE" Then '====================================================================================================
				Select Case CommandID
					Case "VALIDARSADTELEGIBILIDADE"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 160
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_CONSULTA_EXECUCAO" Then '====================================================================================================
				Select Case CommandID
					Case "GERARGUIA"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 110
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "CANCELARGUIA"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTAUTORIZ" Then '====================================================================================================
				Select Case CommandID
					Case "CANCELARSADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 140
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "ATUALIZASADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 120
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "VALIDARSADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 120
						SPP.ParamByName("P_AUTORIZACAO").Value      = Null
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					'Luciano T. Alberti - SMS 94808 - 09/09/2008 - Início
					Case "IMPRIMIRGUIASPSADT"
						Call ChamarRelatorioAutorizacao
					'Luciano T. Alberti - SMS 94808 - 09/09/2008 - Fim
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTPAGAMENTO" Then '====================================================================================================
				Select Case CommandID
		            Case "SADTPAGCANCAUTORIZ"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "SADTPAGCANCGUIA"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "SADTPAGVALIDAR"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 100
						SPP.ParamByName("P_AUTORIZACAO").Value      = Null
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "ATUALIZAR"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 110
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_SADTCANCELAMENTO" Then '====================================================================================================
				Select Case CommandID
					Case "CANCELARSADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
					Case "CANCELARGUIA"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_CONSULTA_CANCELAMENTO" Then '====================================================================================================
				Select Case CommandID
					Case "CANCELARAUTORIZACAO"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
						SPP.ParamByName("P_AUTORIZACAO").Value      = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString      = "C"
				End Select
			ElseIf WebVisionCode = "V_WEB_SPSADT_AUTORIZ" Then '====================================================================================================
				Select Case CommandID
					Case "VALIDARSPSADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
						SPP.ParamByName("P_AUTORIZACAO").AsInteger      = -1
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
				End Select
			ElseIf WebVisionCode = "V_WEB_SPSADT_AUT_REEMBOLSO" Then
				Select Case CommandID
					Case "VALIDARSPSADT"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
						SPP.ParamByName("P_AUTORIZACAO").AsInteger  = -1
						SPP.ParamByName("P_TIPOTISS").AsString      = "S"
						SPP.ParamByName("P_EHREEMBOLSO").AsString   = "S"
				End Select
			ElseIf WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ" Then '====================================================================================================
				Select Case CommandID
					Case "VALIDARINTERNACAO"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 120
						SPP.ParamByName("P_AUTORIZACAO").AsInteger      = -1
						SPP.ParamByName("P_TIPOTISS").AsString      = "I"
					Case "CANCELARSOLICITINTERNACAO"
                        Call CancelarInternacao
				End Select
			ElseIf WebVisionCode = "V_WEB_INTERNACAO_AUT_REEMBOLSO" Then
				Select Case CommandID
					Case "VALIDARINTERNACAO"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
						SPP.ParamByName("P_AUTORIZACAO").AsInteger  = -1
						SPP.ParamByName("P_TIPOTISS").AsString      = "I"
						SPP.ParamByName("P_EHREEMBOLSO").AsString   = "S"
				End Select
			ElseIf WebVisionCode = "V_WEB_ODONTO_AUTORIZ" Or WebVisionCode = "W_WEB_ODONTO_AUTORIZ" Then 'SMS 90455 - Ricardo Rocha - 04/06/2008
				Select Case CommandID
					Case "VALIDARODONTO"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 120
						SPP.ParamByName("P_AUTORIZACAO").AsInteger		= -1
						SPP.ParamByName("P_TIPOTISS").AsString		= "O"
				End Select
			ElseIf WebVisionCode = "V_WEB_ODONTO_AUT_REEMBOLSO" Then
				Select Case CommandID
					Case "VALIDARODONTO"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 121
						SPP.ParamByName("P_AUTORIZACAO").AsInteger	= -1
						SPP.ParamByName("P_TIPOTISS").AsString		= "O"
						SPP.ParamByName("P_EHREEMBOLSO").AsString	= "S"
				End Select
			ElseIf WebVisionCode = "W_WEB_ODONTO_CANCELAMENTO" Then
				Select Case CommandID
					Case "CANCELARODONTO"
						SPP.ParamByName("P_TIPOOPERACAO").AsInteger = 130
						SPP.ParamByName("P_AUTORIZACAO").AsInteger	= CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
						SPP.ParamByName("P_TIPOTISS").AsString		= "O"
				End Select
			End If

            If WebVisionCode = "V_WEB_CONSULTA_AUTORIZ" Or _
               WebVisionCode = "V_WEB_SPSADT_AUTORIZ"   Or _
               WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ" Or _
               WebVisionCode = "V_WEB_ODONTO_AUTORIZ" Or _
               WebVisionCode = "V_WEB_INTERNACAO_AUT_REEMBOLSO" Or _
               WebVisionCode = "V_WEB_ODONTO_AUT_REEMBOLSO" Or _
               WebVisionCode = "V_WEB_SPSADT_AUT_REEMBOLSO" Then
              SPP.ParamByName("P_ORIGEM").AsString        = "1"
            Else
              SPP.ParamByName("P_ORIGEM").AsString        = "2"
            End If

			If SPP.ParamByName("P_TIPOOPERACAO").AsInteger > 0 Then

				If Not InTransaction Then
					StartTransaction
				End If

				Dim sql2 As BPesquisa
				Set sql2 = NewQuery

                sql2.Clear
				sql2.Add("SELECT MAX(HANDLE) HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'")
				sql2.Active = True

				NewCounter2("SAM_AUTORIZ", 0, 1, viNumeroAutorizacao)

				Dim vsDigito2 As String
				vsDigito2 = Modulo11(CStr(viNumeroAutorizacao))

                If vsDigito = "" Then
                  viNumeroAutorizacao = (viNumeroAutorizacao * 10) + CInt("0")
                Else
				  viNumeroAutorizacao = (viNumeroAutorizacao * 10) + CInt(vsDigito)
				End If

				SPP.ParamByName("P_VERSAOTISS").AsInteger   = sql2.FieldByName("HANDLE").AsInteger
				' SMS 104421 - TISS 2.2.1 - Danilo Raisi
				SPP.ParamByName("P_USUARIO").AsInteger      = CurrentUser
				SPP.ParamByName("P_WEBAUTORIZ").AsInteger   = CurrentQuery.FieldByName("HANDLE").AsInteger
				SPP.ParamByName("P_NUMEROAUTORIZACAO").AsFloat = viNumeroAutorizacao

				sql2.Active = False

				Set sql2 = Nothing

				WriteBDebugMessage("Executar BSAUT_AUTORIZWEB")
				SPP.ExecProc
				WriteBDebugMessage("BSAUT_AUTORIZWEB executada")

				If SPP.ParamByName("P_RETORNO").AsString <> "" Then
					WriteBDebugMessage("Retorno da BSAUT_AUTORIZWEB: " + SPP.ParamByName("P_RETORNO").AsString)
					bsShowMessage(SPP.ParamByName("P_RETORNO").AsString, "I")
				End If

				If InTransaction Then
					Commit
				End If
			End If
		End If
		Set SPP = Nothing

		'Foi incluído esta query, pois neste momento a macro ainda não foi atualizada, não contendo o NUMEROAUTORIZACAO da CurrentQuery
		Dim HandleAut As Object
    	Set HandleAut = NewQuery

    	HandleAut.Active = False
		HandleAut.Clear
		HandleAut.Add("SELECT NUMEROAUTORIZACAO, PROTOCOLOTRANSACAO FROM WEB_AUTORIZ WHERE HANDLE = :HANDLEWEB")
		HandleAut.ParamByName("HANDLEWEB").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		HandleAut.Active = True


		If (WebVisionCode = "W_WEB_INTERNACAO_AUTORIZ" Or _
		    WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ" Or _
		    WebVisionCode = "V_WEB_INTERNACAO_AUT_REEMBOLSO") Then
		    If (CommandID = "VALIDARINTERNACAO") Then
              Dim qParametroAtendimento As Object
              Set qParametroAtendimento = NewQuery

              qParametroAtendimento.Active = False
              qParametroAtendimento.Clear
              qParametroAtendimento.Add("SELECT PADRAOACOMODACAO FROM SAM_PARAMETROSATENDIMENTO")
              qParametroAtendimento.Active = True

              If ((qParametroAtendimento.FieldByName("PADRAOACOMODACAO").AsString = "P") And (HandleAut.FieldByName("NUMEROAUTORIZACAO").AsInteger > 0)) Then
				Dim qPossuiDiasInternacao As Object
				Set qPossuiDiasInternacao = NewQuery

				qPossuiDiasInternacao.Active = False
				qPossuiDiasInternacao.Clear
          		qPossuiDiasInternacao.Add("SELECT COALESCE(DIARIASLIBERADAS,0) DIARIASLIBERADAS")
          		qPossuiDiasInternacao.Add("  FROM SAM_AUTORIZ")
          		qPossuiDiasInternacao.Add(" WHERE HANDLE = :HANDLEAUTORIZACAO")
          		qPossuiDiasInternacao.ParamByName("HANDLEAUTORIZACAO").AsInteger = HandleAut.FieldByName("NUMEROAUTORIZACAO").AsInteger
          		qPossuiDiasInternacao.Active = True

				If qPossuiDiasInternacao.FieldByName("DIARIASLIBERADAS").AsInteger > 0 Then
				  Set qPossuiDiasInternacao = Nothing

                  Dim SPPreparaGeracaoDiaria As BStoredProc
  			      Set SPPreparaGeracaoDiaria = NewStoredProc
		          SPPreparaGeracaoDiaria.AutoMode = True
		          SPPreparaGeracaoDiaria.Name = "BSAut_PreparaGeracaoDiaria"

			      SPPreparaGeracaoDiaria.AddParam("P_HANDLEAUTORIZ",ptInput, ftInteger)
			      SPPreparaGeracaoDiaria.AddParam("P_USUARIO",ptInput, ftInteger)

			      SPPreparaGeracaoDiaria.ParamByName("P_HANDLEAUTORIZ").AsInteger = HandleAut.FieldByName("NUMEROAUTORIZACAO").AsInteger
			      SPPreparaGeracaoDiaria.ParamByName("P_USUARIO").AsInteger = CurrentUser

			      SPPreparaGeracaoDiaria.ExecProc

  			      Set SPPreparaGeracaoDiaria = Nothing

  			      Dim vMensagemRetorno As String
  			      Dim vResult As Integer

  			      vMensagemRetorno = ""

   			      Dim dllGerarDiaria As Object
			      Set dllGerarDiaria = CreateBennerObject("SAMAUTO.Autorizador")

			      vResult = dllGerarDiaria.GerarDiariasWEB(CurrentSystem, HandleAut.FieldByName("NUMEROAUTORIZACAO").AsInteger, 0, 0, vMensagemRetorno)

			      If vResult > 0 Then
 			        bsShowMessage(vMensagemRetorno, "I")
			      End If

  			      Set dllGerarDiaria = Nothing

  			      Dim qVinculaEventoGeradoDiariaAoProtocolo As Object
                  Set qVinculaEventoGeradoDiariaAoProtocolo = NewQuery

				  qVinculaEventoGeradoDiariaAoProtocolo.Clear
				  qVinculaEventoGeradoDiariaAoProtocolo.Add("UPDATE SAM_AUTORIZ_EVENTOGERADO SET PROTOCOLOTRANSACAO = :PROTOCOLOTRANSACAO WHERE AUTORIZACAO = :AUTORIZACAO AND PROTOCOLOTRANSACAO IS NULL AND TIPOEVENTO = :TIPOEVENTO")
				  qVinculaEventoGeradoDiariaAoProtocolo.ParamByName("PROTOCOLOTRANSACAO").AsInteger = HandleAut.FieldByName("PROTOCOLOTRANSACAO").AsInteger
				  qVinculaEventoGeradoDiariaAoProtocolo.ParamByName("AUTORIZACAO").AsInteger = HandleAut.FieldByName("NUMEROAUTORIZACAO").AsInteger
				  qVinculaEventoGeradoDiariaAoProtocolo.ParamByName("TIPOEVENTO").AsString = "D"
				  qVinculaEventoGeradoDiariaAoProtocolo.ExecSQL

                  Set qVinculaEventoGeradoDiariaAoProtocolo = Nothing



                  Dim qVinculaProtocoloCentralAtendimento As Object
                  Set qVinculaProtocoloCentralAtendimento = NewQuery

				  qVinculaProtocoloCentralAtendimento.Clear
				  qVinculaProtocoloCentralAtendimento.Add("UPDATE SAM_AUTORIZ_EVENTOGERADO SET ATENDIMENTO = (SELECT MAX(ATENDIMENTO) FROM SAM_AUTORIZ_EVENTOGERADO WHERE AUTORIZACAO = :AUTORIZACAO AND ATENDIMENTO IS NOT NULL) WHERE AUTORIZACAO = :AUTORIZACAO AND ATENDIMENTO IS NULL AND TIPOEVENTO = :TIPOEVENTO")
				  qVinculaProtocoloCentralAtendimento.ParamByName("AUTORIZACAO").AsInteger = HandleAut.FieldByName("NUMEROAUTORIZACAO").AsInteger
				  qVinculaProtocoloCentralAtendimento.ParamByName("TIPOEVENTO").AsString = "D"
				  qVinculaProtocoloCentralAtendimento.ExecSQL

                  Set qVinculaProtocoloCentralAtendimento = Nothing




  			    End If
  			  End If

  		      Set qParametroAtendimento = Nothing
		    End If
		End If

		If (WebVisionCode = "K9_W_WEB_CONSULTA_AUTORIZ" Or _
		    WebVisionCode = "K9_W_WEB_CONSULTA_CONSULTAS" Or _
		    WebVisionCode = "K9_W_WEB_CONSULTA_ELEGIBILIDADE" Or _
		    WebVisionCode = "K9_W_WEB_CONSULTA_EXECUCAO" Or _
	  	    WebVisionCode = "K9_W_WEB_CONSULTA_GUIA" Or _
		    WebVisionCode = "K9_W_WEB_SADTAUTORIZ" Or _
		    WebVisionCode = "K9_W_WEB_SADTAUTORIZPAGAMAENTO" Or _
		    WebVisionCode = "K9_W_WEB_SADTPAGAMENTO" Or _
		    WebVisionCode = "W_WEB_AUTORIZCONSULTA" Or _
		    WebVisionCode = "W_WEB_CONSULTA_AUTORIZ" Or _
		    WebVisionCode = "W_WEB_CONSULTA_CONSULTAS" Or _
		    WebVisionCode = "W_WEB_CONSULTA_ELEGIBILIDADE" Or _
		    WebVisionCode = "W_WEB_CONSULTA_GUIA" Or _
		    WebVisionCode = "W_WEB_INTERNACAO_AUTORIZ" Or _
		    WebVisionCode = "W_WEB_ODONTO_AUTORIZ" Or _
		    WebVisionCode = "W_WEB_SADTAUTORIZ" Or _
		    WebVisionCode = "W_WEB_SADTAUTORIZPAGAMAENTO" Or _
		    WebVisionCode = "W_WEB_SADTELEGIBILIDADE" Or _
		    WebVisionCode = "W_WEB_SADTPAGAMENTO" Or _
		    WebVisionCode = "W_WEB_SPSADT_AUTORIZ" Or _
		    WebVisionCode = "W_WEB_AUTORIZ_EVENTOS_AUTORIZ") Then

		  If (CommandID = "VALIDARSPSADT" Or _
			  CommandID = "VALIDARINTERNACAO" Or _
			  CommandID = "ATUALIZARSADT" Or _
			  CommandID = "VALIDARSADT" Or _
			  CommandID = "ATUALIZASADT" Or _
			  CommandID = "ATUALIZAR" Or _
			  CommandID = "VALIDARSADTELEGIBILIDADE" Or _
			  CommandID = "SADTPAGVALIDAR" Or _
			  CommandID = "VALIDARODONTO" Or _
			  CommandID = "v1" Or _
			  CommandID = "v2") Then

			Dim dllEspecifico As Object
		    Set dllEspecifico = CreateBennerObject("especifico.uespecifico")
			dllEspecifico.ATE_EnviarSmsOnClick(CurrentSystem, HandleAut.FieldByName("NUMEROAUTORIZACAO").AsInteger, 0, False)

			Set dllEspecifico = Nothing
		  End If
		End If
		Set HandleAut = Nothing
	End If
  End If
  Exit Sub

  Erro:

    InfoDescription = Err.Description
    CancelDescription = Err.Description
    CanContinue = False
    If InTransaction Then
		Rollback
	End If
End Sub

Public Sub TABLE_UpdateRequired()
  If Not CurrentQuery.FieldByName("TIPOAUTORIZACAO").IsNull And _
     (WebVisionCode = "V_WEB_CONSULTA_AUTORIZ"   Or _
      WebVisionCode = "V_WEB_INTERNACAO_AUTORIZ" Or _
      WebVisionCode = "V_WEB_SPSADT_AUTORIZ"     Or _
      WebVisionCode = "V_WEB_ODONTO_AUTORIZ") Then
    Dim qSQL       As Object
    Dim viContador As Integer
    Set qSQL = NewQuery

    qSQL.Add("SELECT HERDARSOLICITANTEDE,")
    qSQL.Add("       HERDARRECEBEDORDE,")
    qSQL.Add("       HERDAREXECUTORDE,")
    qSQL.Add("       HERDARLOCALEXECDE")
    qSQL.Add("FROM SAM_TIPOAUTORIZ")
    qSQL.Add("WHERE HANDLE = :HTIPOAUTORIZ")
    qSQL.ParamByName("HTIPOAUTORIZ").AsInteger = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
    qSQL.Active = True

    viContador = 0

    While (CurrentQuery.FieldByName("RECEBEDOR").IsNull     Or _
           CurrentQuery.FieldByName("EXECUTOR").IsNull      Or _
           CurrentQuery.FieldByName("SOLICITANTE").IsNull   Or _
           CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull) And _
          (viContador < 3)

      'Realizar herança para o campo Recebedor
      If CurrentQuery.FieldByName("RECEBEDOR").IsNull And _
         qSQL.FieldByName("HERDARRECEBEDORDE").AsString <> "N" Then
        If     qSQL.FieldByName("HERDARRECEBEDORDE").AsString = "E" And _
               Not CurrentQuery.FieldByName("EXECUTOR").IsNull Then
          CurrentQuery.FieldByName("RECEBEDOR").AsInteger = CurrentQuery.FieldByName("EXECUTOR").AsInteger
        ElseIf qSQL.FieldByName("HERDARRECEBEDORDE").AsString = "S" And _
               Not CurrentQuery.FieldByName("SOLICITANTE").IsNull Then
          CurrentQuery.FieldByName("RECEBEDOR").AsInteger = CurrentQuery.FieldByName("SOLICITANTE").AsInteger
        ElseIf qSQL.FieldByName("HERDARRECEBEDORDE").AsString = "L" And _
               Not CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull Then
          CurrentQuery.FieldByName("RECEBEDOR").AsInteger = CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger
        End If
      End If

      'Realizar herança para o campo Executor
      If CurrentQuery.FieldByName("EXECUTOR").IsNull And _
         qSQL.FieldByName("HERDAREXECUTORDE").AsString <> "N" Then
        If     qSQL.FieldByName("HERDAREXECUTORDE").AsString = "R" And _
               Not CurrentQuery.FieldByName("RECEBEDOR").IsNull Then
          CurrentQuery.FieldByName("EXECUTOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
        ElseIf qSQL.FieldByName("HERDAREXECUTORDE").AsString = "S" And _
               Not CurrentQuery.FieldByName("SOLICITANTE").IsNull Then
          CurrentQuery.FieldByName("EXECUTOR").AsInteger = CurrentQuery.FieldByName("SOLICITANTE").AsInteger
        ElseIf qSQL.FieldByName("HERDAREXECUTORDE").AsString = "L" And _
               Not CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull Then
          CurrentQuery.FieldByName("EXECUTOR").AsInteger = CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger
        End If
      End If

      'Realizar herança para o campo Solicitante
      If CurrentQuery.FieldByName("SOLICITANTE").IsNull And _
         qSQL.FieldByName("HERDARSOLICITANTEDE").AsString <> "N" Then
        If     qSQL.FieldByName("HERDAREXECUTORDE").AsString = "R" And _
               Not CurrentQuery.FieldByName("RECEBEDOR").IsNull Then
          CurrentQuery.FieldByName("SOLICITANTE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
        ElseIf qSQL.FieldByName("HERDARSOLICITANTEDE").AsString = "E" And _
               Not CurrentQuery.FieldByName("EXECUTOR").IsNull Then
          CurrentQuery.FieldByName("SOLICITANTE").AsInteger = CurrentQuery.FieldByName("EXECUTOR").AsInteger
        ElseIf qSQL.FieldByName("HERDARSOLICITANTEDE").AsString = "L" And _
               Not CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull Then
          CurrentQuery.FieldByName("SOLICITANTE").AsInteger = CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger
        End If
      End If

      'Realizar herança para o campo Local de Execução
      If CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull And _
         qSQL.FieldByName("HERDARLOCALEXECDE").AsString <> "N" Then
        If     qSQL.FieldByName("HERDAREXECUTORDE").AsString = "R" And _
               Not CurrentQuery.FieldByName("RECEBEDOR").IsNull Then
          CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
        ElseIf qSQL.FieldByName("HERDARLOCALEXECDE").AsString = "E" And _
               Not CurrentQuery.FieldByName("SOLICITANTE").IsNull Then
          CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger = CurrentQuery.FieldByName("EXECUTOR").AsInteger
        ElseIf qSQL.FieldByName("HERDARLOCALEXECDE").AsString = "S" And _
               Not CurrentQuery.FieldByName("SOLICITANTE").IsNull Then
          CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger = CurrentQuery.FieldByName("SOLICITANTE").AsInteger
        End If
      End If

      viContador = viContador + 1
    Wend

    Set qSQL = Nothing
  End If
End Sub

Public Sub CancelarInternacao
  Dim sql As Object
  Set sql = NewQuery

  sql.Active = False
  sql.Clear
  sql.Add("UPDATE WEB_AUTORIZ_EVENTOS      ")
  sql.Add("   SET SITUACAOEVENTO = '3'     ")
  sql.Add(" WHERE WEBAUTORIZ = :WEBAUTORIZ ")
  sql.ParamByName("WEBAUTORIZ").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL

  sql.Active = False
  sql.Clear
  sql.Add("UPDATE WEB_AUTORIZ      ")
  sql.Add("   SET SITUACAO = 'C'   ")
  sql.Add(" WHERE HANDLE = :HANDLE ")
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL

  Set sql = Nothing

  bsShowMessage("Solicitação cancelada com sucesso!", "I")
End Sub
