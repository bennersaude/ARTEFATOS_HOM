'HASH: 754B0C48345FDD081BB6E531374607EA
'#Uses "*Biblioteca"

Option Explicit


Public Sub Main

  Dim qSolicitacao 			            As Object
  Dim qParametrosWeb 			        As Object
  Dim qBuscaTipoUsuario					As Object
  Dim vHandleSolicitacao 		        As Long
  Dim vHandleRelatorio   				As Long
  Dim vDiretorioTemporarioServidor 	    As String
  Dim vCaminhoArquivo           	    As String
  Dim vLinkExternoPortal 		        As String
  Dim vEmailRemetente			        As String
  Dim qUsuario					        As Object
  Dim vUltimoCaractere                  As String
  Dim qUpdateSol						As Object
  Dim vOcorrencias						As String
  Dim vSucess							As Boolean
  Dim vExiste							As String
  Dim vTipo								As String

  Dim rel As CSReportPrinter

  Set qSolicitacao = NewQuery
  Set qParametrosWeb = NewQuery
  Set qUsuario = NewQuery
  Set qBuscaTipoUsuario = NewQuery
  Set qUpdateSol = NewQuery

  vTipo = SessionVar("TIPO")
  'SELECIONANDO PARAMETROS GERAIS DA WEB
  qParametrosWeb.Active = False
  qParametrosWeb.Clear
  qParametrosWeb.Add("SELECT LINKEXTERNOPORTAL,            ")
  qParametrosWeb.Add("		DIRETORIOTEMPORARIOSERVIDOR,   ")
  qParametrosWeb.Add("		EMAILREMETENTE,                ")
  qParametrosWeb.Add("		RELATPRESTADORPJANALITICO,     ")
  qParametrosWeb.Add("		RELATPRESTADORPJSINTETICO,     ")
  qParametrosWeb.Add("		RELATPRESTADORPFANALITICO,     ")
  qParametrosWeb.Add("		RELATPRESTADORPFSINTETICO,     ")
  qParametrosWeb.Add("		RELATPRESTADORSINTETICOODONTO  ")
  qParametrosWeb.Add("FROM SAM_PARAMETROSWEB               ")
  qParametrosWeb.Active = True

  vLinkExternoPortal = qParametrosWeb.FieldByName("LINKEXTERNOPORTAL").AsString
  vDiretorioTemporarioServidor = qParametrosWeb.FieldByName("DIRETORIOTEMPORARIOSERVIDOR").AsString
  vEmailRemetente = qParametrosWeb.FieldByName("EMAILREMETENTE").AsString

  'VERIFICANDO SE O ULTIMO CARACTERE DO CAMINHO FISICO É "\"
  vUltimoCaractere = Mid(vDiretorioTemporarioServidor,Len(vDiretorioTemporarioServidor))
  If vUltimoCaractere <> "\" Then
    vDiretorioTemporarioServidor = vDiretorioTemporarioServidor + "\"
  End If

  'Buscando dados da solicitação, Email do Usuario Logado, Nome do Prestador
  qSolicitacao.Active = False
  qSolicitacao.Clear
  qSolicitacao.Add("SELECT SOL.HANDLE,")
  qSolicitacao.Add(" SOL.DATAINICIAL, ")
  qSolicitacao.Add(" SOL.DATAFINAL,   ")
  qSolicitacao.Add(" SOL.PRESTADOR,   ")
  qSolicitacao.Add(" SOL.TIPOSOLICITACAO,")
  qSolicitacao.Add(" SOL.OCORRENCIAS,")
  qSolicitacao.Add(" PRE.EMAIL,       ")
  qSolicitacao.Add(" PRE.NOME         ")
  qSolicitacao.Add("FROM SAM_PRESTADOR_SOLICITACOESWEB SOL")
  qSolicitacao.Add("JOIN SAM_PRESTADOR PRE ON PRE.HANDLE = SOL.PRESTADOR")

  If vTipo = "SOLICITACAO" Then
    qSolicitacao.Add("WHERE SOL.SITUACAO = '4' AND SOL.HANDLE = :SOLICITACAO AND SOL.ORIGEM = '1'")
  Else
    qSolicitacao.Add("WHERE SOL.SITUACAO = '1' AND SOL.PRESTADOR = :SOLICITACAO AND SOL.ORIGEM = '1'")
  End If

  qSolicitacao.ParamByName("SOLICITACAO").AsInteger = CLng(SessionVar("HANDLESOLICIT")) 'parametro da solicitacao
  qSolicitacao.Active = True

  qUpdateSol.Clear
  qUpdateSol.Add("UPDATE SAM_PRESTADOR_SOLICITACOESWEB SET SITUACAO = :SITUACAO, OCORRENCIAS = :OCORRENCIAS WHERE HANDLE = :SOLICITACAO")

  qBuscaTipoUsuario.Clear
  qBuscaTipoUsuario.Add("SELECT FISICAJURIDICA            ")
  qBuscaTipoUsuario.Add("  FROM SAM_PRESTADOR             ")
  qBuscaTipoUsuario.Add("WHERE HANDLE=:HPRESTADOR         ")

  While Not qSolicitacao.EOF
	vHandleSolicitacao = qSolicitacao.FieldByName("HANDLE").AsInteger

	On Error GoTo ErroProcesso
	  vSucess = False

	  vHandleRelatorio = 0
	  vOcorrencias = ""
	  vOcorrencias = qSolicitacao.FieldByName("OCORRENCIAS").AsString + vbNewLine + "Início do processamento de geração do demonstrativo." + vbNewLine + "Data/Hora: " + Format(ServerNow(), "DD/mm/YYYY h:mm:ss") + vbNewLine

	  qUpdateSol.Active = False
	  qUpdateSol.ParamByName("SITUACAO").AsString = "2"
	  qUpdateSol.ParamByName("SOLICITACAO").AsInteger = vHandleSolicitacao
	  qUpdateSol.ParamByName("OCORRENCIAS").AsString = vOcorrencias
	  qUpdateSol.ExecSQL

	  'passando SessionVar para o relatório/demonstrativos
	  SessionVar("DATAINICIAL") = Format(qSolicitacao.FieldByName("DATAINICIAL").AsString, "yyyy-mm-dd")
	  SessionVar("DATAFINAL") = Format(qSolicitacao.FieldByName("DATAFINAL").AsString, "yyyy-mm-dd")
	  SessionVar("PRESTADOR") = qSolicitacao.FieldByName("PRESTADOR").AsInteger

	  qBuscaTipoUsuario.Active = False
	  qBuscaTipoUsuario.ParamByName("HPRESTADOR").AsInteger = qSolicitacao.FieldByName("PRESTADOR").AsInteger
	  qBuscaTipoUsuario.Active = True

	  If (qSolicitacao.FieldByName("TIPOSOLICITACAO").AsString = 1) Then 'Demonstrativo de Pagamento Médico(Sintético)
	  	If (qBuscaTipoUsuario.FieldByName("FISICAJURIDICA").AsInteger = 2) Then  'Se for jurídico
			vHandleRelatorio = qParametrosWeb.FieldByName("RELATPRESTADORPJSINTETICO").AsInteger
		End If
		If (qBuscaTipoUsuario.FieldByName("FISICAJURIDICA").AsInteger = 1) Then 'Se for físico
			vHandleRelatorio = qParametrosWeb.FieldByName("RELATPRESTADORPfSINTETICO").AsInteger
		End If
	  End If

	  If (qSolicitacao.FieldByName("TIPOSOLICITACAO").AsString = 2) Then 'Demonstrativo de Analise de conta médica (Analítico)
	  	If (qBuscaTipoUsuario.FieldByName("FISICAJURIDICA").AsInteger = 2) Then   'Se for jurídico
			vHandleRelatorio = qParametrosWeb.FieldByName("RELATPRESTADORPJANALITICO").AsInteger
		End If
		If (qBuscaTipoUsuario.FieldByName("FISICAJURIDICA").AsInteger = 1) Then 'Se for físico
			vHandleRelatorio = qParametrosWeb.FieldByName("RELATPRESTADORPFANALITICO").AsInteger
		End If
	  End If

	  If (qSolicitacao.FieldByName("TIPOSOLICITACAO").AsString = 3) Then 'Demonstrativo de Pagamento Odontológico(Sintético)
	    vHandleRelatorio = qParametrosWeb.FieldByName("RELATPRESTADORSINTETICOODONTO").AsInteger
	  End If

	  Select Case qSolicitacao.FieldByName("TIPOSOLICITACAO").AsString
	  Case "1"
	  	vCaminhoArquivo = vDiretorioTemporarioServidor + "demonstrativoPagamento_" + CStr(vHandleSolicitacao) +  ".pdf"
	  Case "2"
	    vCaminhoArquivo = vDiretorioTemporarioServidor + "demonstrativoAnaliseContaMedica_" + CStr(vHandleSolicitacao) +  ".pdf"
	  Case "3"
	  	vCaminhoArquivo = vDiretorioTemporarioServidor + "demonstrativoPagamentoOdontologico_" + CStr(vHandleSolicitacao) +  ".pdf"
	  End Select


	  Dim qRelatorio As Object
	  Set qRelatorio = NewQuery

	  qRelatorio.Clear
	  qRelatorio.Add("SELECT TIPO FROM R_RELATORIOS WHERE HANDLE = :PHANDLE")
	  qRelatorio.ParamByName("PHANDLE").AsInteger = vHandleRelatorio
	  qRelatorio.Active = True

	  Set rel = NewReport(vHandleRelatorio)

	  If qRelatorio.FieldByName("TIPO").AsInteger = 2 Then

	  		If InStr(SQLServer,"SQL") > 0 Then

	  			Dim vDataInicial As String
	  			Dim vDataFinal   As String

	  			vDataInicial = FormatDateTime2("YYYY-MM-DD",qSolicitacao.FieldByName("DATAINICIAL").AsDateTime)
	  			vDataFinal   = FormatDateTime2("YYYY-MM-DD",qSolicitacao.FieldByName("DATAFINAL").AsDateTime)

				Call CriarFiltro("DEM-PG4", CurrentUser, "DATAINICIAL='" & vDataInicial & "'," & _
				                                         "DATAFINAL  ='" & vDataFinal & "'," & _
				                                         "PRESTADOR=" & qSolicitacao.FieldByName("PRESTADOR").AsString, _
				                                         "DATAINICIAL, DATAFINAL, PRESTADOR", _
				                                         "'" & vDataInicial & "','" & vDataFinal & "'," & _
				                                         qSolicitacao.FieldByName("PRESTADOR").AsString)
			Else
		          Call CriarFiltro("DEM-PG4", CurrentUser, "DATAINICIAL ='" & qSolicitacao.FieldByName("DATAINICIAL").AsDateTime & "',"  & "DATAFINAL='" & qSolicitacao.FieldByName("DATAFINAL").AsString & "'," & "PRESTADOR=" & qSolicitacao.FieldByName("PRESTADOR").AsString, "DATAINICIAL, DATAFINAL, PRESTADOR", "'" & qSolicitacao.FieldByName("DATAINICIAL").AsDateTime & "','" & qSolicitacao.FieldByName("DATAFINAL").AsString & "'," & qSolicitacao.FieldByName("PRESTADOR").AsString)
		    End If

	  End If

      rel.ExportToFile(vCaminhoArquivo)

	  Set qRelatorio = Nothing
	  Set rel = Nothing

	  vExiste = Dir(vCaminhoArquivo)
	  If (vExiste <> "") And (Not IsNull(vExiste)) And (Len(vExiste) > 0) Then
	  	SetFieldDocument("SAM_PRESTADOR_SOLICITACOESWEB","ARQUIVODEMONSTRATIVO",vHandleSolicitacao,vCaminhoArquivo,True)
	  	Kill(vCaminhoArquivo)

	  	vOcorrencias = vOcorrencias + vbNewLine + "Fim do processamento de geração do demonstrativo." + vbNewLine + "Data/Hora: " + Format(ServerNow(), "DD/mm/YYYY h:mm:ss") + vbNewLine
	  Else
		vOcorrencias = vOcorrencias + vbNewLine + "Não existem dados a serem exportados no relatório. Fim do processamento de geração do demonstrativo." + vbNewLine + "Data/Hora: " + Format(ServerNow(), "DD/mm/YYYY h:mm:ss") + vbNewLine
	  End If

	  qUpdateSol.Active = False
	  qUpdateSol.ParamByName("SITUACAO").AsString = "3"
	  qUpdateSol.ParamByName("SOLICITACAO").AsInteger = vHandleSolicitacao
	  qUpdateSol.ParamByName("OCORRENCIAS").AsString = vOcorrencias
	  qUpdateSol.ExecSQL

	  vSucess = True

    ErroProcesso:
	  If Not vSucess Then
	  	vOcorrencias = vOcorrencias + vbNewLine + "Erro durante geração de demonstrativo: " + Err.Description + vbNewLine + "Data/Hora: " + Format(ServerNow(), "DD/mm/YYYY h:mm:ss") + vbNewLine
	  	qUpdateSol.Active = False
	  	qUpdateSol.ParamByName("SITUACAO").AsString = "4"
	  	qUpdateSol.ParamByName("SOLICITACAO").AsInteger = vHandleSolicitacao
	  	qUpdateSol.ParamByName("OCORRENCIAS").AsString = vOcorrencias
	  	qUpdateSol.ExecSQL
	  End If

    If qSolicitacao.FieldByName("EMAIL").AsString <> "" And vEmailRemetente <> "" Then
      On Error GoTo ErroMail
	    Dim MailObj As Object

		Set MailObj = NewMail

 		MailObj.From    = vEmailRemetente
		MailObj.SendTo  =  qSolicitacao.FieldByName("EMAIL").AsString
        MailObj.ContentType = "text/html"

    	Select Case qSolicitacao.FieldByName("TIPOSOLICITACAO").AsString
		Case "1"
			MailObj.Subject = "Envio de demonstrativo de pagamento - Período de " + Format(qSolicitacao.FieldByName("DATAINICIAL").AsString, "DD/mm/YYYY") + " a " + Format(qSolicitacao.FieldByName("DATAFINAL").AsString, "DD/mm/YYYY")
		Case "2"
			MailObj.Subject = "Envio de demonstrativo de Analise de Conta Médica - Período de " + Format(qSolicitacao.FieldByName("DATAINICIAL").AsString, "DD/mm/YYYY") + " a " + Format(qSolicitacao.FieldByName("DATAFINAL").AsString, "DD/mm/YYYY")
		Case "3"
			MailObj.Subject = "Envio de demonstrativo de pagamento Odontológico - Período de " + Format(qSolicitacao.FieldByName("DATAINICIAL").AsString, "DD/mm/YYYY") + " a " + Format(qSolicitacao.FieldByName("DATAFINAL").AsString, "DD/mm/YYYY")
		End Select

  		MailObj.Priority = 0
    	MailObj.Text.Clear
    	MailObj.Text.Add("<html>")
        MailObj.Text.Add("<body>")
        MailObj.Text.Add("<p>")
	    MailObj.Text.Add("Caro <b> " + qSolicitacao.FieldByName("NOME").AsString + ",</b>")
		MailObj.Text.Add("<br>")

		Select Case qSolicitacao.FieldByName("TIPOSOLICITACAO").AsString
		Case "1"
	        MailObj.Text.Add("O demonstrativo de Pagamento referente ao período de " + Format(qSolicitacao.FieldByName("DATAINICIAL").AsString, "DD/mm/YYYY") + " a " + Format(qSolicitacao.FieldByName("DATAFINAL").AsString, "DD/mm/YYYY"))
		Case "2"
	    	MailObj.Text.Add("O demonstrativo de Análise de Conta Médica referente ao período de " + Format(qSolicitacao.FieldByName("DATAINICIAL").AsString, "DD/mm/YYYY") + " a " + Format(qSolicitacao.FieldByName("DATAFINAL").AsString, "DD/mm/YYYY"))
	    Case "3"
	    	MailObj.Text.Add("O demonstrativo de Análise Pagamento Odontológico referente ao período de " + Format(qSolicitacao.FieldByName("DATAINICIAL").AsString, "DD/mm/YYYY") + " a " + Format(qSolicitacao.FieldByName("DATAFINAL").AsString, "DD/mm/YYYY"))
		End Select
        MailObj.Text.Add(" foi gerado e pode ser acessado através do <a href='" + vLinkExternoPortal +  "'>Portal de Serviços</a>.")
		MailObj.Text.Add("<br>")
    	MailObj.Text.Add("</p>")
        MailObj.Text.Add("</body>")
        MailObj.Text.Add("</html>")
        MailObj.Send 'Manda email

	  ErroMail:
		Set MailObj = Nothing
    End If

    qSolicitacao.Next
  Wend

  Set qSolicitacao = Nothing
  Set qParametrosWeb = Nothing
  Set qBuscaTipoUsuario = Nothing
  Set qUpdateSol = Nothing
End Sub
