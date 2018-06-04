'HASH: ECE55B698D338C4CF315B159501C5414
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
	Dim descricaoAnexo As String
	Dim handleRelatorio As Long
	Dim rep As CSReportPrinter

	Dim qParametrosAtendimento As Object
	Set qParametrosAtendimento = NewQuery

	Dim qAnexosProtocolo As Object
	Set qAnexosProtocolo = NewQuery

	qAnexosProtocolo.Add("SELECT COUNT(1) CONTADOR ")

	qParametrosAtendimento.Add("SELECT PROTRELATORIOQUIMIO,      ")
	qParametrosAtendimento.Add("       PROTRELATORIORADIO,       ")
	qParametrosAtendimento.Add("       PROTRELATORIOOPME,        ")
	qParametrosAtendimento.Add("       PROTRELATORIOINTERMEDIACAOOPME")
	qParametrosAtendimento.Add("  FROM SAM_PARAMETROSATENDIMENTO ")

	qParametrosAtendimento.Active = True

	Select Case CurrentQuery.FieldByName("TIPO").AsString
		Case "Q"
			handleRelatorio = qParametrosAtendimento.FieldByName("PROTRELATORIOQUIMIO").AsInteger
			qAnexosProtocolo.Add(" FROM SAM_AUTORIZ_ANEXOQUIMIO ")
			descricaoAnexo = "Quimioterapia"
		Case "R"
			handleRelatorio = qParametrosAtendimento.FieldByName("PROTRELATORIORADIO").AsInteger
			qAnexosProtocolo.Add(" FROM SAM_AUTORIZ_ANEXORADIO  ")
			descricaoAnexo = "Radioterapia"
		Case "O"
        	Dim qAnexoOPME As BPesquisa
        	Set qAnexoOPME = NewQuery

        	qAnexoOPME.Add("SELECT SUM(CASE WHEN ANEX.INTERMEDIACAOCOMPRA = 'S' THEN 0 ELSE 1 END) QTDNAOINTERMEDIADOS,")
        	qAnexoOPME.Add("       SUM(CASE WHEN ANEX.INTERMEDIACAOCOMPRA = 'N' THEN 0 ELSE 1 END) QTDINTERMEDIADOS")
        	qAnexoOPME.Add("FROM SAM_AUTORIZ_ANEXOOPME ANEX")
        	qAnexoOPME.Add("WHERE ANEX.PROTOCOLOTRANSACAO = :HPROTOCOLO")

        	qAnexoOPME.ParamByName("HPROTOCOLO").AsInteger = CLng(SessionVar("PROTOCOLOTRANSACAO"))
        	qAnexoOPME.Active = True

        	If (qAnexoOPME.FieldByName("QTDNAOINTERMEDIADOS").AsInteger = 0) And _
               (qAnexoOPME.FieldByName("QTDINTERMEDIADOS").AsInteger > 0) Then
              handleRelatorio = qParametrosAtendimento.FieldByName("PROTRELATORIOINTERMEDIACAOOPME").AsInteger
        	Else
			  handleRelatorio = qParametrosAtendimento.FieldByName("PROTRELATORIOOPME").AsInteger
        	End If

        	Set qAnexoOPME = Nothing

			qAnexosProtocolo.Add(" FROM SAM_AUTORIZ_ANEXOOPME   ")
			descricaoAnexo = "OPME"
		Case Else
			bsShowMessage("Necessário preencher um tipo de anexo!", "I")
	End Select

    If (handleRelatorio > 0) Then
		qAnexosProtocolo.Add(" WHERE PROTOCOLOTRANSACAO = :PROTOCOLOTRANSACAO ")
		qAnexosProtocolo.Add("   AND AUTORIZACAO = :AUTORIZACAO               ")

		qAnexosProtocolo.ParamByName("PROTOCOLOTRANSACAO").AsString = SessionVar("PROTOCOLOTRANSACAO")
		qAnexosProtocolo.ParamByName("AUTORIZACAO").AsString = SessionVar("AUTORIZACAO")

		qAnexosProtocolo.Active = True

	    If qAnexosProtocolo.FieldByName("CONTADOR").AsInteger <= 0 Then
			bsShowMessage("Protocolo sem anexos de " + descricaoAnexo + " a serem impressos", "I")
			Set qAnexosProtocolo = Nothing
			Set qParametrosAtendimento = Nothing

			Exit Sub
		End If

        If (WebMode) Then
           UserVar("HandleFiltro") = ""
        End If

		Set rep = NewReport(handleRelatorio)
		rep.Preview
		Set rep = Nothing
    Else
        bsShowMessage("Relatório não encontrado", "I")
    End If

	Set qAnexosProtocolo = Nothing
	Set qParametrosAtendimento = Nothing
End Sub
