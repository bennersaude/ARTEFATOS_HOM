'HASH: 503790A9F2C580C32053022C95773F6E

Public Sub Main

  Dim qBuscaRotina              As BPesquisa
  Dim QueryBuscaHandleRelatorio As BPesquisa
  Dim qBuscaDiretorio           As BPesquisa
  Dim qBuscaOcorrencias         As BPesquisa
  Dim qOcorrencias              As Object
  Dim RelatorioHandle           As Long
  Dim vsOcorrencias             As String
  Dim vsDiretorio               As String

  Set qBuscaRotina  = NewQuery
  Set QueryBuscaHandleRelatorio = NewQuery
  Set qOcorrencias = NewQuery
  Set qBuscaDiretorio = NewQuery
  Set qBuscaOcorrencias = NewQuery

  QueryBuscaHandleRelatorio.Active = False
  QueryBuscaHandleRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'BEN202'")
  QueryBuscaHandleRelatorio.Active = True

  RelatorioHandle = QueryBuscaHandleRelatorio.FieldByName("HANDLE").AsInteger

  qBuscaRotina.Active = False
  qBuscaRotina.Add(" SELECT DISTINCT A.TITULAR,                                     ")
  qBuscaRotina.Add("                 B.EMAIL,                                       ")
  qBuscaRotina.Add("                 B.NOME                                         ")
  qBuscaRotina.Add("   FROM SAM_ROTINACANCELAMENTO_BENEF A                          ")
  qBuscaRotina.Add("   JOIN SAM_BENEFICIARIO             B ON A.TITULAR = B.HANDLE  ")
  qBuscaRotina.Add("  WHERE A.CANCELAMENTO = :HANDLE                                ")

  qBuscaRotina.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HANDLEROTINA"))
  qBuscaRotina.Active  = True


  qBuscaDiretorio.Active = False
  qBuscaDiretorio.Add(" SELECT DIRETORIOTEMPORARIOSERVIDOR ")
  qBuscaDiretorio.Add("   FROM SAM_PARAMETROSWEB           ")
  qBuscaDiretorio.Active = True

  qBuscaOcorrencias.Active = False
  qBuscaOcorrencias.Add("  SELECT OCORRENCIAS              ")
  qBuscaOcorrencias.Add("     FROM SAM_ROTINACANCELAMENTO  ")
  qBuscaOcorrencias.Add("    WHERE HANDLE = :HANDLE        ")
  qBuscaOcorrencias.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HANDLEROTINA"))
  qBuscaOcorrencias.Active = True


  vsDiretorio = qBuscaDiretorio.FieldByName("DIRETORIOTEMPORARIOSERVIDOR").AsString

  vsOcorrencias = qBuscaOcorrencias.FieldByName("OCORRENCIAS").AsString + _
                      "===============================================" + _
                                                                Chr(13) + _
                                      "Ocorrências do envio de e-mails" + _
                                                                Chr(13) + _
                      "===============================================" + _
                                                                Chr(13) + _
                                          Str(ServerDate) + " - Inicio" + _
                                                                Chr(13) + _
                                                                Chr(13)


  While Not qBuscaRotina.EOF

    If qBuscaRotina.FieldByName("EMAIL").AsString <> "" Then
	  SessionVar("HANDLE_TITULAR") = qBuscaRotina.FieldByName("TITULAR").AsString
	  ReportExport(RelatorioHandle, "A.HANDLE = " + SessionVar("HANDLEROTINA") ,vsDiretorio + "\" + qBuscaRotina.FieldByName("NOME").AsString + ".pdf", False, False, qBuscaRotina.FieldByName("EMAIL").AsString, "Aviso de Cancelamento")
      vsOcorrencias = vsOcorrencias + qBuscaRotina.FieldByName("NOME").AsString + " " + qBuscaRotina.FieldByName("EMAIL").AsString + " " + Str(Time) + Chr(13)
	  qBuscaRotina.Next
	Else
	  vsOcorrencias = vsOcorrencias + qBuscaRotina.FieldByName("NOME").AsString + " - não possui e-mail cadastrado!" + " " + Str(Time) +  Chr(13)
	  qBuscaRotina.Next
    End If

  Wend


  vsOcorrencias = vsOcorrencias + Chr(13) + Str(ServerDate) + " - Fim" + Chr(13)

  qOcorrencias.Add("  UPDATE SAM_ROTINACANCELAMENTO                    ")
  qOcorrencias.Add("     SET OCORRENCIAS = :OCORRENCIAS                ")
  qOcorrencias.Add("   WHERE HANDLE = :HANDLE                          ")
  qOcorrencias.ParamByName("OCORRENCIAS").AsString = vsOcorrencias
  qOcorrencias.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HANDLEROTINA"))
  qOcorrencias.ExecSQL

  Set QueryBuscaHandleRelatorio = Nothing
  Set qBuscaRotina = Nothing
  Set qOcorrencias = Nothing
  Set qBuscaDiretorio = Nothing
  Set qBuscaOcorrencias = Nothing

End Sub
