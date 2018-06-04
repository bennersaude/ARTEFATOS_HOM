'HASH: 0AA942FFFA5044A09BE9A710DA622053
Option Explicit
Public Sub Main

  Dim q_Solicitacao As BPesquisa
  Dim q_UpdateSolicitacao As BPesquisa
  Dim q_ParametrosPortalServico As BPesquisa
  Dim q_ParametrosWeb As BPesquisa
  Dim q_TipoPrestador As BPesquisa
  Dim q_TabelaGeralEventoInicial As BPesquisa
  Dim q_TabelaGeralEventoFinal As BPesquisa
  Dim r_Relatorio As CSReportPrinter

  Dim v_HandleSolicitacao As Long
  Dim v_HandlePrestador As Long
  Dim v_HandleConvenio As Long
  Dim v_HandleMascaraTge As Long
  Dim v_HandleEventoInicial As Long
  Dim v_HandleEventoFinal As Long
  Dim v_DataReferencia As Date

  Dim v_DirTempServidor As String
  Dim v_UltimoCaractere As String
  Dim v_CaminhoArquivo As String
  Dim v_HandleRelatorio As Long
  Dim v_Existe As String
  Dim v_Sucesso As Boolean

  Set q_Solicitacao = NewQuery
  Set q_UpdateSolicitacao = NewQuery
  Set q_TipoPrestador = NewQuery
  Set q_TabelaGeralEventoInicial = NewQuery
  Set q_TabelaGeralEventoFinal = NewQuery
  Set q_ParametrosPortalServico = NewQuery
  Set q_ParametrosWeb = NewQuery

  q_Solicitacao.Active = False
  q_Solicitacao.Clear

  If InStr(SQLServer, "SQL") > 0 Then
    q_Solicitacao.Add("SELECT TOP 1 HANDLE,        ")
  Else
   If InStr(SQLServer, "ORA") > 0 Then
    q_Solicitacao.Add(" SELECT * FROM (            ")
  	q_Solicitacao.Add("SELECT HANDLE,         	   ")
   End If
  End If
  q_Solicitacao.Add("       PRESTADOR,             ")
  q_Solicitacao.Add("       CONVENIO,              ")
  q_Solicitacao.Add("       MASCARATGE,            ")
  q_Solicitacao.Add("       DATAEMISSAO,           ")
  q_Solicitacao.Add("       SITUACAO,              ")
  q_Solicitacao.Add("       ARQUIVO,               ")
  q_Solicitacao.Add("       DATAREFERENCIA	       ")
  q_Solicitacao.Add("  FROM POR_RELTABPRECOEVENTO  ")
  q_Solicitacao.Add(" WHERE UPPER(PROCESSAR) = 'S' ")
  If SessionVar("CONSULTATABELAPRECO_HANDLESOLICITACAO") <> "" Then
    q_Solicitacao.Add(" AND HANDLE = :HANDLE ")
    q_Solicitacao.ParamByName("HANDLE").AsInteger = CInt(SessionVar("CONSULTATABELAPRECO_HANDLESOLICITACAO"))
  End If
  q_Solicitacao.Add(" ORDER BY DATAEMISSAO DESC   ")

  If InStr(SQLServer, "ORA") > 0 Then
    q_Solicitacao.Add(" ) X  WHERE ROWNUM <= 1     ")
  End If

  q_Solicitacao.Active = True

  If (q_Solicitacao.FieldByName("HANDLE").AsInteger > 0) Then

  If SessionVar("CONSULTATABELAPRECO_HANDLESOLICITACAO") <> "" Then
    v_HandleSolicitacao = CInt(SessionVar("CONSULTATABELAPRECO_HANDLESOLICITACAO"))
  Else
    v_HandleSolicitacao = q_Solicitacao.FieldByName("HANDLE").AsInteger
  End If

  v_HandlePrestador   = q_Solicitacao.FieldByName("PRESTADOR").AsInteger
  v_HandleConvenio    = q_Solicitacao.FieldByName("CONVENIO").AsInteger
  v_HandleMascaraTge  = q_Solicitacao.FieldByName("MASCARATGE").AsInteger
  v_DataReferencia    = q_Solicitacao.FieldByName("DATAREFERENCIA").AsDateTime

  q_TipoPrestador.Active = False
  q_TipoPrestador.Clear
  q_TipoPrestador.Add(" SELECT TP.EVENTOINICIAL,   									")
  q_TipoPrestador.Add("        TP.EVENTOFINAL,										")
  q_TipoPrestador.Add("	       TP.HANDLE HANDLE										")
  q_TipoPrestador.Add("   FROM SAM_PRESTADOR P										")
  q_TipoPrestador.Add("   JOIN SAM_TIPOPRESTADOR TP ON P.TIPOPRESTADOR = TP.HANDLE  ")
  q_TipoPrestador.Add("  WHERE P.HANDLE = :PREST")
  q_TipoPrestador.ParamByName("PREST").AsInteger = v_HandlePrestador
  q_TipoPrestador.Active = True

  q_UpdateSolicitacao.Active = False
  q_UpdateSolicitacao.Clear
  q_UpdateSolicitacao.Add("UPDATE POR_RELTABPRECOEVENTO SET SITUACAO = :SITUACAOSOLICITACAO, OCORRENCIA = :OCORRENCIA, PROCESSAR = :PROCESSAR")
  q_UpdateSolicitacao.Add(" WHERE HANDLE = :HANDLESOLICITACAO")
  q_UpdateSolicitacao.ParamByName("SITUACAOSOLICITACAO").AsString = "G"
  q_UpdateSolicitacao.ParamByName("HANDLESOLICITACAO").AsInteger = v_HandleSolicitacao
  q_UpdateSolicitacao.ParamByName("OCORRENCIA").AsString = "Iniciada a solicitação da geração do relatório. Favor, aguarde até o termino do processo."
  q_UpdateSolicitacao.ParamByName("PROCESSAR").AsString = "N"
  q_UpdateSolicitacao.ExecSQL

  q_TabelaGeralEventoInicial.Active = False
  q_TabelaGeralEventoInicial.Clear
  q_TabelaGeralEventoInicial.Add("SELECT HANDLE                   ")
  q_TabelaGeralEventoInicial.Add("  FROM SAM_TGE                  ")
  q_TabelaGeralEventoInicial.Add(" WHERE MASCARATGE = :MASCARATGE ")
  q_TabelaGeralEventoInicial.Add("   AND ULTIMONIVEL = 'S'        ")
  q_TabelaGeralEventoInicial.Add(" ORDER BY ESTRUTURA ASC         ")
  q_TabelaGeralEventoInicial.ParamByName("MASCARATGE").AsInteger = v_HandleMascaraTge
  q_TabelaGeralEventoInicial.Active = True

  q_TabelaGeralEventoFinal.Active = False
  q_TabelaGeralEventoFinal.Clear
  q_TabelaGeralEventoFinal.Add("SELECT HANDLE                   ")
  q_TabelaGeralEventoFinal.Add("  FROM SAM_TGE                  ")
  q_TabelaGeralEventoFinal.Add(" WHERE MASCARATGE = :MASCARATGE ")
  q_TabelaGeralEventoFinal.Add("   AND ULTIMONIVEL = 'S'        ")
  q_TabelaGeralEventoFinal.Add(" ORDER BY ESTRUTURA DESC        ")
  q_TabelaGeralEventoFinal.ParamByName("MASCARATGE").AsInteger = v_HandleMascaraTge
  q_TabelaGeralEventoFinal.Active = True

  If ((q_TipoPrestador.FieldByName("EVENTOINICIAL").AsInteger) <> 0) Then
    v_HandleEventoInicial = q_TipoPrestador.FieldByName("EVENTOINICIAL").AsInteger
  Else
    v_HandleEventoInicial = q_TabelaGeralEventoInicial.FieldByName("HANDLE").AsInteger
  End If

  If ((q_TipoPrestador.FieldByName("EVENTOFINAL").AsInteger) <> 0) Then
    v_HandleEventoFinal = q_TipoPrestador.FieldByName("EVENTOFINAL").AsInteger
  Else
    v_HandleEventoFinal = q_TabelaGeralEventoFinal.FieldByName("HANDLE").AsInteger
  End If

  q_ParametrosWeb.Active = False
  q_ParametrosWeb.Clear
  q_ParametrosWeb.Add("SELECT DIRETORIOTEMPORARIOSERVIDOR ")
  q_ParametrosWeb.Add("  FROM SAM_PARAMETROSWEB           ")
  q_ParametrosWeb.Active = True

  v_DirTempServidor = q_ParametrosWeb.FieldByName("DIRETORIOTEMPORARIOSERVIDOR").AsString
  v_UltimoCaractere = Mid(v_DirTempServidor,Len(v_DirTempServidor))
  If v_UltimoCaractere <> "\" Then
    v_DirTempServidor = v_DirTempServidor + "\"
  End If

  v_CaminhoArquivo = v_DirTempServidor + "ConsultaTabelaPrecoEvento_" + CStr(v_HandleSolicitacao) + ".pdf"

  On Error GoTo ErroProcesso

	v_Sucesso = False

	q_ParametrosPortalServico.Active = False
	q_ParametrosPortalServico.Clear
	q_ParametrosPortalServico.Add("SELECT RELATORIOTABPRECOEVENTO ")
	q_ParametrosPortalServico.Add("  FROM POR_CONFIGPORTAL        ")
	q_ParametrosPortalServico.Active = True

	v_HandleRelatorio = q_ParametrosPortalServico.FieldByName("RELATORIOTABPRECOEVENTO").AsInteger

	SessionVar("EMISSAOCONSULTATABELAPRECOPROCESSO") = "S"
	SessionVar("CONSULTATABELAPRECO_HANDLEPRESTADOR") = CStr (v_HandlePrestador)
	SessionVar("CONSULTATABELAPRECO_HANDLECONVENIO") = CStr (v_HandleConvenio)
	SessionVar("CONSULTATABELAPRECO_HANDLEMASCARATGE") = CStr (v_HandleMascaraTge)
	SessionVar("CONSULTATABELAPRECO_HANDLEEVENTO_I") = CStr (v_HandleEventoInicial)
	SessionVar("CONSULTATABELAPRECO_HANDLEEVENTO_F") = CStr (v_HandleEventoFinal)
	SessionVar("CONSULTATABELAPRECO_DATAREFERENCIA") = CStr (v_DataReferencia)

  	Set r_Relatorio = NewReport(v_HandleRelatorio)
  	r_Relatorio.ExportToFile(v_CaminhoArquivo)

  	v_Existe = Dir(v_CaminhoArquivo)
  	If (v_Existe <> "") Then
		SetFieldDocument("POR_RELTABPRECOEVENTO", "ARQUIVO", v_HandleSolicitacao, v_CaminhoArquivo, True)
		Kill(v_CaminhoArquivo)
  	End If

  	q_UpdateSolicitacao.Active = False
  	q_UpdateSolicitacao.ParamByName("SITUACAOSOLICITACAO").AsString = "C"
  	q_UpdateSolicitacao.ParamByName("HANDLESOLICITACAO").AsInteger = v_HandleSolicitacao
  	q_UpdateSolicitacao.ParamByName("OCORRENCIA").AsString = "Solicitação concluída com sucesso. Já pode fazer o download do arquivo."
  	q_UpdateSolicitacao.ExecSQL

    v_Sucesso = True
    SessionVar("CONSULTATABELAPRECO_HANDLESOLICITACAO") = Null

  ErroProcesso:
	If (Not v_Sucesso) Then

		q_UpdateSolicitacao.Active = False
  	  	q_UpdateSolicitacao.ParamByName("SITUACAOSOLICITACAO").AsString = "E"
  		q_UpdateSolicitacao.ParamByName("HANDLESOLICITACAO").AsInteger = v_HandleSolicitacao
  		q_UpdateSolicitacao.ParamByName("OCORRENCIA").AsString = Err.Description
  		q_UpdateSolicitacao.ExecSQL

  		SessionVar("CONSULTATABELAPRECO_HANDLESOLICITACAO") = Null

	End If

  End If

  Set q_Solicitacao = Nothing
  Set q_UpdateSolicitacao = Nothing
  Set q_TabelaGeralEventoInicial = Nothing
  Set q_TabelaGeralEventoFinal = Nothing
  Set q_TipoPrestador = Nothing
  Set q_ParametrosWeb = Nothing
  Set r_Relatorio = Nothing

End Sub
