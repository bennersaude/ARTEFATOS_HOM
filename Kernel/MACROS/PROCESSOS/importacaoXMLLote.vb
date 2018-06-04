'HASH: AA1DDBC359A15AFE582BFB90708C34CD
'#Uses "*bsShowMessage"
Public Sub Main

  Dim interface As Object
  Dim qArquivos As Object
  Dim dllValidarMensagem As Object
  Dim qOcorrencias As Object
  Dim qSituacaoMensagemTISS As Object

  Set qArquivos = NewQuery
  Set qOcorrencias = NewQuery
  Set qSituacaoMensagemTISS = NewQuery

  Set interface = CreateBennerObject("Benner.Saude.WSTiss.ImportacaoXmlLote.ImportacaoXmlLote")
  Set dllValidarMensagem = CreateBennerObject("Benner.Saude.WSTiss.Versionador.VersionadorImportarMensagemTISS")

  Dim vsRetorno As String
  Dim viHandleRotina As Long
  Dim arquivoComErro As Boolean

  qSituacaoMensagemTISS.Clear
  qSituacaoMensagemTISS.Add("SELECT OCORRENCIAS, SITUACAO, ARQUIVORECEBIDO FROM SAM_PRESTADOR_MENSAGEMTISS WHERE HANDLE =  :HANDLE")

  arquivoComErro = False


  If SessionVar("HANDLEROTINAXMLLOTE") <> "" Then
    viHandleRotina = CLng(SessionVar("HANDLEROTINAXMLLOTE"))
    vsRetorno = "OK"
    SessionVar("HANDLEROTINAXMLLOTE") = ""
  Else
    viHandleRotina = interface.CriarRegistroAutomatico(CurrentSystem)
    vsRetorno = interface.ImportarXML(CurrentSystem, viHandleRotina )
  End If

  If vsRetorno = "OK" Then

    qOcorrencias.Clear
    qOcorrencias.Add("UPDATE TIS_IMPORTACAOXMLLOTE SET SITUACAO = :SITUACAO WHERE HANDLE = :HANDLE")
    qOcorrencias.ParamByName("HANDLE").AsInteger = viHandleRotina
    qOcorrencias.ParamByName("SITUACAO").AsString = "2"
    qOcorrencias.ExecSQL

    If (InStr(SQLServer, "MSSQL") > 0) Then
      Dim dll As Object
      Set dll = CreateBennerObject("SAMUTIL.Rotinas")
      dll.CriaTabelaTemporariaSqlServer(CurrentSystem, 0)
      Set dll = Nothing
    End If

    qArquivos.Clear
    qArquivos.Add("SELECT MENSAGEMTISS")
    qArquivos.Add("  FROM TIS_IMPORTACAOXMLLOTE_ARQ")
    qArquivos.Add(" WHERE IMPORTACAOXMLLOTE = :XMLLOTE")
    qArquivos.ParamByName("XMLLOTE").AsInteger = viHandleRotina
    qArquivos.Active = True

    interface.AtualizarOcorrenciasRotina(CurrentSystem, viHandleRotina, "===PROCESSAMENTO DOS ARQUIVOS===")

	Dim xml As String

    If qArquivos.EOF Then
      interface.AtualizarOcorrenciasRotina(CurrentSystem, viHandleRotina, "Não há arquivos a serem processados.")
    End If

    While Not qArquivos.EOF

      SessionVar("HANDLE") = qArquivos.FieldByName("MENSAGEMTISS").AsString
      SessionVar("HANDLETABELA_TISS") = qArquivos.FieldByName("MENSAGEMTISS").AsString
      SessionVar("NOMETABELA_TISS") = "SAM_PRESTADOR_MENSAGEMTISS"
      SessionVar("NOMECAMPO_TISS") = "ARQUIVORECEBIDO"
      SessionVar("HANDLE_TISVERSAO") = "0"

	  dllValidarMensagem.Exec(CurrentSystem)

      qSituacaoMensagemTISS.Active = False
      qSituacaoMensagemTISS.ParamByName("HANDLE").AsInteger = qArquivos.FieldByName("MENSAGEMTISS").AsInteger
      qSituacaoMensagemTISS.Active = True

      If qSituacaoMensagemTISS.FieldByName("SITUACAO").AsString = "E" Then
        arquivoComErro = True
        interface.AtualizarOcorrenciasRotina(CurrentSystem, viHandleRotina, "Arquivo: " + qSituacaoMensagemTISS.FieldByName("ARQUIVORECEBIDO").AsString + " - Verificar ocorrências, arquivo processado com erro")
        interface.AtualizarOcorrenciasRotina(CurrentSystem, viHandleRotina, qSituacaoMensagemTISS.FieldByName("OCORRENCIAS").AsString)
      Else
        interface.AtualizarOcorrenciasRotina(CurrentSystem, viHandleRotina, "Arquivo: " + qSituacaoMensagemTISS.FieldByName("ARQUIVORECEBIDO").AsString + " - Arquivo processado com sucesso")
      End If

      qArquivos.Next
    Wend

  Else
    interface.AtualizarOcorrenciasRotina(CurrentSystem, viHandleRotina, vsRetorno)
  End If

  qOcorrencias.Clear
  qOcorrencias.Add("UPDATE TIS_IMPORTACAOXMLLOTE SET SITUACAO = :SITUACAO, USUARIOPROCESSAMENTO = :USUARIOPROC, DATAHORAPROCESSAMENTO = :DATAHORAPROC WHERE HANDLE = :HANDLE")
  qOcorrencias.ParamByName("HANDLE").AsInteger = viHandleRotina

  If arquivoComErro Then
    qOcorrencias.ParamByName("SITUACAO").AsString = "4"

    Dim sp As BStoredProc
	Set sp = NewStoredProc

	sp.Name = "BS_CA741E47"
	sp.AddParam("p_HandleRotImpLote",ptInput, ftInteger,4)

	sp.ParamByName("p_HandleRotImpLote").AsInteger = viHandleRotina
	sp.ExecProc

	Set sp  = Nothing

  Else
    qOcorrencias.ParamByName("SITUACAO").AsString = "3"
  End If
  qOcorrencias.ParamByName("USUARIOPROC").AsInteger = CurrentUser
  qOcorrencias.ParamByName("DATAHORAPROC").AsDateTime = ServerNow
  qOcorrencias.ExecSQL

  Set interface = Nothing
  Set qArquivos = Nothing
  Set dllValidarMensagem = Nothing
  Set qOcorrencias = Nothing
  Set qSituacaoMensagemTISS = Nothing

End Sub
