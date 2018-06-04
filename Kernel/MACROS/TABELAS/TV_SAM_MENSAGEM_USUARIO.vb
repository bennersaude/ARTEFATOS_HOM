'HASH: A8BEFE5A5C7D5E1C4BF10C13A5661A88
Public Sub TABLE_AfterPost()

  Dim qUsuario As BPesquisa
  Set qUsuario = NewQuery

  qUsuario.Add(" SELECT U.HANDLE FROM Z_GRUPOUSUARIOS U JOIN Z_GRUPOS G ON U.GRUPO = G.HANDLE WHERE ")

  If CurrentQuery.FieldByName("TABDESTINATARIO").AsInteger = 1 Then
	qUsuario.Add(" U.HANDLE = :HUSUARIO ")
	qUsuario.ParamByName("HUSUARIO").Value = CurrentQuery.FieldByName("DESTINATARIO").AsInteger
  ElseIf CurrentQuery.FieldByName("TABDESTINATARIO").AsInteger = 2 Then
	qUsuario.Add(" U. GRUPO = :HGRUPO ")
	qUsuario.ParamByName("HGRUPO").Value = CurrentQuery.FieldByName("GRUPOUSUARIODESTINATARIO").AsInteger
  End If

  qUsuario.Active = True

  Dim qAssunto As BPesquisa
  Set qAssunto = NewQuery
  Dim vSQL As String

  vSQL = "SELECT ASSUNTO FROM SAM_MENSAGENSPADRAO WHERE HANDLE = :HANDLE"

  qAssunto.Clear
  qAssunto.Add(vSQL)
  qAssunto.ParamByName("HANDLE").AsInteger      = CurrentQuery.FieldByName("ASSUNTO").AsInteger
  qAssunto.Active = True

  Dim qInsereMensagemUsuario As BPesquisa
  Set qInsereMensagemUsuario = NewQuery
  qInsereMensagemUsuario.Clear
  qInsereMensagemUsuario.Add(" INSERT INTO SAM_MENSAGEM_USUARIO( HANDLE,  REMETENTE,  DESTINATARIO,  ASSUNTO,  TEXTO,  URGENTE,  DATAHORAENVIO) ")
  qInsereMensagemUsuario.Add("                           VALUES(:HANDLE, :REMETENTE, :DESTINATARIO, :ASSUNTO, :TEXTO, :URGENTE, :DATAHORAENVIO) ")

  Dim qInsereMensagemGrupo As BPesquisa
  Set qInsereMensagemGrupo = NewQuery
  qInsereMensagemGrupo.Clear
  qInsereMensagemGrupo.Add(" INSERT INTO Z_GRUPOUSUARIOMENSAGENS( HANDLE,  USUARIO,  DE,  DATAHORA,  ASSUNTO,  TEXTO,  LIDA) ")
  qInsereMensagemGrupo.Add("                              VALUES(:HANDLE, :USUARIO, :DE, :DATAHORA, :ASSUNTO, :TEXTO, :LIDA) ")

  While Not(qUsuario.EOF)

	qInsereMensagemGrupo.ParamByName("HANDLE").AsInteger       = NewHandle("Z_GRUPOUSUARIOMENSAGENS")
	qInsereMensagemGrupo.ParamByName("USUARIO").AsInteger      = qUsuario.FieldByName("HANDLE").AsInteger
	qInsereMensagemGrupo.ParamByName("DE").AsInteger           = CurrentQuery.FieldByName("REMETENTE").AsInteger
	qInsereMensagemGrupo.ParamByName("DATAHORA").AsDateTime    = ServerNow

  	If CurrentQuery.FieldByName("URGENTE").AsString = "S" Then
      qInsereMensagemGrupo.ParamByName("ASSUNTO").AsString    = "URGENTE| " + qAssunto.FieldByName("ASSUNTO").AsString
    ElseIf CurrentQuery.FieldByName("URGENTE").AsString <> "S" Then
      qInsereMensagemGrupo.ParamByName("ASSUNTO").AsString    = qAssunto.FieldByName("ASSUNTO").AsString
    End If

	qInsereMensagemGrupo.ParamByName("LIDA").AsString          = "N"
    qInsereMensagemGrupo.ParamByName("TEXTO").AsString         = CurrentQuery.FieldByName("TEXTO").AsString

	qInsereMensagemGrupo.ExecSQL

    qInsereMensagemUsuario.ParamByName("HANDLE").AsInteger         = NewHandle("SAM_MENSAGEM_USUARIO")
    qInsereMensagemUsuario.ParamByName("REMETENTE").AsInteger      = CurrentQuery.FieldByName("REMETENTE").AsInteger
    qInsereMensagemUsuario.ParamByName("DESTINATARIO").AsInteger   = qUsuario.FieldByName("HANDLE").AsInteger

  	If CurrentQuery.FieldByName("URGENTE").AsString = "S" Then
      qInsereMensagemUsuario.ParamByName("TEXTO").AsString     = "URGENTE| " + CurrentQuery.FieldByName("TEXTO").AsString
    ElseIf CurrentQuery.FieldByName("URGENTE").AsString <> "S" Then
      qInsereMensagemUsuario.ParamByName("TEXTO").AsString     = CurrentQuery.FieldByName("TEXTO").AsString
    End If

    qInsereMensagemUsuario.ParamByName("URGENTE").AsString         = CurrentQuery.FieldByName("URGENTE").AsString
    qInsereMensagemUsuario.ParamByName("DATAHORAENVIO").AsDateTime = ServerNow
    qInsereMensagemUsuario.ParamByName("ASSUNTO").AsInteger        = CurrentQuery.FieldByName("ASSUNTO").AsInteger

    qInsereMensagemUsuario.ExecSQL

	qUsuario.Next
  Wend

  Set qInsereMensagemGrupo   = Nothing
  Set qInsereMensagemUsuario = Nothing
  Set qUsuario               = Nothing
  Set qAssunto               = Nothing

End Sub


Public Sub ASSUNTO_OnChange()
  Dim qMensagem As BPesquisa
  Set qMensagem = NewQuery
  qMensagem.Clear
  qMensagem.Add("SELECT TEXTO FROM SAM_MENSAGENSPADRAO WHERE HANDLE = :HANDLE")
  qMensagem.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ASSUNTO").AsInteger
  qMensagem.Active = True

  CurrentQuery.FieldByName("TEXTO").AsString = qMensagem.FieldByName("TEXTO").AsString

  Set qMensagem = Nothing

End Sub
