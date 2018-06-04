'HASH: AD63A94EF796D7768B36F20D235BD413
'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"

Public Sub BOTAOIMPORTAR_OnClick()
On Error GoTo erro

If (InStr(SQLServer, "MSSQL") > 0) Then
    CriaTabelaTemporariaSqlServer
    'Dim dll As Object
    'Set dll = CreateBennerObject("SAMUTIL.Rotinas")
    'dll.CriaTabelaTemporaria(CurrentSystem, 0)
    'Set dll = Nothing
End If


Procedure:
On Error GoTo Erro

  'Dim dllValidarMensagem As Object
  'Dim xml As String
  'set dllValidarMensagem = CreateBennerObject("Benner.Saude.WSTiss.Versionador.VersionadorImportarValidarMsgTISS")
  'SessionVar("HANDLE") = CStr(RecordHandleOfTable("SAM_PRESTADOR_VALIDAMSGTISS"))
  'SessionVar("HANDLETABELA_TISS") = CStr(RecordHandleOfTable("SAM_PRESTADOR_VALIDAMSGTISS"))
  'SessionVar("NOMETABELA_TISS") = "SAM_PRESTADOR_VALIDAMSGTISS"
  'SessionVar("NOMECAMPO_TISS") = "ARQUIVOVALIDADO"
  'SessionVar("HANDLE_TISVERSAO") = "0"
  'dllValidarMensagem.Exec(CurrentSystem)
  'Set dllValidarMensagem = Nothing
  'Exit Sub

  Dim sql As Object
  Set sql = NewQuery

  sql.Clear
  sql.Add("SELECT NOME FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  sql.Active = False
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  sql.Active = True

  Dim serverExec As CSServerExec
  Set serverExec = NewServerExec

  serverExec.Description = "Importação de XML do TISS - Prestador: " + sql.FieldByName("NOME").AsString
  serverExec.DllClassName = "Benner.Saude.WSTiss.Versionador.VersionadorImportarValidarMsgTISS"
  serverExec.SessionVar("HANDLE") = CStr(RecordHandleOfTable("SAM_PRESTADOR_VALIDAMSGTISS"))
  serverExec.SessionVar("HANDLETABELA_TISS") = CStr(RecordHandleOfTable("SAM_PRESTADOR_VALIDAMSGTISS"))
  serverExec.SessionVar("NOMETABELA_TISS") = "SAM_PRESTADOR_VALIDAMSGTISS"
  serverExec.SessionVar("NOMECAMPO_TISS") = "ARQUIVOVALIDADO"
  serverExec.SessionVar("HANDLE_TISVERSAO") = "0"

  serverExec.Execute

  bsShowMessage("Processo enviado para o servidor", "I")

  Set serverExec = Nothing
  Set sql = Nothing

  Exit Sub


Erro:
    InfoDescription = Err.Description
    CancelDescription = Err.Description
    ServiceResult = Err.Description

End Sub

Public Sub BOTAOVALIDAR_OnClick()
'  Dim dllValidarMensagem As Object
'  Set dllValidarMensagem = CreateBennerObject("Benner.Saude.WSTiss.Versionador.VersionadorValidarValidarMsgTISS")
'  SessionVar("HANDLE") = CStr(RecordHandleOfTable("SAM_PRESTADOR_VALIDAMSGTISS"))
'  SessionVar("HANDLETABELA_TISS") = CStr(RecordHandleOfTable("SAM_PRESTADOR_VALIDAMSGTISS"))
'  SessionVar("NOMETABELA_TISS") = "SAM_PRESTADOR_VALIDAMSGTISS"
'  SessionVar("NOMECAMPO_TISS") = "ARQUIVOVALIDADO"
'  SessionVar("HANDLE_TISVERSAO") = "0"
'  dllValidarMensagem.Exec(CurrentSystem)
'  Set dllValidarMensagem = Nothing
'  Exit Sub

  Dim sql As Object
  Set sql = NewQuery

  sql.Clear
  sql.Add("SELECT NOME FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  sql.Active = False
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  sql.Active = True

  Dim serverExec As CSServerExec
  Set serverExec = NewServerExec

  serverExec.Description = "Validação de XML do TISS - Prestador: " + sql.FieldByName("NOME").AsString
  serverExec.DllClassName = "Benner.Saude.WSTiss.Versionador.VersionadorValidarValidarMsgTISS"
  serverExec.SessionVar("HANDLE") = CStr(RecordHandleOfTable("SAM_PRESTADOR_VALIDAMSGTISS"))
  serverExec.SessionVar("HANDLETABELA_TISS") = CStr(RecordHandleOfTable("SAM_PRESTADOR_VALIDAMSGTISS"))
  serverExec.SessionVar("NOMETABELA_TISS") = "SAM_PRESTADOR_VALIDAMSGTISS"
  serverExec.SessionVar("NOMECAMPO_TISS") = "ARQUIVOVALIDADO"
  serverExec.SessionVar("HANDLE_TISVERSAO") = "0"

  serverExec.Execute

  bsShowMessage("Processo enviado para o servidor", "I")

  Set serverExec = Nothing
  Set sql = Nothing


End Sub

Public Sub TABLE_AfterScroll()
	If VisibleMode Then
		BOTAOIMPORTAR.Visible = False
		BOTAOVALIDAR.Visible = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qVerificaArquivo As Object
  Set qVerificaArquivo = NewQuery

  qVerificaArquivo.Add("SELECT VALIDACAOARQUIVOREPETIDO FROM TIS_PARAMETROS")
  qVerificaArquivo.Active = True

  If CurrentQuery.FieldByName("PRESTADOR").AsInteger <= 0 Then
    bsShowMessage("Usuário atual não está vinculado a um prestador!", "I")
    CancelDescription = "Usuário atual não está vinculado a um prestador!"
    Set qVerificaArquivo = Nothing
    CanContinue = False
    Exit Sub
  End If

  If (qVerificaArquivo.FieldByName("VALIDACAOARQUIVOREPETIDO").AsString = "S") Then

	  qVerificaArquivo.Clear
	  qVerificaArquivo.Add("SELECT 1 ")
	  qVerificaArquivo.Add("  FROM SAM_PRESTADOR_VALIDAMSGTISS ")
	  qVerificaArquivo.Add(" WHERE HANDLE <> :HANDLE")
	  qVerificaArquivo.Add("   AND PRESTADOR = :PRESTADOR")
	  qVerificaArquivo.Add("   AND ARQUIVOVALIDADO = :ARQUIVO")
	  qVerificaArquivo.Active = False
	  qVerificaArquivo.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  qVerificaArquivo.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	  qVerificaArquivo.ParamByName("ARQUIVO").AsString = CurrentQuery.FieldByName("ARQUIVOVALIDADO").AsString
	  qVerificaArquivo.Active = True

	  If (Not qVerificaArquivo.EOF) Then
		bsShowMessage("Não foi possível inserir o registro. Existe um arquivo com o mesmo nome do arquivo a ser inserido!", "E")

		CanContinue = False
		Exit Sub
	  End If
  End If
End Sub

Public Sub TABLE_NewRecord()
  Dim qLocalizaPrestador As Object
  Set qLocalizaPrestador = NewQuery
  qLocalizaPrestador.Clear
  qLocalizaPrestador.Add("SELECT PRESTADOR                         ")
  qLocalizaPrestador.Add("  FROM Z_GRUPOUSUARIOS_PRESTADOR         ")
  qLocalizaPrestador.Add(" WHERE USUARIO = :USUARIO                ")
  qLocalizaPrestador.ParamByName("USUARIO").AsInteger = CurrentUser
  qLocalizaPrestador.Active = True


  CurrentQuery.FieldByName("PRESTADOR").Value = qLocalizaPrestador.FieldByName("PRESTADOR").AsInteger

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOVALIDAR" Then
		BOTAOVALIDAR_OnClick
	ElseIf CommandID = "BOTAOIMPORTAR" Then
		BOTAOIMPORTAR_OnClick
	End If
End Sub
