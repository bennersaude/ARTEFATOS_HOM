'HASH: 7829900C05CF73F8D129336ABB3259F9
'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"

Public Sub BOTAOIMPORTAR_OnClick()

On Error GoTo erro

If (InStr(SQLServer, "MSSQL") > 0) Then
    CriaTabelaTemporariaSqlServer
End If

Procedure:
On Error GoTo Erro


	Dim sql As Object
	Set sql = NewQuery
	Dim sql2 As Object
	Set sql2 = NewQuery
	Dim sql3 As Object
	Set sql3 = NewQuery

	sql.Clear
	sql.Add("SELECT NOME FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
	sql.Active = False
	sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	sql.Active = True


	Dim ServerExec As CSServerExec

	If (Not CurrentQuery.FieldByName("IDPROCESSO").IsNull) Or _
	   (CurrentQuery.FieldByName("IDPROCESSO").AsString <> "") Then
		Set ServerExec = GetServerExec(CLng(CurrentQuery.FieldByName("IDPROCESSO").AsString))
	Else
		Set ServerExec = NewServerExec
	End If
	If (ServerExec.IsExisting) Then
		If (ServerExec.Status = esRunning) Then
			bsShowMessage("A mensagem TISS está sendo processada","I")
			Exit Sub
		ElseIf (ServerExec.Status = esNone) Then
			ServerExec.RequestAbort
			Set ServerExec = NewServerExec
		Else
			If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
				bsShowMessage("A mensagem TISS já foi processada","I")
				Exit Sub
			End If
		End If
	End If

	sql2.Clear
	sql2.Add("SELECT HANDLE FROM Z_MACROS WHERE NOME = :NOME ")
	sql2.Active = False
	sql2.ParamByName("NOME").AsString = "importaTissLotePortalPrestador"
	sql2.Active = True

	Dim descricao As String

	descricao = "Importação de XML do TISS - Prestador: " + sql.FieldByName("NOME").AsString + " Arquivo Recebido: " + CurrentQuery.FieldByName("ARQUIVORECEBIDO").AsString
	If (Len(descricao) > 120) Then
		descricao = Left(descricao, 120)
	End If
	sql3.Clear
	sql3.Add("Select HANDLE FROM Z_AGENDAMENTOLOG  ")
    sql3.Add(" WHERE DESCRICAO = :DESCRICAO        ")
    sql3.Add("   AND STATUS = 1                    ")
    sql3.ParamByName("DESCRICAO").AsString = descricao
    sql3.Active = True

    If sql3.EOF Then
		ServerExec.Description = descricao
		ServerExec.Process = sql2.FieldByName("HANDLE").AsInteger
		ServerExec.SessionVar("HANDLETISS") = CurrentQuery.FieldByName("HANDLE").AsString

		ServerExec.Execute

		CurrentQuery.Edit
		CurrentQuery.FieldByName("IDPROCESSO").AsString = CStr(ServerExec.LogHandle)
		CurrentQuery.Post

		bsShowMessage("Processo enviado para processamento no servidor", "I")
	Else
	  bsshowmessage("Já existe um porcessamento em execução. Tente novamente mais tarde.", "I")
	End If

	Set ServerExec = Nothing
	Set sql = Nothing
	Set sql2 = Nothing
	Set sql3 = Nothing

	Exit Sub

Erro:
    InfoDescription = Err.Description
    CancelDescription = Err.Description
    If WebMode Then
      ServiceResult = Err.Description
    End If

End Sub

Public Sub TABLE_AfterEdit()
  ARQUIVORETORNO.ReadOnly = True
End Sub

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("USUARIODIGITADOR").AsInteger = CurrentUser
End Sub

Public Sub TABLE_AfterScroll()
	If VisibleMode Then
		BOTAOIMPORTAR.Visible = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qLocalizaPrestador As Object
  Set qLocalizaPrestador = NewQuery

  qLocalizaPrestador.Clear
  qLocalizaPrestador.Add("SELECT PRESTADOR                         ")
  qLocalizaPrestador.Add("  FROM Z_GRUPOUSUARIOS_PRESTADOR         ")
  qLocalizaPrestador.Add(" WHERE USUARIO = :USUARIO                ")
  qLocalizaPrestador.ParamByName("USUARIO").AsInteger = CurrentUser
  qLocalizaPrestador.Active = True

  If (qLocalizaPrestador.FieldByName("PRESTADOR").IsNull) Then
    bsShowMessage("Não foi possível salvar o registro, pois o usuário atual não está vinculado a um prestador!", "E")

    Set qLocalizaPrestador = Nothing

    CanContinue = False
    Exit Sub
  End If

  Set qLocalizaPrestador = Nothing

  Dim qVerificaArquivo As Object
  Set qVerificaArquivo = NewQuery

  qVerificaArquivo.Add("SELECT VALIDACAOARQUIVOREPETIDO FROM TIS_PARAMETROS")
  qVerificaArquivo.Active = True

  If (qVerificaArquivo.FieldByName("VALIDACAOARQUIVOREPETIDO").AsString = "S") Then

	  qVerificaArquivo.Clear
	  qVerificaArquivo.Add("SELECT 1 ")
	  qVerificaArquivo.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS ")
	  qVerificaArquivo.Add(" WHERE HANDLE <> :HANDLE")
	  qVerificaArquivo.Add("   AND PRESTADOR = :PRESTADOR")
	  qVerificaArquivo.Add("   AND CRC = :CRC")
	  qVerificaArquivo.Active = False
	  qVerificaArquivo.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  qVerificaArquivo.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	  qVerificaArquivo.ParamByName("CRC").AsString = CurrentQuery.FieldByName("CRC").AsString
	  qVerificaArquivo.Active = True

	  If (Not qVerificaArquivo.EOF) Then
		bsShowMessage("Não foi possível inserir o registro. O arquivo a ser inserido já foi processado no sistema!", "E")

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
  Set qLocalizaPrestador = Nothing

  ARQUIVORETORNO.ReadOnly = True

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

	If CommandID = "BOTAOIMPORTAR" Then
		BOTAOIMPORTAR_OnClick
	End If
	If CommandID = "BOTAOIMPORTARABERTOS" Then
		ImportarArquivosAbertos()
	End If
End Sub
Public Function ImportarArquivosAbertos()
	If (InStr(SQLServer, "MSSQL") > 0) Then
    	CriaTabelaTemporariaSqlServer
	End If

	Dim sql As Object
	Dim sql2 As Object
	Dim sql3 As Object
	Dim handles As String
	handles = ""

	Dim sqltiss As Object
	Set sqltiss = NewQuery
	Set sql = NewQuery
	Set sql2 = NewQuery
	Set sql3 = NewQuery

	sql.Clear
	sql.Add("SELECT GP.PRESTADOR, P.NOME                         ")
    sql.Add("  FROM Z_GRUPOUSUARIOS_PRESTADOR GP                 ")
    sql.Add("  JOIN SAM_PRESTADOR P ON (P.HANDLE = GP.PRESTADOR) ")
    sql.Add(" WHERE GP.USUARIO = :USUARIO")

    sql.ParamByName("USUARIO").AsInteger = CurrentUser
    sql.Active = True

    If sql.EOF Then
    	bsShowMessage("Usuário atual não está vinculado a um prestador!", "I")

		Set sqltiss = Nothing
		Set sql = Nothing
		Set sql2 = Nothing
		Set sql3 = Nothing

		Exit Function
    End If

	sql2.Clear
	sql2.Add("SELECT HANDLE FROM Z_MACROS WHERE NOME = :NOME ")
	sql2.Active = False
	sql2.ParamByName("NOME").AsString = "importaTissLotePortalPrestador"
	sql2.Active = True

	Dim descricao As String
	descricao = "Importação de XML do TISS (Importação em lote) - Prestador: " + sql.FieldByName("NOME").AsString + " (" +sql.FieldByName("PRESTADOR").AsString + ")"
	If (Len(descricao) > 120) Then
		descricao = Left(descricao, 120)
	End If
	sql3.Clear
	sql3.Add("Select HANDLE FROM Z_AGENDAMENTOLOG  ")
    sql3.Add(" WHERE DESCRICAO = :DESCRICAO        ")
    sql3.Add("   AND STATUS = 1                    ")
    sql3.ParamByName("DESCRICAO").AsString = descricao
    sql3.Active = True

	Dim ServerExec As CSServerExec
	Set ServerExec = NewServerExec

	If sql3.EOF Then
		sqltiss.Clear
		sqltiss.Add("SELECT M.HANDLE                     					")
		sqltiss.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS M					")
		sqltiss.Add(" WHERE M.PRESTADOR = :PRESTADOR     					")
		sqltiss.Add("   AND M.SITUACAO = :SITUACAO       					")
		sqltiss.Add("   AND NOT EXISTS (SELECT 1 							")
		sqltiss.Add("                     FROM TIS_RECURSOGLOSA T			")
		sqltiss.Add("                    WHERE T.MENSAGEMTISS = M.HANDLE) ")

		sqltiss.ParamByName("PRESTADOR").AsInteger = sql.FieldByName("PRESTADOR").AsString
		sqltiss.ParamByName("SITUACAO").AsString = "A"
		sqltiss.Active = True
		While Not sqltiss.EOF
		  If Trim(handles) = "" Then
			handles = handles + sqltiss.FieldByName("HANDLE").AsString
		  Else
		    handles = handles + "," + sqltiss.FieldByName("HANDLE").AsString
		  End If
		  sqltiss.Next
		Wend

	  ServerExec.Description = descricao
	  ServerExec.Process = sql2.FieldByName("HANDLE").AsInteger
	  ServerExec.SessionVar("HANDLETISS") = handles

	  ServerExec.Execute

	  bsShowMessage("Processo enviado para processamento no servidor", "I")
	Else
	  bsshowmessage("Já existe um porcessamento em execução. Tente novamente mais tarde.", "I")
	End If

	Set sql = Nothing
	Set sql2 = Nothing
	Set sql3 = Nothing
	Set sqltiss = Nothing
	Set ServerExec = Nothing

End Function
