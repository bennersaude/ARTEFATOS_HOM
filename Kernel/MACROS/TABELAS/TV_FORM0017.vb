'HASH: D27C0464F3038EFCB968BC65D1EA86B7
'#USES "*CriaTabelaTemporariaSqlServer"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()


	Dim sql As BPesquisa
	Set sql=NewQuery

    Dim vHandleAutorizacao As Long
    vHandleAutorizacao = RetornaNumeroAutorizacao

    sql.Clear
	sql.RequestLive = True
	sql.Add("SELECT HANDLE, ACIDENTEPESSOAL, REGIMEATENDIMENTO, TIPOTRATAMENTO, FINALIDADEATENDIMENTO, CID, INTERCORRENCIA FROM SAM_AUTORIZ WHERE HANDLE=:HANDLE")
	sql.ParamByName("HANDLE").AsInteger = vHandleAutorizacao
	sql.Active=	True

	CurrentQuery.FieldByName("ACIDENTEPESSOAL").Value = sql.FieldByName("ACIDENTEPESSOAL").Value
	CurrentQuery.FieldByName("REGIMEATENDIMENTO").Value= sql.FieldByName("REGIMEATENDIMENTO").Value
	CurrentQuery.FieldByName("TIPOTRATAMENTO").Value = sql.FieldByName("TIPOTRATAMENTO").Value
	CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").Value = sql.FieldByName("FINALIDADEATENDIMENTO").Value
	CurrentQuery.FieldByName("CID").Value= sql.FieldByName("CID").Value
	Set sql=Nothing
End Sub

Public Sub TABLE_AfterPost()
  If (Not CurrentQuery.FieldByName("ACIDENTEPESSOAL").IsNull) Or _
  	(Not CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull) Or _
  	(Not CurrentQuery.FieldByName("TIPOTRATAMENTO").IsNull) Or _
  	(Not CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").IsNull) Or _
  	(Not CurrentQuery.FieldByName("CID").IsNull) Then

	Dim sql As BPesquisa
	Set sql=NewQuery

    Dim vHandleAutorizacao As Long
    vHandleAutorizacao = RetornaNumeroAutorizacao

    sql.Clear
	sql.RequestLive = True
	sql.Add("SELECT HANDLE, ACIDENTEPESSOAL, REGIMEATENDIMENTO, TIPOTRATAMENTO, FINALIDADEATENDIMENTO, CID, INTERCORRENCIA FROM SAM_AUTORIZ WHERE HANDLE=:HANDLE")
	sql.ParamByName("HANDLE").AsInteger = vHandleAutorizacao
	sql.Active=	True

	'Gravar log da intercorrência
	Dim logIntercorrencia As CSEntity

	Set logIntercorrencia = Entity.CreateNew("SAM_AUTORIZ_INTERCORRENCIA")
	logIntercorrencia.SetValue("AUTORIZACAO", vHandleAutorizacao)
	'Valores anteriores
	logIntercorrencia.SetValue("REGIMEATENDIMENTOANTERIOR", sql.FieldByName("REGIMEATENDIMENTO").AsInteger)
	logIntercorrencia.SetValue("TIPOTRATAMENTOANTERIOR", sql.FieldByName("TIPOTRATAMENTO").AsInteger)
	logIntercorrencia.SetValue("FINALIDADEATENDIMENTOANTERIOR", sql.FieldByName("FINALIDADEATENDIMENTO").AsInteger)
	If Not sql.FieldByName("CID").IsNull Then
		logIntercorrencia.SetValue("CIDANTERIOR", sql.FieldByName("CID").AsInteger)
	End If
	logIntercorrencia.SetValue("ACIDENTEPESSOALANTERIOR", sql.FieldByName("ACIDENTEPESSOAL").AsString)

	sql.Edit
	If Not CurrentQuery.FieldByName("ACIDENTEPESSOAL").IsNull Then
	  sql.FieldByName("ACIDENTEPESSOAL").AsString = CurrentQuery.FieldByName("ACIDENTEPESSOAL").Value
	  logIntercorrencia.SetValue("ACIDENTEPESSOALINTERCORRENCIA", CurrentQuery.FieldByName("ACIDENTEPESSOAL").AsString)
	End If
	If Not CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull Then
	  sql.FieldByName("REGIMEATENDIMENTO").AsInteger= CurrentQuery.FieldByName("REGIMEATENDIMENTO").Value
	  logIntercorrencia.SetValue("REGIMEATENDIMENTOINTER", CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger)
	End If
	If Not CurrentQuery.FieldByName("TIPOTRATAMENTO").IsNull Then
	  sql.FieldByName("TIPOTRATAMENTO").AsInteger = CurrentQuery.FieldByName("TIPOTRATAMENTO").Value
	  logIntercorrencia.SetValue("TIPOTRATAMENTOINTERCORRENCIA", CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger)
	End If
	If Not CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").IsNull Then
	  sql.FieldByName("FINALIDADEATENDIMENTO").AsInteger = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").Value
	  logIntercorrencia.SetValue("FINALIDADEATENDIMENTOINTER", CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger)
	End If
	If Not CurrentQuery.FieldByName("CID").IsNull Then
	  sql.FieldByName("CID").AsInteger= CurrentQuery.FieldByName("CID").Value
	  logIntercorrencia.SetValue("CIDINTERCORRENCIA", CurrentQuery.FieldByName("CID").AsInteger)
	End If
	sql.FieldByName("INTERCORRENCIA").AsString="S"
	sql.Post
	Set sql=Nothing

	'Gravando log da intercorrência
	logIntercorrencia.Save

	'REVALIDAR A AUTORIZAÇÃO
	revalidarAutoriz
  End If
End Sub

Public Sub revalidarAutoriz
    CriaTabelaTemporariaSqlServer
	Dim sql As BPesquisa
	Set sql = NewQuery

    Dim vHandleAutorizacao As Long
    vHandleAutorizacao = RetornaNumeroAutorizacao

    sql.Clear
	sql.Add("SELECT ES.HANDLE ")
	sql.Add("  FROM SAM_AUTORIZ_EVENTOSOLICIT ES ")
	If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then
	  sql.Add(" JOIN SAM_AUTORIZ_EVENTOGERADO EG ON ES.HANDLE = EG.EVENTOSOLICITADO ")
	End If

	sql.Add(" WHERE ES.AUTORIZACAO=:AUTORIZACAO")

	If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then
	  sql.Add(" AND EG.PROTOCOLOTRANSACAO = :PROTOCOLOTRANSACAO ")
	  sql.ParamByName("PROTOCOLOTRANSACAO").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
	End If
	sql.ParamByName("AUTORIZACAO").AsInteger = vHandleAutorizacao
	sql.Active=True

	Dim SQLAgendado As Object
	Set SQLAgendado = NewQuery
	SQLAgendado.Add("SELECT EXECUTAAUTORIZACAOAGENDADA FROM SAM_PARAMETROSWEB")
	SQLAgendado.Active = True

	If SQLAgendado.FieldByName("EXECUTAAUTORIZACAOAGENDADA").AsString = "S" Then
		Dim vsMensagemErro As String
		Dim Obj As Object
		Dim vcContainer As CSDContainer
		Set vcContainer = NewContainer
		vcContainer.AddFields("HANDLE:INTEGER")
		vcContainer.Insert
		Dim retorno As Long

		vcContainer.Field("HANDLE").AsInteger = sql.FieldByName("HANDLE").AsInteger

		Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
		retorno = Obj.ExecucaoImediata(CurrentSystem, _
									   "CA043", _
	                                   "RevalidarSolicitado", _
	                                   "Processamento de Autorização", _
	                                   sql.FieldByName("HANDLE").AsInteger, _
		                               "SAM_AUTORIZ_EVENTOSOLICIT", _
		                               "SITUACAOPROCESSAMENTO", _
		                               "", _
		                               "", _
		                               "P", _
		                               True, _
		                               vsMensagemErro, _
		                               vcContainer)
		Set sql=Nothing
		If retorno = 0 Then
			bsShowMessage("Processo enviado para execução no servidor!", "I")
			bsShowMessage("Autorização sendo revalidada!", "I")

		Else
			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		End If

	Else
		Dim erro As String
		Dim alertas As String
		Dim dll As Object
		Set dll=CreateBennerObject("ca043.autorizacao")
		dll.inicializar(CurrentSystem, "I")

		While Not sql.EOF
			If dll.revalidarSolicitado(CurrentSystem, sql.FieldByName("HANDLE").AsInteger, alertas, erro) < 0 Then
				Exit While
			End If
			sql.Next
		Wend

		dll.finalizar
		Set sql=Nothing

		If erro<>"" Then
			bsShowMessage(erro, "I")
		Else
			bsShowMessage("Operação concluída com sucesso!", "I")
		End If
	End If

End Sub

Public Function RetornaNumeroAutorizacao As Long

  RetornaNumeroAutorizacao = RecordHandleOfTable("SAM_AUTORIZ")

  If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then
    Dim qBuscaHandleAutorizacao As Object
    Set qBuscaHandleAutorizacao  = NewQuery

    qBuscaHandleAutorizacao.Clear
    qBuscaHandleAutorizacao.Add("SELECT AUTORIZACAO FROM SAM_PROTOCOLOTRANSACAOAUTORIZ WHERE HANDLE = :HANDLE")
    qBuscaHandleAutorizacao.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
    qBuscaHandleAutorizacao.Active = True
    RetornaNumeroAutorizacao = qBuscaHandleAutorizacao.FieldByName("AUTORIZACAO").AsInteger

    Set qBuscaHandleAutorizacao  = Nothing
  End If

End Function
