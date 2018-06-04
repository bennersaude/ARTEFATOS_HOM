'HASH: 84BDA902BC22901AE25690554142D66B
'Macro: SAM_GUIA
'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"
'#Uses "*PermissaoAlteracao"
'#Uses "*TV_FORM0143_VALIDACAO"

Option Explicit

Dim cancontinue              As Boolean
Dim OldAutorizacao           As Long
Dim OldExecutor              As Long
Dim Regime_Atend             As Long
Dim Data_Atend               As Date
Dim DataBase                 As Date
Dim old_modeloguia           As Long
Dim gModeloGuia              As Long
Dim ERROIDENTIFICADOR        As Boolean
Dim OldNumeroContingencia    As String
Dim OLDBENEFICIARIO          As Long
Dim vSituacaoGuia            As String
Dim vSituacaoPeg             As String
Dim vSMensagem               As String
Dim viState                  As Long
Dim vFormacaoPrestador       As String
Dim vDllBSPro006             As Object

' Botões ------------------------------------------------------------

Public Sub BOTAOALTERARBENEFICIARIO_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOALTERARBENEFICIARIO_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOATUALIZARVALOR_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOATUALIZARVALOR_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOODONTOGRAMA_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOODONTOGRAMA_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOAPROVARVALOR_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOAPROVARVALOR_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOCANCELARPROVISAO_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOCANCELARPROVISAO_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOCONFERIDA_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOCONFERIDA_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAODEVOLVERGUIA_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAODEVOLVERGUIA_OnClick(CurrentSystem, CurrentQuery.TQuery, cancontinue)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAODIGITAR_OnClick()

	'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
	Dim aux As Boolean
	aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "P")
	If aux = False Then
		Exit Sub
	End If

	If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
		If CurrentQuery.State <>1 Then
			bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
			Exit Sub
		End If

		Dim INTERFACE As Object

		Set INTERFACE = CreateBennerObject("BSPRO006.DIGITACAO")

		INTERFACE.Digitar(CurrentSystem, CurrentQuery.FieldByName("PEG").AsInteger)

		Set INTERFACE = Nothing
	Else
		bsShowMessage("Comando válido somente na fase de digitação", "I")
	End If

End Sub

Public Sub BOTAOGLOSATOTAL_OnClick()

	'Em modo web o código do click do botão será executado após o post da tabela virtual
	'Desta forma, uma parte deste evento será executada apenas em modo desktop
	If VisibleMode Then
		If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
			bsShowMessage("Guia já paga não pode ser glosada", "I")
			Exit Sub
		End If

		Dim aux As Boolean
		'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
		If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then 'digitacao
			aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "A")
		Else
			aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "P")
		End If

		If aux = False Then
			Exit Sub
		End If

		Dim vSMensagem As String

		'Verificar se será permitido acionar a funcionalidade se o PEG estiver sendo provisionado
		If PermissaoAlteracao(0, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, False, vSMensagem) = 1 Then
			MsgBox(vSMensagem)
			Exit Sub
		End If

		Dim INTERFACE As Object
		Dim vsMsgErro As String
		Dim viRetorno As Integer

		Set INTERFACE = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
		viRetorno = INTERFACE.Exec(CurrentSystem, _
									1, _
									"TV_FORM0053", _
									"Glosa Total", _
									0, _
									290, _
									516, _
									False, _
									vsMsgErro, _
									Null)
		If viRetorno = 1 Then
			bsShowMessage(vsMsgErro, "I")
			Exit Sub
		ElseIf viRetorno = -1 Then
			Exit Sub
		End If

		Set INTERFACE = Nothing
	End If

	If VisibleMode Then
		Dim qglosasOK As Object
		Set qglosasOK = NewQuery

		qglosasOK.Clear
		qglosasOK.Add("SELECT COUNT(1) QTDE             ")
		qglosasOK.Add("  FROM SAM_GUIA_EVENTOS_GLOSA  GL")
		qglosasOK.Add("  JOIN SAM_GUIA_EVENTOS E ON E.HANDLE = GL.GUIAEVENTO ")
		qglosasOK.Add(" WHERE E.GUIA = :GUIA")
		qglosasOK.Add("   AND GL.GLOSAREVISADA = 'N' ")
		qglosasOK.Add("   AND E.COPIAEVENTOORIGINAL <> 'S'")
		qglosasOK.ParamByName("GUIA").Value = CurrentQuery.FieldByName("HANDLE").Value
		qglosasOK.Active = True

		If qglosasOK.FieldByName("QTDE").AsInteger = 0 Then
			Dim qnegacoesOK As Object
			Set qnegacoesOK = NewQuery

			qnegacoesOK.Clear
			qnegacoesOK.Add("SELECT COUNT(1) QTDE              ")
			qnegacoesOK.Add("  FROM SAM_GUIA_EVENTOS_NEGACAO EN")
			qnegacoesOK.Add("  JOIN SAM_GUIA_EVENTOS E ON E.HANDLE = EN.GUIAEVENTO ")
			qnegacoesOK.Add(" WHERE E.GUIA = :GUIA")
			qnegacoesOK.Add("   AND EN.NEGACAOREVISADA = 'N' ")
			qnegacoesOK.Add("   AND E.COPIAEVENTOORIGINAL <> 'S'")
			qnegacoesOK.ParamByName("GUIA").Value = CurrentQuery.FieldByName("HANDLE").Value
			qnegacoesOK.Active = True

			If qnegacoesOK.FieldByName("QTDE").AsInteger = 0 Then

				Dim qSituacaoGuia As Object
				Set qSituacaoGuia = NewQuery

				qSituacaoGuia.Clear
				qSituacaoGuia.Add("SELECT SITUACAO       ")
				qSituacaoGuia.Add("  FROM SAM_GUIA       ")
				qSituacaoGuia.Add(" WHERE HANDLE = :GUIA")
				qSituacaoGuia.ParamByName("GUIA").Value = CurrentQuery.FieldByName("HANDLE").Value
				qSituacaoGuia.Active = True

				If VisibleMode Then
					If (qSituacaoGuia.FieldByName("SITUACAO").AsString <> vSituacaoGuia) Then
						RefreshNodesWithTable("SAM_GUIA")
					End If
				End If

				Set qSituacaoGuia = Nothing
			End If

			Set qnegacoesOK = Nothing
		End If

		Set qglosasOK = Nothing
	End If

End Sub

Public Sub BOTAOTROCARLEIAUTE_OnClick()

	'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
	Dim aux As Boolean

	aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "P")

	If aux = False Then
		Exit Sub
	End If

End Sub

Public Sub BOTAOGUIAORIGINAL_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOGUIAORIGINAL_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAONAOFINANCIAR_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAONAOFINANCIAR_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOPFINTEGRAL_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOPFINTEGRAL_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOREPROCESSARGUIA_OnClick()

	'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
	Dim aux As Boolean

	Dim qSituacaoAntiga As Object
	Dim vsSituacaoPeg As String
	Dim vsSituacaoGuia As String
	Dim vsSituacaoEvento As String

	Set qSituacaoAntiga = NewQuery

	qSituacaoAntiga.Clear
	qSituacaoAntiga.Add("SELECT SITUACAO ")
	qSituacaoAntiga.Add("  FROM SAM_PEG ")
	qSituacaoAntiga.Add(" WHERE HANDLE IN (SELECT PEG ")
	qSituacaoAntiga.Add("                    FROM SAM_GUIA ")
	qSituacaoAntiga.Add("                   WHERE HANDLE = :HGUIA) ")
	qSituacaoAntiga.ParamByName("HGUIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qSituacaoAntiga.Active = True

	vsSituacaoPeg = qSituacaoAntiga.FieldByName("SITUACAO").AsString
	vsSituacaoGuia = CurrentQuery.FieldByName("SITUACAO").AsString

	Set qSituacaoAntiga = Nothing

	aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "P")

	If aux = False Then
		bsShowMessage("Usuário sem permissão nesta filial", "I")
		Exit Sub
	End If

	Dim Continuar As Boolean
	Dim qPeg As Object

	VERIFICAJAPAGO(Continuar)

	If Continuar = False Then
		bsShowMessage("A Guia já foi paga e não pode ser modificada", "I")
		Exit Sub
	End If

	Dim INTERFACE As Object

	BOTAOREPROCESSARGUIA.Enabled = False

	If VisibleMode Then
		Set INTERFACE = CreateBennerObject("BSINTERFACE0046.Rotinas")
		INTERFACE.ReprocessarGuia(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
	Else
		Dim vsMensagemErro As String
		Dim Obj As Object
		Dim viRet As Long
		Dim vcContainer As CSDContainer
		Set vcContainer = NewContainer
		vcContainer.AddFields("HANDLE:INTEGER")

		vcContainer.Insert
		vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

		Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
		viRet = Obj.ExecucaoImediata(CurrentSystem, _
									"BSPRO000", _
									"VerificarEventosGuia", _
									"Reprocessamento de GUIA", _
									CurrentQuery.FieldByName("HANDLE").AsInteger, _
									"SAM_GUIA", _
									"SITUACAOPROCESSAMENTO", _
									"", _
									"", _
									"P", _
									True, _
									vsMensagemErro, _
									vcContainer)

		If viRet = 0 Then
			bsShowMessage("Processo enviado para execução no servidor!", "I")
		Else
			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		End If
	End If

	BOTAOREPROCESSARGUIA.Enabled = True
	CurrentQuery.Active = False
	CurrentQuery.Active = True

	CHECATETOREEMBOLSO

	Set INTERFACE = Nothing

	Dim qSituacao As Object
	Set qSituacao = NewQuery

	qSituacao.Clear
	qSituacao.Add("SELECT SITUACAO ")
	qSituacao.Add("  FROM SAM_PEG ")
	qSituacao.Add(" WHERE HANDLE IN (SELECT PEG ")
	qSituacao.Add("                    FROM SAM_GUIA ")
	qSituacao.Add("                    WHERE HANDLE = :HGUIA)")
	qSituacao.ParamByName("HGUIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qSituacao.Active = True

	If qSituacao.FieldByName("SITUACAO").AsString <> vsSituacaoPeg Then
		RefreshNodesWithTable("SAM_PEG")
	Else
		If CurrentQuery.FieldByName("SITUACAO").AsString <> vsSituacaoGuia Then
			RefreshNodesWithTable("SAM_GUIA")
		Else
			qSituacao.Clear
			qSituacao.Add("SELECT COUNT(1) QTDE ")
			qSituacao.Add("  FROM SAM_GUIA_EVENTOS ")
			qSituacao.Add(" WHERE GUIA = :HGUIA")
			qSituacao.Add("   AND SITUACAO <> :SITUACAO")
			qSituacao.Add("   AND COPIAEVENTOORIGINAL <> 'S'")
			qSituacao.ParamByName("HGUIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
			qSituacao.ParamByName("SITUACAO").AsString = vsSituacaoGuia
			qSituacao.Active = True

			If qSituacao.FieldByName("QTDE").AsInteger > 0 Then
				RefreshNodesWithTable("SAM_GUIA_EVENTOS")
			End If
		End If
	End If

	Set qSituacao = Nothing

End Sub

Public Sub BOTAOVERIFICAMONITORAMENTO_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.BOTAOVERIFICAMONITORAMENTO_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

' Campos ------------------------------------------------------------

Public Sub AUTORIZACAO_OnExit()

	HerdarAutorizacao

End Sub

Public Sub AUTORIZACAO_OnPopup(ShowPopup As Boolean)
	Dim ProcuraDLL As Object
	Dim handlexx As Long
	ShowPopup = False
	Dim vPos As Integer

	Dim SQL As Object
	Set SQL = NewQuery

	On Error GoTo prox
	SQL.Add("SELECT HANDLE FROM SAM_AUTORIZ WHERE AUTORIZACAO=:AUTORIZ")
	SQL.ParamByName("AUTORIZ").AsString = AUTORIZACAO.Text
	SQL.Active = True

	If Not SQL.EOF Then
		CurrentQuery.FieldByName("AUTORIZACAO").AsInteger = SQL.FieldByName("HANDLE").AsInteger
		Set SQL = Nothing
		Exit Sub
	End If

	prox:
	Set SQL = Nothing
	Set ProcuraDLL = CreateBennerObject("Procura.Procurar")

	If (IsNumeric(AUTORIZACAO.Text)) Then
		vPos = 2 'realiza a busca por autorizacao
	Else
		vPos = 4 'realiza a busca por beneficiario
	End If

	handlexx = ProcuraDLL.Exec(CurrentSystem, "SAM_AUTORIZ|SAM_BENEFICIARIO[SAM_BENEFICIARIO.HANDLE=SAM_AUTORIZ.BENEFICIARIO]", "SAM_AUTORIZ.DATAAUTORIZACAO|SAM_AUTORIZ.AUTORIZACAO|SAM_AUTORIZ.RADIOSOLICITACAO|SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_AUTORIZ.SENHAGUIASOLICITACAO", vPos, "Data da Autorização|Autorização|Solicitação|Nome|Matrícula Funcional|Senha da solicitação", "SAM_AUTORIZ.HANDLE > 0 ", "Procura Autorização", False, AUTORIZACAO.Text)

	If handlexx <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("AUTORIZACAO").Value = handlexx
	End If

	Set ProcuraDLL = Nothing

End Sub

Public Sub BENEFICIARIO_OnExit()

	Dim msg As String
	SugerirIdadeBeneficiario
	msg = VerificarCC

	If (msg <> "") Then
		bsShowMessage(msg, "I")
	End If

End Sub

Public Sub CIDALTA_OnPopup(ShowPopup As Boolean)

  ShowPopup = False

  Dim OLEAutorizador As Object
  Dim handlexx As Long

  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")

  handlexx = OLEAutorizador.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", True, "")

  If handlexx <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CIDALTA").Value = handlexx
  End If

  Set OLEAutorizador = Nothing
End Sub

Public Sub CIDINTERNACAO_OnPopup(ShowPopup As Boolean)

  ' Leonardo Inicio 25/01/2001
  Dim OLEAutorizador As Object
  Dim handlexx As Long
  ShowPopup = False
  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
  handlexx = OLEAutorizador.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CIDINTERNACAO").Value = handlexx
  End If
  Set OLEAutorizador = Nothing
  ' Leonardo Fim 25/01/2001
End Sub

Public Sub CIDPRINCIPAL_OnPopup(ShowPopup As Boolean)
  ' Leonardo Inicio 25/01/2001

  Dim OLEAutorizador As Object
  Dim handlexx As Long
  ShowPopup = False
  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
  handlexx = OLEAutorizador.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CIDPRINCIPAL").Value = handlexx
  End If
  Set OLEAutorizador = Nothing
  ' Leonardo Fim 25/01/2001
End Sub

Public Sub DATAATENDIMENTO_OnExit()
  Dim vAux As Long
  Dim qParamGeral As Object
  Set qParamGeral = NewQuery

  qParamGeral.Clear
  qParamGeral.Add("SELECT SUGERIRIDADEBENEF FROM SAM_PARAMETROSPROCCONTAS")
  qParamGeral.Active = True

  If qParamGeral.FieldByName("SUGERIRIDADEBENEF").AsString = "S" Then
    If CurrentQuery.State <>1 Then
      If CurrentQuery.FieldByName("DATAATENDIMENTO").IsNull Then
        DataBase = ServerDate
      Else
        DataBase = CurrentQuery.FieldByName("DATAATENDIMENTO").AsDateTime
      End If
      SugerirIdadeBeneficiario
    End If
  End If
  Set qParamGeral = Nothing

End Sub

Public Sub DATAHORAINICIALCIRURGIA_OnExit()
  If Not CurrentQuery.FieldByName("DATAHORAINICIALCIRURGIA").IsNull Then
    If CurrentQuery.FieldByName("DATAHORAFINALCIRURGIA").IsNull Then
      CurrentQuery.FieldByName("DATAHORAFINALCIRURGIA").AsDateTime = Round(CurrentQuery.FieldByName("DATAHORAINICIALCIRURGIA").AsDateTime)
    End If
  End If
End Sub

Public Sub ENDERECOEXECUTOR_OnChange()

  SelecionaCnesEndPrest("Executor")

End Sub

Public Sub ENDERECOLOCALEXEC_OnChange()

  SelecionaCnesEndPrest("LocalExec")

End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vTexto As String

  vTexto = EVENTO.LocateText

  If ShowPopup = False Then
    Exit Sub
  End If

  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, vTexto)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

End Sub

Public Sub EXECUTOR_OnChange()
  ContaQtdeEndPrest ("Executor")
End Sub

Public Sub GRAUPARTICIPACAO_OnPopup(ShowPopup As Boolean)

  GRAUPARTICIPACAO.LocalWhere = VersaoTissDoPEG

End Sub

Public Sub GUIAAUXILIAR_OnExit()

	If(CurrentQuery.State = 2 Or CurrentQuery.State = 3)Then

		Dim INTERFACE As Object
		Dim pGuiaAuxiliar As String
		Dim msg As String

		Set INTERFACE = CreateBennerObject("SamPEgDigit.digitacao")

		msg = INTERFACE.FormataGuiaAuxiliar(CurrentSystem, CurrentQuery.FieldByName("GUIAAUXILIAR").AsString, pGuiaAuxiliar)

		If msg <> "" Then
			bsShowMessage(msg, "I")
		End If

		CurrentQuery.FieldByName("GUIAAUXILIAR").AsString = pGuiaAuxiliar

		CurrentQuery.FieldByName("GUIA").AsFloat = IIf(CurrentQuery.FieldByName("GUIAAUXILIAR").AsString = "", 0, LongHint(CurrentQuery.FieldByName("GUIAAUXILIAR").Value))

		Set INTERFACE = Nothing
	End If

End Sub

Public Sub HORAATENDIMENTO_OnExit()

	If(CurrentQuery.State >1)Then
		If CurrentQuery.FieldByName("GIHDATAHORAINTERNACAO").IsNull Then
			CurrentQuery.FieldByName("GIHDATAHORAINTERNACAO").AsDateTime = CurrentQuery.FieldByName("DATAATENDIMENTO").AsDateTime + _
																		   CurrentQuery.FieldByName("HORAATENDIMENTO").AsDateTime - _
																		   Round(CurrentQuery.FieldByName("HORAATENDIMENTO").AsDateTime)
	    End If
	End If

End Sub

Public Sub LOCALEXECUCAO_OnChange()
  ContaQtdeEndPrest("LocalExec")
End Sub

Public Sub LOCALEXECUCAO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", LOCALEXECUCAO.Text)' pelo CPF e todos
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("LOCALEXECUCAO").Value = vHandle
    LOCALEXECUCAO_OnChange
  End If
End Sub

Public Sub MOTIVOALTA_OnChange()
	Dim sqlMotivoSaida As Object
	Set sqlMotivoSaida = NewQuery

	sqlMotivoSaida.Clear
	sqlMotivoSaida.Add("SELECT HANDLE MOTIVOSAIDA        ")
	sqlMotivoSaida.Add("  FROM TIS_MOTIVOSAIDAINTERNACAO ")
	sqlMotivoSaida.Add(" WHERE MOTIVOALTA = :MOTIVOALTA  ")
	sqlMotivoSaida.Add("   AND                           ")
	sqlMotivoSaida.Add(VersaoTissDoPEG)

	sqlMotivoSaida.ParamByName("MOTIVOALTA").AsInteger = CurrentQuery.FieldByName("MOTIVOALTA").AsInteger
	sqlMotivoSaida.Active = True

	Dim sqlContador As Object
	Set sqlContador = NewQuery
	sqlContador.Add("SELECT COUNT(1) TOTAL            ")
	sqlContador.Add("  FROM TIS_MOTIVOSAIDAINTERNACAO ")
	sqlContador.Add(" WHERE MOTIVOALTA = :MOTIVOALTA  ")
	sqlContador.Add("   AND                           ")
	sqlContador.Add(VersaoTissDoPEG)

	sqlContador.ParamByName("MOTIVOALTA").AsInteger = CurrentQuery.FieldByName("MOTIVOALTA").AsInteger
	sqlContador.Active = True

	CurrentQuery.FieldByName("MOTIVOSAIDA").Value = IIf(sqlMotivoSaida.FieldByName("MOTIVOSAIDA").IsNull Or sqlContador.FieldByName("TOTAL").AsInteger > 1, Null, sqlMotivoSaida.FieldByName("MOTIVOSAIDA").AsInteger)

	Set sqlContador = Nothing

	Set sqlMotivoSaida = Nothing
End Sub

Public Sub MOTIVOSAIDA_OnChange()
	Dim sqlMotivoAlta As Object
	Set sqlMotivoAlta = NewQuery

	sqlMotivoAlta.Clear
	sqlMotivoAlta.Add("SELECT MOTIVOALTA                ")
	sqlMotivoAlta.Add("  FROM TIS_MOTIVOSAIDAINTERNACAO ")
	sqlMotivoAlta.Add(" WHERE HANDLE = :MOTIVOSAIDA     ")
	sqlMotivoAlta.Add("   AND                           ")
	sqlMotivoAlta.Add(VersaoTissDoPEG)

	sqlMotivoAlta.ParamByName("MOTIVOSAIDA").AsInteger = CurrentQuery.FieldByName("MOTIVOSAIDA").AsInteger
	sqlMotivoAlta.Active = True

	CurrentQuery.FieldByName("MOTIVOALTA").Value = IIf(sqlMotivoAlta.FieldByName("MOTIVOALTA").IsNull, Null, sqlMotivoAlta.FieldByName("MOTIVOALTA").AsInteger)
	Set sqlMotivoAlta = Nothing
End Sub

Public Sub NUMEROCONTINGENCIA_OnExit()
	Dim qAux As Object

	If CurrentQuery.State <>1 Then
		If(CurrentQuery.FieldByName("NUMEROCONTINGENCIA").AsString <>"")And(OldNumeroContingencia <>CurrentQuery.FieldByName("NUMEROCONTINGENCIA").AsString)Then
			Set qAux = NewQuery

			qAux.Active = False
			qAux.Clear
			qAux.Add("SELECT HANDLE                                         ")
			qAux.Add("  FROM SAM_AUTORIZ                                    ")
			qAux.Add(" WHERE UPPER(NUMEROCONTINGENCIA) = :NUMEROCONTINGENCIA")
			qAux.ParamByName("NUMEROCONTINGENCIA").AsString = CurrentQuery.FieldByName("NUMEROCONTINGENCIA").AsString
			qAux.Active = True
			' Preenche o número da autorização de acordo com o número de contingência
			If(Not qAux.EOF)Then
				CurrentQuery.FieldByName("AUTORIZACAO").AsInteger = qAux.FieldByName("HANDLE").AsInteger
				' Executa o procedimento para carregar os dados para o form
				HerdarAutorizacao
			End If

			qAux.Active = False
			Set qAux = Nothing
		End If

		OldNumeroContingencia = CurrentQuery.FieldByName("NUMEROCONTINGENCIA").AsString
	End If
End Sub

Public Sub RECEBEDOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", RECEBEDOR.Text)' pelo CPF e todos
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("RECEBEDOR").Value = vHandle
  End If
End Sub

Public Sub TABREGIMEPGTO_OnChanging(AllowChange As Boolean)
  bsShowMessage("Alteração não permitida", "I")
  AllowChange = False
End Sub

Public Sub SOLICITANTE_OnPopup(ShowPopup As Boolean)
  TABLE_BeforeEdit(ShowPopup)
  If ShowPopup = False Then
    Exit Sub
  End If
  '  If Len(SOLICITANTE.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", SOLICITANTE.LocateText)' pelo CPF e Solicitante
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("SOLICITANTE").Value = vHandle
  End If
  '  End If
End Sub

Public Sub EXECUTOR_OnPopup(ShowPopup As Boolean)
  TABLE_BeforeEdit(ShowPopup)
  If ShowPopup = False Then
    Exit Sub
  End If

  Dim vHandle As Long
  ShowPopup = False

  Dim vUtilizarConsultaCentral As Boolean
  Dim qParametros As Object
  Set qParametros = NewQuery
  qParametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
  qParametros.Active = True
  vUtilizarConsultaCentral = (qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString = "S")
  Set qParametros = Nothing

  If RecordHandleOfTable("SAM_TIPOPRESTADOR") > 0 And vUtilizarConsultaCentral Then
	Dim q1 As Object

	Set q1 = NewQuery

    q1.Clear
    q1.Add("SELECT FORMACAOPRESTADOR FROM SAM_TIPOPRESTADOR WHERE HANDLE = :HTIPOPRESTADOR")
    q1.ParamByName("HTIPOPRESTADOR").AsInteger = RecordHandleOfTable("SAM_TIPOPRESTADOR")
    q1.Active = True

    vFormacaoPrestador = q1.FieldByName("FORMACAOPRESTADOR").AsString

	vHandle = ProcuraPrestador("C", "E", EXECUTOR.LocateText)

	Set q1 = Nothing
  Else
    ShowPopup = False
    vHandle = ProcuraPrestador("C", "T", EXECUTOR.LocateText)' pelo CPF e executor
  End If

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EXECUTOR").Value = vHandle
    EXECUTOR_OnChange
  End If

End Sub

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  TABLE_BeforeEdit(ShowPopup)
  If ShowPopup = False Then
    Exit Sub
  End If

  Dim vHandle As Long
  ShowPopup = False

  Dim vRecebedor As Long
  Dim qRecebedor As Object
  Dim InterfaceBenef As Object
  'Separação da Interface da regra de negocio para consulta de Beneficiários
  Set InterfaceBenef =CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")
  Dim qParametros As Object
  Dim vUtilizarConsultaCentral As Boolean

  If (Not CurrentQuery.FieldByName("RECEBEDOR").IsNull) Then
    vRecebedor = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
  Else
    Set qRecebedor = NewQuery
    qRecebedor.Add("SELECT P.RECEBEDOR RECEBEDORPEG ")
    qRecebedor.Add("  FROM SAM_PEG P ")
    qRecebedor.Add(" WHERE P.HANDLE = :PEG ")
    qRecebedor.ParamByName("PEG").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
    qRecebedor.Active = True
    vRecebedor = qRecebedor.FieldByName("RECEBEDORPEG").AsInteger
    Set qRecebedor = Nothing
  End If

  Set qParametros = NewQuery
  qParametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
  qParametros.Active = True
  vUtilizarConsultaCentral = (qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString = "S")
  Set qParametros = Nothing

  If (CurrentQuery.FieldByName("DATAATENDIMENTO").IsNull) Then
    InterfaceBenef.AlteraDataAtend(ServerDate, vRecebedor)
  Else
    InterfaceBenef.AlteraDataAtend(CurrentQuery.FieldByName("DATAATENDIMENTO").AsDateTime, vRecebedor)
  End If


  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT P.TABREGIMEPGTO, P.BENEFICIARIO")
  SQL.Add("  FROM SAM_PEG  P ")
  SQL.Add(" WHERE P.HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
  SQL.Active = True

  If SQL.FieldByName("TABREGIMEPGTO").AsInteger = 2 Then
    If SQL.FieldByName("BENEFICIARIO").AsInteger >= 0 Then
      If (vUtilizarConsultaCentral) Then
        vHandle = InterfaceBenef.FiltroTitular(CurrentSystem, 1, BENEFICIARIO.LocateText, SQL.FieldByName("BENEFICIARIO").AsInteger)
      Else
        Dim vBenefReembolso As Long
        vBenefReembolso = SQL.FieldByName("BENEFICIARIO").AsInteger
        vHandle = ProcuraBeneficiarioAtivoReembolso(False,ServerDate,BENEFICIARIO.LocateText, vBenefReembolso, True)
      End If

    Else
      vHandle =ProcuraBeneficiarioAtivo(False,ServerDate,BENEFICIARIO.LocateText)
    End If
  Else
    vHandle =ProcuraBeneficiarioAtivo(False,ServerDate,BENEFICIARIO.LocateText)
  End If

  Set InterfaceBenef = Nothing

  If (vHandle <> 0) Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
End Sub

Public Sub TIPOACOMODACAO_OnChange()
  TIPOACOMODACAO.LocalWhere = " ACOMODACAO IN (SELECT HANDLE FROM SAM_ACOMODACAO) "
  If (CurrentQuery.FieldByName("TIPOACOMODACAO").AsInteger > 0) And (ACOMODACAO.ReadOnly) Then
    Dim Q As Object
    Set Q = NewQuery
    Q.Add(" SELECT ACOMODACAO FROM TIS_TIPOACOMODACAO WHERE HANDLE = :HANDLE ")
    Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOACOMODACAO").AsInteger
    Q.Active = True
    If (Q.FieldByName("ACOMODACAO").AsInteger > 0) Then
      CurrentQuery.FieldByName("ACOMODACAO").AsInteger = Q.FieldByName("ACOMODACAO").AsInteger
    End If
    Set Q = Nothing
  End If
End Sub

' Tabela -----------------------------------------------

Public Sub TABLE_AfterCommitted()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_AfterCommitted(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_AfterDelete()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_AfterDelete(CurrentSystem, CurrentQuery.TQuery, IIf(RecordHandleOfTable("SAM_PEG")=-1,SessionVar("HPEG"),RecordHandleOfTable("SAM_PEG")), vSituacaoPeg)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_AfterEdit()

	CIDALTA.AnyLevel = True
	CIDINTERNACAO.AnyLevel = True
	CIDPRINCIPAL.AnyLevel = True

	OldExecutor = CurrentQuery.FieldByName("EXECUTOR").AsInteger

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_AfterEdit(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_AfterInsert()

	Dim vbBeneficiario      As Boolean
    Dim vbRegimeAtendimento As Boolean
    Dim vbLocalAtendimento  As Boolean
	Dim viPeg               As Integer

	vbBeneficiario      = BENEFICIARIO.ReadOnly
	vbRegimeAtendimento = REGIMEATENDIMENTO.ReadOnly
	vbLocalAtendimento  = LOCALATENDIMENTO.ReadOnly

	viPeg = IIf(RecordHandleOfTable("SAM_PEG")=-1, SessionVar("HPEG"), RecordHandleOfTable("SAM_PEG"))

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_AfterInsert(CurrentSystem, CurrentQuery.TQuery, viPeg, gModeloGuia, cancontinue, vbBeneficiario, vbRegimeAtendimento, vbLocalAtendimento)

	Set vDllBSPro006 = Nothing

	BENEFICIARIO.ReadOnly      = vbBeneficiario
	REGIMEATENDIMENTO.ReadOnly = vbRegimeAtendimento
	LOCALATENDIMENTO.ReadOnly  = vbLocalAtendimento

End Sub

Public Sub TABLE_AfterPost()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_AfterPost(CurrentSystem, CurrentQuery.TQuery, viState, vSituacaoPeg, vSituacaoGuia)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_AfterScroll()

	REGIMEINTERNACAO.LocalWhere    = VersaoTissDoPEG
	REGIMEINTERNACAO.WebLocalWhere = " A." + VersaoTissDoPEG
	TIPOACOMODACAO.LocalWhere      = VersaoTissDoPEG
	TIPOACOMODACAO.WebLocalWhere   = " A." + VersaoTissDoPEG
	MOTIVOSAIDA.LocalWhere         = VersaoTissDoPEG
	MOTIVOSAIDA.WebLocalWhere      = " A." + VersaoTissDoPEG

	If RecordHandleOfTable("SAM_PEG") <> -1 Then
		SessionVar("HPEG") = CStr(RecordHandleOfTable("SAM_PEG"))
	End If

	MostraCamposLeiaute

	old_modeloguia = CurrentQuery.FieldByName("MODELOGUIA").AsFloat

	Dim vbBotaoOdontogramaVisible        As Boolean
	Dim vbBotaoOdontogramaEnable         As Boolean
	Dim vbBotaoAprovarValorEnable        As Boolean
	Dim vbBotaoDevolverGuiaEnable        As Boolean
	Dim vbBotaoReprocessarGuiaEnable     As Boolean
	Dim vbBotaoConferidaEnable           As Boolean
	Dim vbBotaoContaTerceiroEnable       As Boolean
	Dim vbBotaoDigitarEnable             As Boolean
	Dim vbBotaoGlosaTotalEnable          As Boolean
	Dim vbBotaoIncluirPrestadorEnable    As Boolean
	Dim vbBotaoNaoFinanciarEnable        As Boolean
	Dim vbBotaoPFIntegralEnable          As Boolean
	Dim vbBotaoSipEnable                 As Boolean
	Dim vbBotaoCancelarProvisaoEnable    As Boolean
	Dim vbBotaoAtualizarValorEnable      As Boolean
	Dim vbModeloGuia                     As Boolean
	Dim vbBotaoAlterarBeneficiarioEnable As Boolean
	Dim vbBotaoGuiaOriginalEnable        As Boolean
	Dim vbModeloGuiaWebLocalWhere        As String
	Dim vsTitularGuia                    As String
	Dim vsRotulo1                        As String
	Dim vsRotulo2                        As String
	Dim vsRotulo3                        As String
	Dim vsRotulo4                        As String
	Dim vbAbatePFEventoReadOnly          As Boolean
	Dim vbTableReadOnly                  As Boolean

	vbBotaoOdontogramaVisible        = BOTAOODONTOGRAMA.Visible
	vbBotaoOdontogramaEnable         = BOTAOODONTOGRAMA.Enabled
	vbBotaoAprovarValorEnable        = BOTAOAPROVARVALOR.Enabled
	vbBotaoDevolverGuiaEnable        = BOTAODEVOLVERGUIA.Enabled
	vbBotaoReprocessarGuiaEnable     = BOTAOREPROCESSARGUIA.Enabled
	vbBotaoConferidaEnable           = BOTAOCONFERIDA.Enabled
	vbBotaoContaTerceiroEnable       = BOTAOCONTATERCEIRO.Enabled
	vbBotaoDigitarEnable             = BOTAODIGITAR.Enabled
	vbBotaoGlosaTotalEnable          = BOTAOGLOSATOTAL.Enabled
	vbBotaoIncluirPrestadorEnable    = BOTAOINCLUIRPRESTADOR.Enabled
	vbBotaoNaoFinanciarEnable        = BOTAONAOFINANCIAR.Enabled
	vbBotaoPFIntegralEnable          = BOTAOPFINTEGRAL.Enabled
	vbBotaoSipEnable                 = BOTAOSIP.Enabled
	vbBotaoCancelarProvisaoEnable    = BOTAOCANCELARPROVISAO.Enabled
	vbBotaoAtualizarValorEnable      = BOTAOATUALIZARVALOR.Enabled
	vbBotaoAlterarBeneficiarioEnable = BOTAOALTERARBENEFICIARIO.Enabled
	vbBotaoGuiaOriginalEnable        = BOTAOGUIAORIGINAL.Enabled
	vbModeloGuia                     = MODELOGUIA.ReadOnly
	vbModeloGuiaWebLocalWhere        = MODELOGUIA.WebLocalWhere
	vsTitularGuia                    = TITULARGUIA.Text
	vsRotulo1                        = ROTULO1.Text
	vsRotulo2                        = ROTULO2.Text
	vsRotulo3                        = ROTULO3.Text
	vsRotulo4                        = ROTULO4.Text
	vbAbatePFEventoReadOnly          = ABATEPFDOEVENTO.ReadOnly
	vbTableReadOnly                  = TableReadOnly

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_AfterScroll(CurrentSystem, _
	                               CurrentQuery.TQuery, _
								   vbBotaoOdontogramaVisible, _
								   vbBotaoOdontogramaEnable, _
								   vbBotaoAprovarValorEnable, _
								   vbBotaoDevolverGuiaEnable, _
								   vbBotaoReprocessarGuiaEnable, _
								   vbBotaoConferidaEnable, _
								   vbBotaoContaTerceiroEnable, _
								   vbBotaoDigitarEnable, _
								   vbBotaoGlosaTotalEnable, _
								   vbBotaoIncluirPrestadorEnable, _
								   vbBotaoNaoFinanciarEnable, _
								   vbBotaoPFIntegralEnable, _
								   vbBotaoSipEnable, _
								   vbBotaoCancelarProvisaoEnable, _
								   vbBotaoAtualizarValorEnable, _
								   vbBotaoAlterarBeneficiarioEnable, _
								   vbBotaoGuiaOriginalEnable, _
								   vbModeloGuia, _
								   vbModeloGuiaWebLocalWhere, _
								   vsTitularGuia, _
								   vsRotulo1, _
								   vsRotulo2, _
								   vsRotulo3, _
								   vsRotulo4, _
								   vbAbatePFEventoReadOnly, _
 								   OLDBENEFICIARIO, _
 								   vbTableReadOnly, _
								   vSituacaoPeg, _
								   vSituacaoGuia)

	Set vDllBSPro006 = Nothing

	BOTAOODONTOGRAMA.Visible		= vbBotaoOdontogramaVisible
	BOTAOODONTOGRAMA.Enabled		= vbBotaoOdontogramaEnable
	BOTAOAPROVARVALOR.Enabled		= vbBotaoAprovarValorEnable
	BOTAODEVOLVERGUIA.Enabled		= vbBotaoDevolverGuiaEnable
	BOTAOREPROCESSARGUIA.Enabled	= vbBotaoReprocessarGuiaEnable
	BOTAOCONFERIDA.Enabled			= vbBotaoConferidaEnable
	BOTAOCONTATERCEIRO.Enabled		= vbBotaoContaTerceiroEnable
	BOTAODIGITAR.Enabled			= vbBotaoDigitarEnable
	BOTAOGLOSATOTAL.Enabled			= vbBotaoGlosaTotalEnable
	BOTAOINCLUIRPRESTADOR.Enabled	= vbBotaoIncluirPrestadorEnable
	BOTAONAOFINANCIAR.Enabled		= vbBotaoNaoFinanciarEnable
	BOTAOPFINTEGRAL.Enabled			= vbBotaoPFIntegralEnable
	BOTAOSIP.Enabled				= vbBotaoSipEnable
	BOTAOCANCELARPROVISAO.Enabled	= vbBotaoCancelarProvisaoEnable
	BOTAOATUALIZARVALOR.Enabled     = vbBotaoAtualizarValorEnable
	BOTAOALTERARBENEFICIARIO.Enabled= vbBotaoAlterarBeneficiarioEnable
	BOTAOGUIAORIGINAL.Enabled       = vbBotaoGuiaOriginalEnable
	MODELOGUIA.ReadOnly				= vbModeloGuia
	MODELOGUIA.WebLocalWhere		= vbModeloGuiaWebLocalWhere
	TITULARGUIA.Text				= vsTitularGuia
	ROTULO1.Text					= vsRotulo1
	ROTULO2.Text					= vsRotulo2
	ROTULO3.Text					= vsRotulo3
	ROTULO4.Text					= vsRotulo4
	ABATEPFDOEVENTO.ReadOnly 		= vbAbatePFEventoReadOnly
	TableReadOnly        			= vbTableReadOnly
End Sub

Public Sub TABLE_BeforeDelete(cancontinue As Boolean)

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_BeforeDelete(CurrentSystem, CurrentQuery.TQuery, cancontinue)

	Set vDllBSPro006 = Nothing

	MostraCamposLeiaute

End Sub

Public Sub TABLE_BeforeEdit(cancontinue As Boolean)

    'Não era verificado se a agencia estava ativa ou inativa, exibindo assim todas as agências.
	If WebMode Then
		AGENCIA.WebLocalWhere = "SITUACAO = 'A'"
	ElseIf VisibleMode Then
		AGENCIA.LocalWhere = "SITUACAO = 'A'"
	End If

	Dim vbBotaoReprocessarGuia As Boolean
	vbBotaoReprocessarGuia = BOTAOREPROCESSARGUIA.Enabled

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_BeforeEdit(CurrentSystem, CurrentQuery.TQuery, cancontinue, vbBotaoReprocessarGuia)

	Set vDllBSPro006 = Nothing

	BOTAOREPROCESSARGUIA.Enabled = vbBotaoReprocessarGuia

	OldAutorizacao           = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
	OldNumeroContingencia    = CurrentQuery.FieldByName("NUMEROCONTINGENCIA").AsString
    Regime_Atend             = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
    Data_Atend               = CurrentQuery.FieldByName("DATAATENDIMENTO").AsDateTime

End Sub

Public Sub TABLE_BeforeInsert(cancontinue As Boolean)

	CIDALTA.AnyLevel = True
	CIDINTERNACAO.AnyLevel = True
	CIDPRINCIPAL.AnyLevel = True

	If WebMode Then
		AGENCIA.WebLocalWhere = "SITUACAO = 'A'"
	ElseIf VisibleMode Then
		AGENCIA.LocalWhere = "SITUACAO = 'A'"
	End If

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_BeforeInsert(CurrentSystem, CurrentQuery.TQuery, RecordHandleOfTable("SAM_PEG"), cancontinue, gModeloGuia)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_BeforePost(cancontinue As Boolean)

	Dim vbMotivoAlta As Boolean
	Dim vbCanContinue As Boolean

	vbMotivoAlta  = MOTIVOALTA.ReadOnly
	vbCanContinue = cancontinue

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_BeforePost(CurrentSystem, CurrentQuery.TQuery, vSituacaoGuia, vbMotivoAlta, OLDBENEFICIARIO, viState, vbCanContinue)

	Set vDllBSPro006 = Nothing

	MOTIVOALTA.ReadOnly = vbMotivoAlta
	cancontinue         = vbCanContinue

End Sub

Public Sub TABLE_NewRecord()

	REGIMEINTERNACAO.LocalWhere = VersaoTissDoPEG
	TIPOACOMODACAO.LocalWhere   = VersaoTissDoPEG

	Dim vbLocalAtendimento As Boolean
	vbLocalAtendimento = LOCALATENDIMENTO.ReadOnly

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.TABLE_NewRecord(CurrentSystem, CurrentQuery.TQuery, RecordHandleOfTable("SAM_PEG"), gModeloGuia, vSituacaoPeg, vbLocalAtendimento)

	Set vDllBSPro006 = Nothing

	LOCALATENDIMENTO.ReadOnly  = vbLocalAtendimento

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, cancontinue As Boolean)

	Select Case CommandID
      Case "BOTAOAPROVARVALOR"
        BOTAOAPROVARVALOR_OnClick
      Case "BOTAOCONFERIDA"
        BOTAOCONFERIDA_OnClick
	  'Case "BOTAODEVOLVERGUIA"
	  '	BOTAODEVOLVERGUIA_OnClick
      Case "BOTAODIGITAR"
        BOTAODIGITAR_OnClick
      Case "BOTAOGLOSATOTAL"
        BOTAOGLOSATOTAL_OnClick
      Case "BOTAOTROCARLEIAUTE"
        BOTAOTROCARLEIAUTE_OnClick
      Case "BOTAODEVOLVERGUIA"
        BOTAODEVOLVERGUIA_OnClick
      Case "BOTAOREPROCESSARGUIA"
        BOTAOREPROCESSARGUIA_OnClick
      Case "BOTAOODONTOGRAMA"
        BOTAOODONTOGRAMA_OnClick
      Case "BOTAOPFINTEGRAL"
        BOTAOPFINTEGRAL_OnClick
      Case "BOTAOVERIFICAMONITORAMENTO"
        IncluiSessionVarMonitoramento
	End Select

End Sub

Public Sub TABLE_UpdateRequired()

	TratarCaracteristicasAtendimento

	If WebMode Then
		Dim SQL As Object
		Set SQL = NewQuery
		SQL.Add("SELECT LOCALATENDIMENTO, CONDICAOATENDIMENTO, FINALIDADEATENDIMENTO, OBJETIVOTRATAMENTO, REGIMEATENDIMENTO, TIPOTRATAMENTO FROM SAM_TIPOGUIA_MDGUIA WHERE HANDLE=:HANDLE")
		SQL.ParamByName("HANDLE").AsInteger =  CurrentQuery.FieldByName("MODELOGUIA").AsInteger
		SQL.Active = True

		If CurrentQuery.FieldByName("LOCALATENDIMENTO").IsNull And _
			Not SQL.FieldByName("LOCALATENDIMENTO").IsNull Then
			CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger = SQL.FieldByName("LOCALATENDIMENTO").AsInteger
		End If

		If CurrentQuery.FieldByName("CONDICAOATENDIMENTO").IsNull And _
			Not SQL.FieldByName("CONDICAOATENDIMENTO").IsNull Then
			CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger = SQL.FieldByName("CONDICAOATENDIMENTO").AsInteger
		End If

		If CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").IsNull And _
			Not SQL.FieldByName("FINALIDADEATENDIMENTO").IsNull Then
			CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger = SQL.FieldByName("FINALIDADEATENDIMENTO").AsInteger
		End If

		If CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").IsNull And _
			Not SQL.FieldByName("OBJETIVOTRATAMENTO").IsNull Then
			CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger = SQL.FieldByName("OBJETIVOTRATAMENTO").AsInteger
		End If

		If CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull And _
			Not SQL.FieldByName("REGIMEATENDIMENTO").IsNull Then
			CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger = SQL.FieldByName("REGIMEATENDIMENTO").AsInteger
		End If

		If CurrentQuery.FieldByName("TIPOTRATAMENTO").IsNull And _
			Not SQL.FieldByName("TIPOTRATAMENTO").IsNull Then
			CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger = SQL.FieldByName("TIPOTRATAMENTO").AsInteger
		End If

		Set SQL = Nothing
	End If
End Sub

' Funções --------------------------------------

Public Sub EscondeCamposLeiaute
  'esses são os campos que estariam no panel da SamPegDigit.dll que é montado conforme o modelo de guia
  If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Or CurrentQuery.FieldByName("SITUACAO").AsString = "4" Or _
     WebMode Then
    Exit Sub
  End If

  ORDEM.ReadOnly = True
  GUIA.ReadOnly = True
  CONDICAOATENDIMENTO.ReadOnly = True
  HORAATENDIMENTO.ReadOnly = True
  DATAATENDIMENTO.ReadOnly = True
  GODATAINICIAL.ReadOnly = True
  GODATAFINAL.ReadOnly = True
  GIHDATAHORAINTERNACAO.ReadOnly = True
  GIHDATAHORAALTA.ReadOnly = True
  DATAHORAINICIALCIRURGIA.ReadOnly = True
  DATAHORAFINALCIRURGIA.ReadOnly = True
  DVCARTAO.ReadOnly = True
  PERCENTUALDESCONTO.ReadOnly = True
  EXECUTOR.ReadOnly = True
  BENEFICIARIO.ReadOnly = True
  MATRICULA.ReadOnly = True
  SOLICITANTE.ReadOnly = True
  TIPOTRATAMENTO.ReadOnly = True
  REGIMEATENDIMENTO.ReadOnly = True
  REVISOR.ReadOnly = True
  MOTIVOALTA.ReadOnly = True
  FINALIDADEATENDIMENTO.ReadOnly = True
  EVENTO.ReadOnly = True
  CIDINTERNACAO.ReadOnly = True
  CIDALTA.ReadOnly = True
  CIDPRINCIPAL.ReadOnly = True
  LOCALEXECUCAO.ReadOnly = True
  IDADEBENEFICIARIO.ReadOnly = True
  VALORAPRESENTADO.ReadOnly = True
  LOCALATENDIMENTO.ReadOnly = True
  OBJETIVOTRATAMENTO.ReadOnly = True

  INDICACAOCLINICA.ReadOnly = True
  TIPOCONSULTA.ReadOnly = True
  TIPODOENCA.ReadOnly = True
  TEMPODOENCA.ReadOnly = True
  UNIDADETEMPODOENCA.ReadOnly = True
  INDICADORDEACIDENTE.ReadOnly = True
  TIPOSAIDACONSULTA.ReadOnly = True
  TIPOSAIDASPSADT.ReadOnly = True
  TIPOFATURAMENTO.ReadOnly = True
  TIPOATENDIMENTO.ReadOnly = True
  TIPOINTERNACAO.ReadOnly = True
  QTDDIARIASUTI.ReadOnly = True
  DECLARACAOOBITO.ReadOnly = True
  CIDOBITO.ReadOnly = True
  ACOMODACAO.ReadOnly = True
  TIPOACOMODACAO.ReadOnly = True
  CID2.ReadOnly = True
  CID3.ReadOnly = True
  CID4.ReadOnly = True
  DATAEMISSAO.ReadOnly = True
  SENHATISS.ReadOnly = True
  EVENTOSUS.ReadOnly = True


End Sub

Public Sub MostraCamposLeiaute

  If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Or CurrentQuery.FieldByName("SITUACAO").AsString = "4" Or _
     WebMode Then
    Exit Sub
  End If

  If old_modeloguia = CurrentQuery.FieldByName("MODELOGUIA").AsFloat Then
    Exit Sub
  End If

  Dim vAux As String

  ListaCamposLeiaute IIf(gModeloGuia = 0, CurrentQuery.FieldByName("MODELOGUIA").AsInteger, gModeloGuia), 2

  EscondeCamposLeiaute

  vAux = UserVar("CAMPOS_LEIAUTE_GUIA")

  If(InStr(vAux, "|" + "ORDEM")>0)And(CurrentQuery.RequestLive)Then ORDEM.ReadOnly = False
  If(InStr(vAux, "|" + "OBJETIVOTRATAMENTO")>0)And(CurrentQuery.RequestLive)Then OBJETIVOTRATAMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "LOCALATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then LOCALATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "GUIA")>0)And(CurrentQuery.RequestLive)Then GUIA.ReadOnly = False
  If(InStr(vAux, "|" + "CONDICAOATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then CONDICAOATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "HORAATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then HORAATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "DATAATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then DATAATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "GODATAINICIAL")>0)And(CurrentQuery.RequestLive)Then GODATAINICIAL.ReadOnly = False
  If(InStr(vAux, "|" + "GODATAFINAL")>0)And(CurrentQuery.RequestLive)Then GODATAFINAL.ReadOnly = False
  If(InStr(vAux, "|" + "GIHDATAHORAINTERNACAO")>0)And(CurrentQuery.RequestLive)Then GIHDATAHORAINTERNACAO.ReadOnly = False
  If(InStr(vAux, "|" + "GIHDATAHORAALTA")>0)And(CurrentQuery.RequestLive)Then GIHDATAHORAALTA.ReadOnly = False
  If(InStr(vAux, "|" + "DATAHORAINICIALCIRURGIA")>0)And(CurrentQuery.RequestLive)Then DATAHORAINICIALCIRURGIA.ReadOnly = False
  If(InStr(vAux, "|" + "DATAHORAFINALCIRURGIA")>0)And(CurrentQuery.RequestLive)Then DATAHORAFINALCIRURGIA.ReadOnly = False
  If(InStr(vAux, "|" + "DVCARTAO")>0)And(CurrentQuery.RequestLive)Then DVCARTAO.ReadOnly = False
  If(InStr(vAux, "|" + "PERCENTUALDESCONTO")>0)And(CurrentQuery.RequestLive)Then PERCENTUALDESCONTO.ReadOnly = False
  If(InStr(vAux, "|" + "SOLICITANTE")>0)And(CurrentQuery.RequestLive)Then SOLICITANTE.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOTRATAMENTO")>0)And(CurrentQuery.RequestLive)Then TIPOTRATAMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "REGIMEATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then REGIMEATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "REVISOR")>0)And(CurrentQuery.RequestLive)Then REVISOR.ReadOnly = False
  If(InStr(vAux, "|" + "MOTIVOALTA")>0)And(CurrentQuery.RequestLive)Then MOTIVOALTA.ReadOnly = False
  If(InStr(vAux, "|" + "FINALIDADEATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then FINALIDADEATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "EVENTO")>0)And(CurrentQuery.RequestLive)Then EVENTO.ReadOnly = False
  If(InStr(vAux, "|" + "EXECUTOR")>0)And(CurrentQuery.RequestLive)Then EXECUTOR.ReadOnly = False
  If(InStr(vAux, "|" + "CIDINTERNACAO")>0)And(CurrentQuery.RequestLive)Then CIDINTERNACAO.ReadOnly = False
  If(InStr(vAux, "|" + "CIDALTA")>0)And(CurrentQuery.RequestLive)Then CIDALTA.ReadOnly = False
  If(InStr(vAux, "|" + "CIDPRINCIPAL")>0)And(CurrentQuery.RequestLive)Then CIDPRINCIPAL.ReadOnly = False
  If(InStr(vAux, "|" + "LOCALEXECUCAO")>0)And(CurrentQuery.RequestLive)Then LOCALEXECUCAO.ReadOnly = False
  If(InStr(vAux, "|" + "IDADEBENEFICIARIO")>0)And(CurrentQuery.RequestLive)Then IDADEBENEFICIARIO.ReadOnly = False
  If(InStr(vAux, "|" + "BENEFICIARIO")>0)And(CurrentQuery.RequestLive)Then BENEFICIARIO.ReadOnly = False
  If(InStr(vAux, "|" + "VALORAPRESENTADO")>0)And(CurrentQuery.RequestLive)Then VALORAPRESENTADO.ReadOnly = False


  If(InStr(vAux, "|" + "INDICACAOCLINICA")>0)And(CurrentQuery.RequestLive)Then INDICACAOCLINICA.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOCONSULTA")>0)And(CurrentQuery.RequestLive)Then TIPOCONSULTA.ReadOnly = False
  If(InStr(vAux, "|" + "TIPODOENCA")>0)And(CurrentQuery.RequestLive)Then TIPODOENCA.ReadOnly = False
  If(InStr(vAux, "|" + "TEMPODOENCA")>0)And(CurrentQuery.RequestLive)Then TEMPODOENCA.ReadOnly = False
  If(InStr(vAux, "|" + "UNIDADETEMPODOENCA")>0)And(CurrentQuery.RequestLive)Then UNIDADETEMPODOENCA.ReadOnly = False
  If(InStr(vAux, "|" + "INDICADORDEACIDENTE")>0)And(CurrentQuery.RequestLive)Then INDICADORDEACIDENTE.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOSAIDACONSULTA")>0)And(CurrentQuery.RequestLive)Then TIPOSAIDACONSULTA.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOSAIDASPSADT")>0)And(CurrentQuery.RequestLive)Then TIPOSAIDASPSADT.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOFATURAMENTO")>0)And(CurrentQuery.RequestLive)Then TIPOFATURAMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then TIPOATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOINTERNACAO")>0)And(CurrentQuery.RequestLive)Then TIPOINTERNACAO.ReadOnly = False
  If(InStr(vAux, "|" + "QTDDIARIASUTI")>0)And(CurrentQuery.RequestLive)Then QTDDIARIASUTI.ReadOnly = False
  If(InStr(vAux, "|" + "DECLARACAOOBITO")>0)And(CurrentQuery.RequestLive)Then DECLARACAOOBITO.ReadOnly = False
  If(InStr(vAux, "|" + "CIDOBITO")>0)And(CurrentQuery.RequestLive)Then CIDOBITO.ReadOnly = False
  If(InStr(vAux, "|" + "ACOMODACAO")>0)And(CurrentQuery.RequestLive)Then ACOMODACAO.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOACOMODACAO")>0)And(CurrentQuery.RequestLive)Then TIPOACOMODACAO.ReadOnly = False
  If(InStr(vAux, "|" + "CID2")>0)And(CurrentQuery.RequestLive)Then CID2.ReadOnly = False
  If(InStr(vAux, "|" + "CID3")>0)And(CurrentQuery.RequestLive)Then CID3.ReadOnly = False
  If(InStr(vAux, "|" + "CID4")>0)And(CurrentQuery.RequestLive)Then CID4.ReadOnly = False
  If(InStr(vAux, "|" + "DATAEMISSAO")>0)And(CurrentQuery.RequestLive)Then DATAEMISSAO.ReadOnly = False
  If(InStr(vAux, "|" + "SENHATISS")>0)And(CurrentQuery.RequestLive)Then SENHATISS.ReadOnly = False
  If(InStr(vAux, "|" + "EVENTOSUS")>0)And(CurrentQuery.RequestLive)Then EVENTOSUS.ReadOnly = False


End Sub

Public Sub CaracteristicasTIPOINTERNACAO()
  Dim q1 As Object

  Set q1 = NewQuery

  q1.Clear
  q1.Add("SELECT TA.FINALIDADEATENDIMENTO,                 ")
  q1.Add("       TA.LOCALATENDIMENTO,                      ")
  q1.Add("       TA.REGIMEATENDIMENTO,                     ")
  q1.Add("       TA.TIPOTRATAMENTO,                        ")
  q1.Add("       TA.OBJETIVOTRATAMENTO                     ")
  q1.Add("  FROM TIS_TIPOINTERNACAO TA                     ")
  q1.Add(" WHERE TA.CODIGO = :CODIGO                       ")
  q1.Add("   And TA.VERSAOTISS In (Select MAX (HANDLE)        ")
  q1.Add("                           FROM TIS_VERSAO          ")
  q1.Add("                          WHERE ATIVODESKTOP = 'S') ")

  If Not CurrentQuery.FieldByName("TIPOINTERNACAO").IsNull Then
    q1.ParamByName("CODIGO").AsString = CurrentQuery.FieldByName("TIPOINTERNACAO").AsString
    q1.Active = True
    If (Not q1.FieldByName("FINALIDADEATENDIMENTO").IsNull) And _
       (FINALIDADEATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger = q1.FieldByName("FINALIDADEATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("LOCALATENDIMENTO").IsNull) And _
       (LOCALATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger = q1.FieldByName("LOCALATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("REGIMEATENDIMENTO").IsNull) And _
       (REGIMEATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger = q1.FieldByName("REGIMEATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("TIPOTRATAMENTO").IsNull) And _
       (TIPOTRATAMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger = q1.FieldByName("TIPOTRATAMENTO").AsInteger
    End If
    If (Not q1.FieldByName("OBJETIVOTRATAMENTO").IsNull) And _
       (OBJETIVOTRATAMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger = q1.FieldByName("OBJETIVOTRATAMENTO").AsInteger
    End If
  End If
  Set q1 = Nothing
End Sub

Public Sub CaracteristicasTIPOATENDIMENTO()
  Dim q1 As Object
  Dim q2 As Object
  Dim vCaraterSolicitacao As String

  Set q1 = NewQuery
  Set q2 = NewQuery

  q1.Clear
  q1.Add("SELECT TA.FINALIDADEATENDIMENTO,                 ")
  q1.Add("       TA.LOCALATENDIMENTO,                      ")
  q1.Add("       TA.REGIMEATENDIMENTO,                     ")
  q1.Add("       TA.TIPOTRATAMENTO,                        ")
  q1.Add("       TA.OBJETIVOTRATAMENTO                     ")
  q1.Add("  FROM TIS_TIPOATENDIMENTO    TA,                ")
  q1.Add("       TIS_CARATERATENDIMENTO CS                 ")
  q1.Add(" WHERE TA.CODIGO = :CODIGO                       ")
  q1.Add("   AND CS.CODIGO = :CARATERSOLICITACAO           ")
  q1.Add("   AND TA.CARATERSOLICITACAO = CS.HANDLE         ")
  q1.Add("   And TA.VERSAOTISS In (Select MAX (HANDLE)        ")
  q1.Add("                           FROM TIS_VERSAO          ")
  q1.Add("                          WHERE ATIVODESKTOP = 'S') ")

  q2.Clear
  q2.Add(" SELECT URGENCIA FROM SAM_CONDATENDIMENTO WHERE HANDLE = :HANDLE ")
  q2.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger
  q2.Active = True
  If (q2.FieldByName("URGENCIA").AsString = "S") Then
    vCaraterSolicitacao = "U"
  Else
    vCaraterSolicitacao = "E"
  End If


  If Not CurrentQuery.FieldByName("TIPOATENDIMENTO").IsNull Then
    q1.ParamByName("CODIGO").AsString = CurrentQuery.FieldByName("TIPOATENDIMENTO").AsString
    q1.ParamByName("CARATERSOLICITACAO").AsString = vCaraterSolicitacao
    q1.Active = True
    If (Not q1.FieldByName("FINALIDADEATENDIMENTO").IsNull) And _
       (FINALIDADEATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger = q1.FieldByName("FINALIDADEATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("LOCALATENDIMENTO").IsNull) And _
       (LOCALATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger = q1.FieldByName("LOCALATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("REGIMEATENDIMENTO").IsNull) And _
       (REGIMEATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger = q1.FieldByName("REGIMEATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("TIPOTRATAMENTO").IsNull) And _
       (TIPOTRATAMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger = q1.FieldByName("TIPOTRATAMENTO").AsInteger
    End If
    If (Not q1.FieldByName("OBJETIVOTRATAMENTO").IsNull) And _
       (OBJETIVOTRATAMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger = q1.FieldByName("OBJETIVOTRATAMENTO").AsInteger
    End If
  End If
  Set q1 = Nothing
End Sub

Public Sub CaracteristicasINDICADORDEACIDENTE()
  Dim q1 As Object

  Set q1 = NewQuery

  q1.Clear
  q1.Add("SELECT FINALIDADEATENDIMENTO,                    ")
  q1.Add("       LOCALATENDIMENTO,                         ")
  q1.Add("       REGIMEATENDIMENTO,                        ")
  q1.Add("       TIPOTRATAMENTO,                           ")
  q1.Add("       OBJETIVOTRATAMENTO                        ")
  q1.Add("  FROM TIS_INDICADORDEACIDENTE                   ")
  q1.Add(" WHERE VERSAOTISS In (Select MAX (HANDLE)        ")
  q1.Add("                        FROM TIS_VERSAO          ")
  q1.Add("                       WHERE ATIVODESKTOP = 'S') ")

  q1.Add("   AND CODIGO = :CODIGO        ")
  If Not CurrentQuery.FieldByName("INDICADORDEACIDENTE").IsNull Then
    q1.ParamByName("CODIGO").AsString = CurrentQuery.FieldByName("INDICADORDEACIDENTE").AsString
    q1.Active = True
    If (Not q1.FieldByName("FINALIDADEATENDIMENTO").IsNull) And _
       (FINALIDADEATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger = q1.FieldByName("FINALIDADEATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("LOCALATENDIMENTO").IsNull) And _
       (LOCALATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger = q1.FieldByName("LOCALATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("REGIMEATENDIMENTO").IsNull) And _
       (REGIMEATENDIMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger = q1.FieldByName("REGIMEATENDIMENTO").AsInteger
    End If
    If (Not q1.FieldByName("TIPOTRATAMENTO").IsNull) And _
       (TIPOTRATAMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger = q1.FieldByName("TIPOTRATAMENTO").AsInteger
    End If
    If (Not q1.FieldByName("OBJETIVOTRATAMENTO").IsNull) And _
       (OBJETIVOTRATAMENTO.ReadOnly)                     Then
      CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger = q1.FieldByName("OBJETIVOTRATAMENTO").AsInteger
    End If
  End If
  Set q1 = Nothing
End Sub

Public Sub CaracteristicasREGIMEINTERNACAO
  Dim q1 As Object
  Set q1 = NewQuery

  q1.Add(" SELECT REGIMEATENDIMENTO                         ")
  q1.Add("   FROM TIS_REGIMEINTERNACAO                      ")
  q1.Add("  WHERE CODIGO = :CODIGO                          ")
  q1.Add("    And VERSAOTISS In (Select MAX (HANDLE)        ")
  q1.Add("                         FROM TIS_VERSAO          ")
  q1.Add("                        WHERE ATIVODESKTOP = 'S') ")

  q1.ParamByName("CODIGO").AsString = CurrentQuery.FieldByName("REGIMEINTERNACAO").AsString
  q1.Active = True


  If (Not CurrentQuery.FieldByName("REGIMEINTERNACAO").IsNull) And (Not REGIMEINTERNACAO.ReadOnly) And ( Not q1.FieldByName("REGIMEATENDIMENTO").IsNull ) Then
    CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger = q1.FieldByName("REGIMEATENDIMENTO").AsInteger
  End If

  Set q1 = Nothing
End Sub

Public Sub ListaCamposLeiaute(MODELOGUIAx, Fase As Integer)
  Dim q1 As Object
  Dim vAux As String
  Dim vVisualizar As String

  If Fase=1 Then
    vVisualizar="D"
  Else
    vVisualizar="C"
  End If
  Set q1=NewQuery
  q1.Clear
  q1.Add("SELECT SIS.ZCAMPO FROM SIS_MODELOGUIA_CAMPOS SIS, SAM_TIPOGUIA_MDGUIA_EVENTO E")
  q1.Add("WHERE")
  q1.Add("E.SISCAMPO = SIS.HANDLE")
  q1.Add("AND E.VISUALIZAR IN (:VISUALIZAR, 'A')")
  q1.Add("AND E.MODELOGUIA="+Str(MODELOGUIAx))
  q1.ParamByName("VISUALIZAR").AsString=vVisualizar
  q1.Active=True
  vAux="|"
  While Not q1.EOF
    vAux=vAux+q1.FieldByName("ZCAMPO").AsString+"|"
    q1.Next
  Wend
  q1.Active=False
  UserVar("CAMPOS_LEIAUTE_GUIA_EVENTO") = vAux


  q1.Clear
  q1.Add("SELECT SIS.ZCAMPO FROM SIS_MODELOGUIA_CAMPOS SIS, SAM_TIPOGUIA_MDGUIA_GUIA E")
  q1.Add("WHERE")
  q1.Add("E.SISCAMPO = SIS.HANDLE")
  q1.Add("AND E.VISUALIZAR IN (:VISUALIZAR, 'A')")
  q1.Add("AND E.MODELOGUIA="+Str(MODELOGUIAx))
  q1.ParamByName("VISUALIZAR").AsString=vVisualizar
  q1.Active=True
  vAux="|"
  While Not q1.EOF
    vAux=vAux+q1.FieldByName("ZCAMPO").AsString+"|"
    q1.Next
  Wend
  q1.Active=False
  UserVar("CAMPOS_LEIAUTE_GUIA") = vAux

  Set q1 = Nothing
End Sub

Public Function ProcuraBeneficiarioAtivo(pSoAtivos As Boolean,  pData As Date, TextoBenAtivo As String) As Long
  Dim INTERFACE As Object
  Dim vWhere As String
  Dim vColunas As String
  Dim vDllDigit As Object
  Dim vOrdem As Integer
  Dim qParametros As Object
  Set qParametros=NewQuery
  Set vDllDigit = CreateBennerObject("SAMPEGDIGIT.DIGITACAO")

  If (TextoBenAtivo <> "") Then
  	ProcuraBeneficiarioAtivo = vDllDigit.ValidarBeneficiario(CurrentSystem, TextoBenAtivo)
  End If

  If ProcuraBeneficiarioAtivo <= 0 Then
    qParametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
    qParametros.Active=True
    If qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString="S" Then
       If IsNumeric(TextoBenAtivo) Then
          vOrdem = 0
       Else
          vOrdem = 1
       End If

      'Separação da Interface da regra de negocio para consulta de Beneficiários
      Set INTERFACE =CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")
      ProcuraBeneficiarioAtivo=INTERFACE.Filtro(CurrentSystem,vOrdem, TextoBenAtivo)
      Set INTERFACE=Nothing
    End If
    If qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString="N" Then
      vColunas="SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_CONTRATO.CONTRATANTE|SAM_BENEFICIARIO.CODIGODEAFINIDADE|SAM_BENEFICIARIO.CODIGOANTIGO|SAM_BENEFICIARIO.DATACANCELAMENTO|SAM_CONVENIO.DESCRICAO|SAM_BENEFICIARIO.CODIGODEORIGEM|SAM_BENEFICIARIO.CODIGODEREPASSE"
      vWhere= ""
      If pSoAtivos = True Then
        Dim vData As String
        vData=SQLDate(pData)
         vWhere=vWhere+"(SAM_BENEFICIARIO.DATABLOQUEIO IS NULL) And "
         vWhere=vWhere+" ((SAM_BENEFICIARIO.ATENDIMENTOATE Is NOT NULL AND SAM_BENEFICIARIO.ATENDIMENTOATE >= "+vData+") OR (SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL OR SAM_BENEFICIARIO.DATACANCELAMENTO >= "+vData+"))"
      End If
      If IsNumeric(TextoBenAtivo) Then
         vOrdem = 1
      Else
        vOrdem = 2
      End If
      Set INTERFACE=CreateBennerObject("Procura.Procurar")
      ProcuraBeneficiarioAtivo=INTERFACE.Exec(CurrentSystem,"SAM_BENEFICIARIO|SAM_CONTRATO[SAM_BENEFICIARIO.CONTRATO=SAM_CONTRATO.HANDLE]|SAM_CONVENIO[SAM_BENEFICIARIO.CONVENIO=SAM_CONVENIO.HANDLE]",vColunas,vOrdem,"Matrícula Funcional|Nome|Beneficiario|Contratante|Código Afinidade|Código Antigo|Data Cancelamento|Convenio|Código de origem|Código de repasse",vWhere,"Procura por Beneficiário",False,TextoBenAtivo ,"CA006.ConsultaBeneficiario")
      Set INTERFACE=Nothing
    End If
  End If
End Function

Public Function ProcuraBeneficiarioAtivoReembolso(pSoAtivos As Boolean,  pData As Date, TextoBenAtivo As String, pBeneficiarioTitular As Long, pReembolso As Boolean) As Long
	Dim INTERFACE As Object
	Dim vWhere As String
	Dim vColunas As String
	Dim vDllDigit As Object
	Dim qParametros As Object
	Set qParametros=NewQuery

	Set vDllDigit = CreateBennerObject("SAMPEGDIGIT.DIGITACAO")

	If (TextoBenAtivo <> "") Then
		ProcuraBeneficiarioAtivoReembolso = vDllDigit.ValidarBeneficiario(CurrentSystem, TextoBenAtivo)
	End If

	If ProcuraBeneficiarioAtivoReembolso <= 0 Then
		qParametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
		qParametros.Active=True
		If qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString="S" Then
			'Separação da Interface da regra de negocio para consulta de Beneficiários
			Set INTERFACE =CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")
			ProcuraBeneficiarioAtivoReembolso=INTERFACE.FiltroTitular(CurrentSystem,1,TextoBenAtivo, pBeneficiarioTitular)
			Set INTERFACE=Nothing
		End If

		If qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString="N" Then
			vColunas="SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_CONTRATO.CONTRATANTE|SAM_BENEFICIARIO.CODIGODEAFINIDADE|SAM_BENEFICIARIO.CODIGOANTIGO|SAM_BENEFICIARIO.DATACANCELAMENTO|SAM_CONVENIO.DESCRICAO|SAM_BENEFICIARIO.CODIGODEORIGEM|SAM_BENEFICIARIO.CODIGODEREPASSE"

			vWhere= ""

			If pSoAtivos = True Then
				Dim vData As String

				vData=SQLDate(pData)

				vWhere = vWhere + "(SAM_BENEFICIARIO.DATABLOQUEIO IS NULL) And "
				vWhere = vWhere + " ((SAM_BENEFICIARIO.ATENDIMENTOATE Is NOT NULL AND SAM_BENEFICIARIO.ATENDIMENTOATE >= "+vData+") OR (SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL OR SAM_BENEFICIARIO.DATACANCELAMENTO >= "+vData+"))"
			End If

			If pBeneficiarioTitular > 0 Then
				Dim qAux As Object
				Set qAux = NewQuery
				qAux.Clear
				qAux.Add("SELECT FAMILIA, MATRICULA FROM SAM_BENEFICIARIO WHERE HANDLE=:BN")
				qAux.ParamByName("BN").AsInteger = pBeneficiarioTitular
				qAux.Active = True

				vWhere = vWhere + "SAM_BENEFICIARIO.HANDLE IN ("
				vWhere = vWhere + "SELECT HANDLE FROM SAM_BENEFICIARIO WHERE FAMILIA=" + qAux.FieldByName("FAMILIA").AsString
				vWhere = vWhere + " UNION "
				vWhere = vWhere + "SELECT HANDLE FROM SAM_BENEFICIARIO WHERE MATRICULAINDICADORA = " + qAux.FieldByName("MATRICULA").AsString + ")  AND SAM_BENEFICIARIO.CONVENIO=SAM_CONVENIO.HANDLE " + Chr(13) + " "

				Set qAux = Nothing
			End If

			If pReembolso Then
				If Len(vWhere) > 0 Then
					vWhere = vWhere + " AND "
				End If

				vWhere = vWhere + "SAM_BENEFICIARIO.PERMITEREEMBOLSO='S'"
			End If

			Set INTERFACE=CreateBennerObject("Procura.Procurar")

			ProcuraBeneficiarioAtivoReembolso=INTERFACE.Exec(CurrentSystem,"SAM_BENEFICIARIO|SAM_CONTRATO[SAM_BENEFICIARIO.CONTRATO=SAM_CONTRATO.HANDLE]|SAM_CONVENIO[SAM_BENEFICIARIO.CONVENIO=SAM_CONVENIO.HANDLE]",vColunas,2,"Matrícula Funcional|Nome|Beneficiario|Contratante|Código Afinidade|Código Antigo|Data Cancelamento|Convenio|Código de origem|Código de repasse",vWhere,"Procura por Beneficiário",False,TextoBenAtivo ,"CA006.ConsultaBeneficiario")

			Set INTERFACE=Nothing
		End If
	End If

End Function

Public Function ProcuraPrestador(CPF_Nome As String, Sol_Exe_Rec_Todos As String, TextoPrestador As String ) As Long

	Dim INTERFACE As Object
	Dim vCriterio As String
	Dim qParametros As Object
	Set qParametros=NewQuery
	Dim vCampos As String
	Dim vColunas As String

	qParametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
	qParametros.Active=True

	If qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString="S" Then
		'Verifica se há apenas um prestador com aquele nome ou código de prestador;
		'se isto acontecer, retorna este prestador e não abre interface de busca
		Dim qPreConsulta As Object
		Dim vContinua As Boolean

		Set qPreConsulta = NewQuery
		vContinua = True
		ProcuraPrestador = -1
		qPreConsulta.Add("SELECT HANDLE FROM SAM_PRESTADOR")

		If CPF_Nome = "C" Then
			qPreConsulta.Add("WHERE PRESTADOR LIKE :TextoPrestador")
		Else
			qPreConsulta.Add("WHERE UPPER(NOME) LIKE UPPER(:TextoPrestador)")
		End If

		qPreConsulta.ParamByName("TextoPrestador").AsString = TextoPrestador & "%"
		qPreConsulta.Active = True

		If Not qPreConsulta.EOF Then
			ProcuraPrestador = qPreConsulta.FieldByName("HANDLE").AsInteger
			qPreConsulta.Next
			vContinua = Not qPreConsulta.EOF
		Else
			qPreConsulta.Clear
			qPreConsulta.Add("SELECT HANDLE FROM SAM_PRESTADOR")
			qPreConsulta.Add("WHERE CPFCNPJ LIKE :TextoPrestador")
			qPreConsulta.ParamByName("TextoPrestador").AsString = TextoPrestador & "%"
			qPreConsulta.Active = True

			If Not qPreConsulta.EOF Then
				ProcuraPrestador = qPreConsulta.FieldByName("HANDLE").AsInteger
				qPreConsulta.Next
				vContinua = Not qPreConsulta.EOF
			End If
		End If

		Set qPreConsulta = Nothing

		If Not vContinua Then
			Set qParametros = Nothing
			Exit Function
		End If

		Set INTERFACE = CreateBennerObject("BSINTERFACE0001.BuscaPrestador")

		If CPF_Nome = "C" Then

			Dim vAux As String
			Dim vPosicao As Integer
			Dim VIRETORNO As Integer

			vAux = Left(TextoPrestador,1)

			If Len(vAux) > 0 Then
				If vAux = "0" Or _
				   vAux = "1" Or _
				   vAux = "2" Or _
				   vAux = "3" Or _
				   vAux = "4" Or _
				   vAux = "5" Or _
				   vAux = "6" Or _
				   vAux = "7" Or _
				   vAux = "8" Or _
				   vAux = "9" Then
					vPosicao = 0
				Else
					vPosicao = 1
				End If
			End If

			If (Sol_Exe_Rec_Todos = "E") Then
				VIRETORNO= INTERFACE.Abrir(CurrentSystem,vSMensagem,vPosicao,TextoPrestador ,"E",ProcuraPrestador, 0, vFormacaoPrestador)

			ElseIf (Sol_Exe_Rec_Todos = "T") Then
				VIRETORNO= INTERFACE.Abrir(CurrentSystem,vSMensagem,vPosicao,TextoPrestador ,"T",ProcuraPrestador)
			End If
		Else
			If (Sol_Exe_Rec_Todos = "E") Then
				VIRETORNO= INTERFACE.SelecionaPrestador(CurrentSystem,vSMensagem,1,TextoPrestador ,"E",ProcuraPrestador, 0, vFormacaoPrestador)
			ElseIf (Sol_Exe_Rec_Todos = "T") Then
				VIRETORNO= INTERFACE.SelecionaPrestador(CurrentSystem,vSMensagem,1,TextoPrestador ,"T",ProcuraPrestador)
			End If
		End If

		Set INTERFACE = Nothing
	End If

	If qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString="N" Then

		Set INTERFACE = CreateBennerObject("Procura.Procurar")

		vColunas = "SAM_PRESTADOR.PRESTADOR"
		vColunas =vColunas + "|SAM_PRESTADOR.Z_NOME"
		vColunas = vColunas + "|SAM_PRESTADOR.CPFCNPJ"
		vColunas = vColunas + "|SAM_PRESTADOR.INSCRICAOCR"
		vColunas = vColunas + "|SAM_PRESTADOR.DATACREDENCIAMENTO"
		vColunas = vColunas + "|SAM_PRESTADOR.DATADESCREDENCIAMENTO"
		vColunas = vColunas + "|SAM_PRESTADOR.SOLICITANTE"
		vColunas = vColunas + "|SAM_PRESTADOR.EXECUTOR"
		vColunas = vColunas + "|SAM_PRESTADOR.RECEBEDOR"
		vColunas = vColunas + "|ESTADOS.NOME NOMEESTADO"
		vColunas = vColunas + "|MUNICIPIOS.NOME NOMEMUNICIPIO"

		vCriterio = "(1=1)"

		If (Sol_Exe_Rec_Todos = "L") Then
			vCriterio = vCriterio + " AND SAM_PRESTADOR.LOCALEXECUCAO = 'S' "
		Else
			If (Sol_Exe_Rec_Todos = "S") Then
				vCriterio = vCriterio + " AND SAM_PRESTADOR.SOLICITANTE = 'S' "
			Else
				If (Sol_Exe_Rec_Todos = "E") Then
					vCriterio = vCriterio +  " AND SAM_PRESTADOR.EXECUTOR = 'S' "
				Else
					If (Sol_Exe_Rec_Todos = "R") Then
						vCriterio = vCriterio +  " AND SAM_PRESTADOR.RECEBEDOR = 'S'  "
					End If
				End If
			End If
		End If

		vCampos = "Prestador|Nome do Prestador|CPF/CNPJ|Nr.Conselho|Credenciam.|Descredenc|Sol|Exe|Rec|Estado|Município"

		If CPF_Nome = "C" Then
			ProcuraPrestador = INTERFACE.Exec(CurrentSystem,"SAM_PRESTADOR|*SAM_CONSELHO[SAM_CONSELHO.HANDLE =SAM_PRESTADOR.CONSELHOREGIONAL]|*ESTADOS[ESTADOS.HANDLE =SAM_PRESTADOR.ESTADOPAGAMENTO]|*MUNICIPIOS[MUNICIPIOS.HANDLE =SAM_PRESTADOR.MUNICIPIOPAGAMENTO]",vColunas,1,vCampos,vCriterio,"Prestadores", False,TextoPrestador, "CA005.ConsultaPrestador")
		Else
			ProcuraPrestador = INTERFACE.Exec(CurrentSystem,"SAM_PRESTADOR|*SAM_CONSELHO[SAM_CONSELHO.HANDLE =SAM_PRESTADOR.CONSELHOREGIONAL]|*ESTADOS[ESTADOS.HANDLE =SAM_PRESTADOR.ESTADOPAGAMENTO]|*MUNICIPIOS[MUNICIPIOS.HANDLE =SAM_PRESTADOR.MUNICIPIOPAGAMENTO]",vColunas,2,vCampos,vCriterio,"Prestadores",  False,TextoPrestador,"CA005.ConsultaPrestador")
		End If

		Set INTERFACE = Nothing
	End If

End Function

Public Function IsInt(pValor As String) As Boolean
  Dim vAux As Long

  On Error GoTo Erro
  vAux = CLng(pValor)
  IsInt = True
  Exit Function
  Erro:
  IsInt = False
End Function

Public Function ProcuraEvento (pUltimoNivel As Boolean, TextoEvento As String) As Long

  Dim sql As Object
  Set sql=NewQuery
  On Error GoTo Pula1
  sql.Clear
  sql.Add("SELECT HANDLE FROM SAM_TGE WHERE ESTRUTURANUMERICA=:P1")

  If pUltimoNivel Then
    sql.Add(" AND ULTIMONIVEL = 'S' ")
  End If

  sql.ParamByName("P1").AsInteger= CLng(TextoEvento)
  sql.Active=True
  If sql.FieldByName("HANDLE").AsInteger>0 Then
    ProcuraEvento=sql.FieldByName("HANDLE").AsInteger
    Set sql=Nothing
    Exit Function
  End If
  Pula1:
  On Error GoTo Pula2
  sql.Clear
  sql.Add("SELECT HANDLE FROM SAM_TGE WHERE ESTRUTURA=:P1")

  If pUltimoNivel Then
    sql.Add(" AND ULTIMONIVEL = 'S' ")
  End If

  sql.ParamByName("P1").AsString=TextoEvento
  sql.Active=True
  If sql.FieldByName("HANDLE").AsInteger>0 Then
    ProcuraEvento=sql.FieldByName("HANDLE").AsInteger
    Set sql=Nothing
    Exit Function
  End If
  Pula2:
  Set sql=Nothing
  Dim INTERFACE As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vOrdem As Integer

  Set INTERFACE = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.Z_DESCRICAO|SAM_CBHPM.DESCRICAO|SAM_TGE.DESCRICAOABREVIADA|SAM_TGE.NIVELAUTORIZACAO"

  If pUltimoNivel Then
    vCriterio = "SAM_TGE.ULTIMONIVEL = 'S' "
  Else
    vCriterio = "SAM_TGE.HANDLE > 0"
  End If

  If IsInt(TiraAcento(TextoEvento,True)) Then
    vOrdem = 2
  Else
    vOrdem = 3
  End If


  vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM|Descrição abreviada TGE|Nível"

  ProcuraEvento = INTERFACE.Exec(CurrentSystem, "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]",vColunas,vOrdem ,vCampos,vCriterio, _
  "Tabela Geral de Eventos",False,TextoEvento,"CA011.ConsultaTge")

  Set INTERFACE = Nothing

End Function

Function CalculaIdadeBeneficiario(ByVal pBeneficiario As Long, ByVal pDataAtendimento As Date)As Integer
  Dim vDias           As Integer
  Dim vMeses          As Integer
  Dim vAnos           As Integer
  Dim VDataNascimento As Date
  Dim query As Object


  Set query = NewQuery
  query.Clear
  query.Add("SELECT M.DATANASCIMENTO ")
  query.Add("  FROM SAM_MATRICULA M ")
  query.Add("  JOIN SAM_BENEFICIARIO B ON B.MATRICULA = M.HANDLE ")
  query.Add(" WHERE B.HANDLE  = :HANDLE ")
  query.ParamByName("HANDLE").AsInteger = pBeneficiario
  query.Active = True

  VDataNascimento = query.FieldByName("DATANASCIMENTO").AsDateTime
  If (VDataNascimento > ServerDate) Then
    CalculaIdadeBeneficiario = 0
  Else
    DiferencaData2 pDataAtendimento, VDataNascimento, vDias, vMeses, vAnos
    CalculaIdadeBeneficiario = vAnos
  End If
End Function

Public Sub DiferencaData2(ByVal Data1, Data2 As Date, Dias, Meses, Anos As Integer)
  Dim DtSwap As Date
  Dim Day1, Day2, Month1, Month2, Year1, Year2 As Integer

  If Data1 >Data2 Then
    DtSwap = Data1
    Data1 = Data2
    Data2 = DtSwap
  End If

  Year1 = Val(Format(Data1, "yyyy"))
  Month1 = Val(Format(Data1, "mm"))
  Day1 = Val(Format(Data1, "dd"))

  Year2 = Val(Format(Data2, "yyyy"))
  Month2 = Val(Format(Data2, "mm"))
  Day2 = Val(Format(Data2, "dd"))

  Anos = Year2 - Year1
  Meses = 0
  Dias = 0
  If Month2 <Month1 Then
    Meses = Meses + 12
    Anos = Anos -1
  End If
  Meses = Meses + (Month2 - Month1)
  If Day2 <Day1 Then
    Dias = Dias + DiasPorMes(Year1, Val(Month1))
    If Meses = 0 Then
      Anos = Anos -1
      Meses = 11
    Else
      Meses = Meses -1
    End If
  End If
  Dias = Dias + (Day2 - Day1)
End Sub

Function DiasPorMes(ByVal Ano, Mes As Integer)As Integer
  Dim Meses31 As String
  Dim Meses30 As String

  Meses31 = "'1','3','5','7','8','10','12'"
  Meses30 = "'4','6','9','11'"

  If InStr(Meses31, "'" + Str(Mes) + "'")>0 Then
    DiasPorMes = 31
  ElseIf InStr(Meses30, "'" + Str(Mes) + "'")>0 Then
    DiasPorMes = 30
  Else
    If Ano Mod 4 = 0 Then
      DiasPorMes = 29
    Else
      DiasPorMes = 28
    End If
  End If

End Function

Public Sub Log (Texto As String )
	Dim query As Object
	Set query = NewQuery

	query.Add("INSERT INTO ABREV (HANDLE, TEXTO) VALUES (:HANDLE, :TEXTO)")
	query.ParamByName("HANDLE").AsInteger = NewHandle("ABREV")
	query.ParamByName("TEXTO").AsString = Texto
	query.ExecSQL

	Set query = Nothing

End Sub

Public Sub CHECATETOREEMBOLSO()
    Dim qSituacaoPeg As Object
    Set qSituacaoPeg = NewQuery

    qSituacaoPeg.Clear
    qSituacaoPeg.Add("SELECT P.SITUACAO, P.TABREGIMEPGTO, P.HANDLE ")
    qSituacaoPeg.Add("  FROM SAM_PEG    P ")
    qSituacaoPeg.Add(" WHERE P.HANDLE = :HANDLE")
    qSituacaoPeg.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
    qSituacaoPeg.Active = True

    If (qSituacaoPeg.FieldByName("SITUACAO").AsString = "3") And (qSituacaoPeg.FieldByName("TABREGIMEPGTO").AsInteger = 2) Then ' pronto deve checar o teto de reembolso pois se foi feito a revisão ele retira e não considera o teto
      Dim spCalcTetoFinanciamento As BStoredProc
      Set spCalcTetoFinanciamento = NewStoredProc

      spCalcTetoFinanciamento.Name = "BSPROPEG_CALCTETOFINANCIAMENTO"
      spCalcTetoFinanciamento.AddParam("P_PEG",ptInput,ftInteger)
      spCalcTetoFinanciamento.AddParam("P_CHAVE",ptInput,ftInteger)
      spCalcTetoFinanciamento.AddParam("P_USUARIO",ptInput,ftInteger)
      spCalcTetoFinanciamento.ParamByName("P_PEG").AsInteger = qSituacaoPeg.FieldByName("HANDLE").AsInteger
      spCalcTetoFinanciamento.ParamByName("P_CHAVE").AsInteger = qSituacaoPeg.FieldByName("HANDLE").AsInteger
      spCalcTetoFinanciamento.ParamByName("P_USUARIO").AsInteger = CurrentUser
      spCalcTetoFinanciamento.ExecProc

      Set spCalcTetoFinanciamento = Nothing
    End If

    Set qSituacaoPeg = Nothing

End Sub

Public Sub SelecionaCnesEndPrest(psPrestador As String)

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.SelecionaCnesEndPrest(CurrentSystem, CurrentQuery.TQuery, psPrestador)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub ContaQtdeEndPrest (prestador As String)
  Dim vPrestador As Long
  Dim qQtdEndPrestador As Object
  Set qQtdEndPrestador = NewQuery

  If prestador = "Executor" Then
    vPrestador = CurrentQuery.FieldByName("EXECUTOR").AsInteger
  ElseIf prestador = "LocalExec" Then
    vPrestador = CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger
  End If

  qQtdEndPrestador.Active = False
  qQtdEndPrestador.Clear
  qQtdEndPrestador.Add("SELECT COUNT(HANDLE) QTDENDERECOS                            ")
  qQtdEndPrestador.Add("  FROM SAM_PRESTADOR_ENDERECO                                ")
  qQtdEndPrestador.Add(" WHERE PRESTADOR = :PRESTADOR                                ")
  qQtdEndPrestador.Add("   AND DATACANCELAMENTO IS NULL                              ")
  qQtdEndPrestador.Add("   AND ATENDIMENTO = 'S'                                     ")

  qQtdEndPrestador.ParamByName("PRESTADOR").AsInteger = vPrestador
  qQtdEndPrestador.Active = True

  If qQtdEndPrestador.FieldByName("QTDENDERECOS").AsInteger = 1 Then
    Dim qEndPrestador As Object
    Set qEndPrestador = NewQuery

    qEndPrestador.Active = False
    qEndPrestador.Clear
    qEndPrestador.Add("SELECT HANDLE FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = :PRESTADOR")
    qEndPrestador.ParamByName("PRESTADOR").AsInteger = vPrestador
    qEndPrestador.Active = True

	If prestador = "Executor" Then
      CurrentQuery.FieldByName("ENDERECOEXECUTOR").AsInteger = qEndPrestador.FieldByName("HANDLE").AsString
	  ENDERECOEXECUTOR_OnChange
	ElseIf prestador = "LocalExec" Then
      CurrentQuery.FieldByName("ENDERECOLOCALEXEC").AsInteger = qEndPrestador.FieldByName("HANDLE").AsString
      ENDERECOLOCALEXEC_OnChange
    End If

    Set qEndPrestador = Nothing
  End If

  If prestador = "Executor" Then
    If CurrentQuery.FieldByName("ENDERECOEXECUTOR").IsNull Then
      CurrentQuery.FieldByName("EXECUTORCNES").Clear
    End If
  ElseIf prestador = "LocalExec" Then
    If CurrentQuery.FieldByName("ENDERECOLOCALEXEC").IsNull Then
      CurrentQuery.FieldByName("LOCALEXECUCAOCNES").Clear
    End If
  End If

  Set qQtdEndPrestador = Nothing
End Sub

Public Sub IncluiSessionVarMonitoramento()

  SessionVar("MONITORAMENTOPEG") = "0"
  SessionVar("MONITORAMENTOGUIA") = CurrentQuery.FieldByName("HANDLE").AsString

End Sub

Public Function VersaoTissDoPEG As String

	Dim sqlVersaoTISS As Object
	Set sqlVersaoTISS = NewQuery

	Dim vsVersaoTiss As String

	sqlVersaoTISS.Active = False
	sqlVersaoTISS.Clear
	sqlVersaoTISS.Add(" SELECT P.VERSAOTISS    ")
	sqlVersaoTISS.Add("   FROM SAM_PEG P       ")
	sqlVersaoTISS.Add("  WHERE P.HANDLE = :PEG ")
	sqlVersaoTISS.ParamByName("PEG").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
	sqlVersaoTISS.Active = True

	If sqlVersaoTISS.FieldByName("VERSAOTISS").IsNull Then
		VersaoTissDoPEG = "VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S') "
	Else
		VersaoTissDoPEG = "VERSAOTISS = " + sqlVersaoTISS.FieldByName("VERSAOTISS").AsString
	End If

	Set sqlVersaoTISS = Nothing

End Function

Public Sub HerdarAutorizacao

	If CurrentQuery.State <>1 Then
		If(OldAutorizacao <>CurrentQuery.FieldByName("AUTORIZACAO").AsInteger)Then
			If VisibleMode Then
				If Not CurrentQuery.FieldByName("AUTORIZACAO").IsNull Then
					Dim INTERFACE As Object
					Set INTERFACE = CreateBennerObject("SAMPEGDIGIT.Digitacao")
					cancontinue = INTERFACE.HerdardaAutorizacao(CurrentSystem)
					Set INTERFACE = Nothing
				End If
			End If
			OldAutorizacao = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
		End If
	End If

End Sub

Public Function VerificarCC As String

	If (CurrentQuery.State = 1) Then
		Exit Function
	End If

	OLDBENEFICIARIO = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

	vDllBSPro006.VerificarCC(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Function

Public Sub SugerirIdadeBeneficiario

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.Geral")

	vDllBSPro006.SugerirIdadeBeneficiario(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TratarCaracteristicasAtendimento()

  If Not TIPOINTERNACAO.ReadOnly Then
    CaracteristicasTIPOINTERNACAO
  End If

  If Not TIPOATENDIMENTO.ReadOnly Then
    CaracteristicasTIPOATENDIMENTO
  End If

  If Not INDICADORDEACIDENTE.ReadOnly Then
    CaracteristicasINDICADORDEACIDENTE
  End If

  If (Not REGIMEINTERNACAO.ReadOnly) Then
    CaracteristicasREGIMEINTERNACAO
  End If

End Sub

Public Sub VERIFICAJAPAGO(cancontinue As Boolean)

	Dim SITUACAO As String
	Dim q1 As Object

	Set q1 = NewQuery

	q1.Clear
	q1.Add("SELECT G.SITUACAO FROM SAM_GUIA G WHERE G.HANDLE=:GUIA")
	q1.ParamByName("GUIA").Value = RecordHandleOfTable("SAM_GUIA")
	q1.Active = True
	SITUACAO = q1.FieldByName("SITUACAO").AsString
	q1.Active = False

	Set q1 = Nothing

	If SITUACAO = "4" Then
		cancontinue = False
	Else
		cancontinue = True
	End If

End Sub
