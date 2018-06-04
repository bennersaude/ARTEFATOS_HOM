'HASH: BC4AC10183CABD7929C7C327AF82BCBB
'Macro: SAM_PRESTADOR_PROC_CREDEN

'#Uses "*bsShowMessage"

'Mauricio Ibelli - 04/01/2002 - sms3165 - Se filial padrao do prestador for nulo não checar responsavel

Dim Mensagem As String

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Dim S As Object
  Set S = NewQuery
  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True

  SQL.Add("SELECT SAM_PRESTADOR_PROC.DATAFINAL,SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.FILIALPADRAO FROM SAM_PRESTADOR_PROC, SAM_PRESTADOR WHERE SAM_PRESTADOR_PROC.HANDLE = :HANDLE And  SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PROC.PRESTADOR")
  If VisibleMode Then
  	SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  Else
  	SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger
  End If

  SQL.Active = True
  Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And _
       ( (SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser) Or (SQL.FieldByName("FILIALPADRAO").IsNull) ), _
       True, False)

  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo finalizado! Operação não permitida." + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

Public Sub BOTAOGERAROFICIODOCUMENTOS_OnClick()

	On Error GoTo Exception

	Dim componente As CSBusinessComponent
	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcCredenBLL, Benner.Saude.Prestadores.Business")
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.Execute("GerarOficioDocumentos")
	Set componente = Nothing
	RefreshNodesWithTable("")
	Exit Sub

	Exception:
    	Set componente = Nothing
    	bsShowMessage(Err.Description, "I")
    	Exit Sub

End Sub

Public Sub BOTAOSOLICITARDOCUMENTOS_OnClick()

	On Error GoTo Exception

	Dim componente As CSBusinessComponent
	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcCredenBLL, Benner.Saude.Prestadores.Business")
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.Execute("SolicitarDocumentosEnvioEmail")
	Set componente = Nothing
	RefreshNodesWithTable("")
	Exit Sub

	Exception:
    	Set componente = Nothing
    	bsShowMessage(Err.Description, "I")
    	Exit Sub

End Sub

Public Sub TABLE_AfterScroll()

	Dim qBuscaEmail As BPesquisa
	Set qBuscaEmail = NewQuery

	qBuscaEmail.Active = False
	qBuscaEmail.Clear

	qBuscaEmail.Add("SELECT EMAIL FROM Z_GRUPOUSUARIOS WHERE HANDLE = :PHANDLE")
	qBuscaEmail.ParamByName("PHANDLE").AsInteger = CurrentUser

	qBuscaEmail.Active = True

	SessionVar("TIPONOTIFICACAO") = "4"
	SessionVar("REMETENTE") = qBuscaEmail.FieldByName("EMAIL").AsString
	SessionVar("HANDLEPROCESSO") = CurrentQuery.FieldByName("HANDLE").AsString

	Set qBuscaEmail = Nothing

	If WebMode Then
		If CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger > 0 Then
			Dim SQL As BPesquisa
			Set SQL = NewQuery

			SQL.Clear
			SQL.Add("SELECT ATUALIZADATACREDENCIAMENTO FROM SAM_TIPOPROCESSOCREDENCTO WHERE HANDLE = :HANDLE")
			SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger
			SQL.Active = True

			If SQL.FieldByName("ATUALIZADATACREDENCIAMENTO").AsString = "S" Then
				DATACREDENCIAMENTO.ReadOnly = False
			Else
				DATACREDENCIAMENTO.ReadOnly = True
			End If
		End If
	End If


	If(Not WebMode) Then
		Dim vExibeBotoes As Boolean

		Dim componente As CSBusinessComponent
		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcCredenBLL, Benner.Saude.Prestadores.Business")
		componente.AddParameter(pdtInteger, RecordHandleOfTable("SAM_PRESTADOR_PROC_CREDEN"))
		vExibeBotoes = componente.Execute("VerificarControleDocumentacaoProcesso")

		BOTAOGERAROFICIODOCUMENTOS.Visible = vExibeBotoes
		BOTAOSOLICITARDOCUMENTOS.Visible = vExibeBotoes

		Set componente = Nothing
		RefreshNodesWithTable("")
		Exit Sub

	End If


End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT CATEGORIA, ISS FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  AdicionarParametroPrestadorNaQuery(SQL)
  SQL.Active = True

  If WebMode Then
    TIPOCREDENCIAMENTO.WebLocalWhere = "A.HANDLE " + _
                                       "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_CATEGORIA WHERE TIPOPROCESSO = A.HANDLE AND CATEGORIA = " + SQL.FieldByName("CATEGORIA").AsInteger + ") AND " + _
                                       "A.HANDLE " + _
                                       "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_ISS WHERE TIPOPROCESSO = A.HANDLE AND ISS = " + SQL.FieldByName("ISS").AsInteger + ")"
  Else
	TIPOCREDENCIAMENTO.LocalWhere = "SAM_TIPOPROCESSOCREDENCTO.HANDLE " + _
                                    "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_CATEGORIA WHERE TIPOPROCESSO = SAM_TIPOPROCESSOCREDENCTO.HANDLE AND CATEGORIA = " + SQL.FieldByName("CATEGORIA").AsInteger + ") AND " + _
                                    "SAM_TIPOPROCESSOCREDENCTO.HANDLE " + _
                                    "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_ISS WHERE TIPOPROCESSO = SAM_TIPOPROCESSOCREDENCTO.HANDLE AND ISS = " + SQL.FieldByName("ISS").AsInteger + ")"
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT CATEGORIA, ISS FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  AdicionarParametroPrestadorNaQuery(SQL)
  SQL.Active = True

  If WebMode Then
    TIPOCREDENCIAMENTO.WebLocalWhere = "A.HANDLE " + _
                                       "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_CATEGORIA WHERE TIPOPROCESSO = A.HANDLE AND CATEGORIA = " + SQL.FieldByName("CATEGORIA").AsInteger + ") AND " + _
                                       "A.HANDLE " + _
                                       "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_ISS WHERE TIPOPROCESSO = A.HANDLE AND ISS = " + SQL.FieldByName("ISS").AsInteger + ")"
  Else
	TIPOCREDENCIAMENTO.LocalWhere = "SAM_TIPOPROCESSOCREDENCTO.HANDLE " + _
                                    "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_CATEGORIA WHERE TIPOPROCESSO = SAM_TIPOPROCESSOCREDENCTO.HANDLE AND CATEGORIA = " + SQL.FieldByName("CATEGORIA").AsInteger + ") AND " + _
                                    "SAM_TIPOPROCESSOCREDENCTO.HANDLE " + _
                                    "IN (SELECT TIPOPROCESSO FROM SAM_TIPOPROCESSO_ISS WHERE TIPOPROCESSO = SAM_TIPOPROCESSOCREDENCTO.HANDLE AND ISS = " + SQL.FieldByName("ISS").AsInteger + ")"
  End If

  Set SQL = Nothing
End Sub



Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL, PRESTADOR As Object
  Set SQL = NewQuery
  Set PRESTADOR = NewQuery

  SQL.Clear
  SQL.Add("SELECT ATUALIZAFILIAL, ATUALIZADATACREDENCIAMENTO FROM SAM_TIPOPROCESSOCREDENCTO WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger
  SQL.Active = True

  If (SQL.FieldByName("ATUALIZAFILIAL").Value = "N") And _
     Not(CurrentQuery.FieldByName ("NOVAFILIAL").IsNull) Then
    CurrentQuery.FieldByName("NOVAFILIAL").Clear
    bsShowMessage("O Tipo do Processo está configurado para não atualizar a filial! O conteúdo do campo ""Nova filial"" foi apagado.", "I")
  End If

  If (SQL.FieldByName("ATUALIZADATACREDENCIAMENTO").Value = "N") And _
     Not (CurrentQuery.FieldByName ("DATACREDENCIAMENTO").IsNull) Then
    CurrentQuery.FieldByName("DATACREDENCIAMENTO").Clear
    bsShowMessage("O Tipo do Processo está configurado para não atualizar a data de credenciamento! O conteúdo do campo ""Data credenciamento"" foi apagado.", "I")
  End If

  Set PRESTADOR = NewQuery
  PRESTADOR.Add("SELECT CATEGORIA, ISS, DATADESCREDENCIAMENTO, MOTIVODESCREDENCIAMENTO FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  AdicionarParametroPrestadorNaQuery(PRESTADOR)
  PRESTADOR.Active = True

  SQL.Clear
  SQL.Add("SELECT COUNT(*) TOT FROM SAM_TIPOPROCESSO_CATEGORIA WHERE  TIPOPROCESSO = :TIPOPROCESSO AND CATEGORIA = :CATEGORIA")
  SQL.ParamByName("TIPOPROCESSO").Value = CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger
  SQL.ParamByName("CATEGORIA").Value = PRESTADOR.FieldByName("CATEGORIA").AsInteger
  SQL.Active = True


  If SQL.FieldByName("TOT").AsInteger = 0 Then
    CanContinue = Ok
    bsShowMessage("Tipo de processo inválido para a categoria do prestador", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(*) TOT FROM SAM_TIPOPROCESSO_ISS WHERE  TIPOPROCESSO = :TIPOPROCESSO AND ISS = :ISS")
  SQL.ParamByName("TIPOPROCESSO").Value = CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger
  SQL.ParamByName("ISS").Value = PRESTADOR.FieldByName("ISS").AsInteger
  SQL.Active = True
  If SQL.FieldByName("TOT").AsInteger = 0 Then
    CanContinue = Ok
    bsShowMessage("Tipo de processo inválido para o tipo do prestador (ISS) do prestador", "E")
    CanContinue = False
    Set SQL = Nothing
  End If

  If Not PRESTADOR.FieldByName("MOTIVODESCREDENCIAMENTO").IsNull Then

    Dim S As Object
    Set S = NewQuery

    S.Add("SELECT PERMITERECREDENCIAMENTO FROM SAM_MOTIVODESCREDENCIAMENTO WHERE HANDLE = :HANDLE")
    S.ParamByName("HANDLE").Value = PRESTADOR.FieldByName("MOTIVODESCREDENCIAMENTO").AsInteger
    S.Active = True

    If S.FieldByName("PERMITERECREDENCIAMENTO").Value = "N" Then
      CanContinue = False
      bsShowMessage("Motivo do descredenciamento não permite re-credenciamento", "E")
      CanContinue = False
      Exit Sub
    End If

    Set S = Nothing
  End If

  'Claudemir - 06.08.2003 - SMS 18210 - inicio -----
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_CREDEN WHERE PRESTADORPROCESSO = :PRESTADORPROCESSO AND HANDLE <> :HANDLE")
  SQL.ParamByName("PRESTADORPROCESSO").Value = CurrentQuery.FieldByName("PRESTADORPROCESSO").Value
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.Active = True
  If Not SQL.EOF Then
    bsShowMessage("Já existe outro tipo de processo cadastrado!", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If
  'Claudemir - 06.08.2003 - SMS 18210 - fim   -----

  Set SQL = Nothing


  If CurrentQuery.FieldByName ("DATACREDENCIAMENTO").IsNull Then
    Dim qTipoProcesso As Object
    Set qTipoProcesso = NewQuery

    qTipoProcesso.Add("SELECT HANDLE ")
    qTipoProcesso.Add("  FROM SAM_TIPOPROCESSOCREDENCTO  ")
    qTipoProcesso.Add(" WHERE HANDLE =:PTIPOPROCESSO     ")
    qTipoProcesso.Add("   AND ATUALIZADATACREDENCIAMENTO = 'S' ")
    qTipoProcesso.ParamByName("PTIPOPROCESSO").AsInteger = CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger
    qTipoProcesso.Active = True

    If qTipoProcesso.FieldByName("HANDLE").AsInteger > 0 Then
      Set qTipoProcesso = Nothing
      bsShowMessage("Tipo de processo exige a data de credenciamento!", "E")
      DATACREDENCIAMENTO.SetFocus
      CanContinue = False
    End If

    Set qTipoProcesso = Nothing
  End If


End Sub

Public Sub TABLE_NewRecord()
  Dim SQL As Object
  Set SQL = NewQuery
  Dim FILIAL As Long
  Dim FILIALPROC As Long
  Dim Msg As String

  SQL.Add("SELECT DATADESCREDENCIAMENTO FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  AdicionarParametroPrestadorNaQuery(SQL)
  SQL.Active = True

  If Not SQL.FieldByName("DATADESCREDENCIAMENTO").IsNull Then
    BuscarFiliais(CurrentSystem, FILIAL, FILIALPROC, Msg)
    CurrentQuery.FieldByName ("NOVAFILIAL").Value = FILIAL
    NOVAFILIAL.ReadOnly = True
  End If

End Sub

Public Sub TIPOCREDENCIAMENTO_OnChange()
  Dim SQL As Object
  Dim vCondicao As String

  Set SQL = NewQuery

  SQL.Add("SELECT ATUALIZAFILIAL, ATUALIZADATACREDENCIAMENTO FROM SAM_TIPOPROCESSOCREDENCTO WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("ATUALIZAFILIAL").Value = "N" Then
    CurrentQuery.FieldByName ("NOVAFILIAL").Clear
    NOVAFILIAL.ReadOnly = True
  Else
    CurrentQuery.FieldByName ("NOVAFILIAL").Clear
    NOVAFILIAL.ReadOnly = False
  End If

  If SQL.FieldByName("ATUALIZADATACREDENCIAMENTO").Value = "S" Then
    CurrentQuery.FieldByName ("DATACREDENCIAMENTO").Clear
    DATACREDENCIAMENTO.ReadOnly = False
  Else
    CurrentQuery.FieldByName ("DATACREDENCIAMENTO").Clear
    DATACREDENCIAMENTO.ReadOnly = True
  End If

  Dim SQL2 As Object
  Set SQL2 = NewQuery
  Dim FILIAL As Long
  Dim FILIALPROC As Long
  Dim Msg As String

  SQL2.Add("SELECT DATADESCREDENCIAMENTO FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
  AdicionarParametroPrestadorNaQuery(SQL2)

  SQL2.Active = True

  If Not SQL2.FieldByName("DATADESCREDENCIAMENTO").IsNull Then
    BuscarFiliais(CurrentSystem, FILIAL, FILIALPROC, Msg)
    CurrentQuery.FieldByName ("NOVAFILIAL").Value = FILIAL
    NOVAFILIAL.ReadOnly = True
  End If
End Sub
Public Sub AdicionarParametroPrestadorNaQuery(SQL As Object)
  If WebMode And CurrentEntity.TransitoryVars("HandlePrestadorCredenciamentoWizard").IsPresent Then
  SQL.ParamByName("HANDLE").Value = CurrentEntity.TransitoryVars("HandlePrestadorCredenciamentoWizard").AsInteger
  Else
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR")
  End If
End Sub

Public Sub VerificarSolicitacaoDocumentos()

	On Error GoTo Exception

	Dim componente As CSBusinessComponent
	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcCredenBLL, Benner.Saude.Prestadores.Business")
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.Execute("VerificarCondicoesSolicitacaoDocumentos")
	Set componente = Nothing

	Exit Sub

	Exception:
    	Set componente = Nothing
    	bsShowMessage(Err.Description, "I")
    	Exit Sub

End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

  If (CommandID = "BOTAOGERAROFICIODOCUMENTOS") Then
  	BOTAOGERAROFICIODOCUMENTOS_OnClick
  End If

  If (CommandID = "BOTAOSOLICITARDOCUMENTOS") Then
  	VerificarSolicitacaoDocumentos
  End If


End Sub
