'HASH: 6D82241DC4CF3786FB83C8A93FC0FFD9
'Macro: ANS_SIP_COMPETENCIA
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj, qConsulta As Object
  Dim vMensagem As String

  Set qConsulta = NewQuery

  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  qConsulta.Active = False
  qConsulta.Clear
  qConsulta.Add("SELECT MIN(DATAHORA) PRIMEIRAOCORRENCIA")
  qConsulta.Add("  FROM ANS_SIP_COMPETENCIA_OCORRENCIA  ")
  qConsulta.Add(" WHERE SIPCOMPET = :SIPCOMPET          ")
  qConsulta.ParamByName("SIPCOMPET").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qConsulta.Active = True

  If CurrentQuery.FieldByName("SITUACAO").AsInteger = 5 Then
  	If bsShowMessage("Deseja realmente cancelar o Anexo?", "Q") = vbYes Then
  		Set Obj = CreateBennerObject("BSANS001.Cancelar")
  		Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,vMensagem)
  		Set Obj = Nothing
    	bsShowMessage(vMensagem, "I")
  	End If
  ElseIf CurrentQuery.FieldByName("SITUACAO").AsInteger = 4 Then
  	Dim vsVerifica As String
  	vsVerifica = "A rotina está processando há mais de 1 dia. Provavelmente a mesma não esteja mais em execução e poderá ser cancelada sem maiores problemas. Deseja continuar?"

  	If qConsulta.FieldByName("PRIMEIRAOCORRENCIA").AsDateTime + 1 < ServerNow Then
    	If bsShowMessage(vsVerifica, "Q") = vbYes Then
  		Set Obj = CreateBennerObject("BSANS001.Cancelar")
  		Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,vMensagem)
  		Set Obj = Nothing
    	bsShowMessage(vMensagem, "I")
  	End If
	Else
		bsShowMessage("A rotina está processando e não é permitido o cancelamento.","I")
	End If
  Else
  	bsShowMessage("Situação não permite cancelamento!","I")
  End If

  If Not WebMode Then
	RefreshNodesWithTable("ANS_SIP_COMPETENCIA")
  End If
End Sub

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object
  Dim vmensagemretorno As String
  Dim viRetorno As Long

  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If VisibleMode Then
  	Set Obj = CreateBennerObject("BSINTERFACE0055.Rotinas")
  	Obj.GerarAnexo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vmensagemretorno)
  	Set Interface = Nothing
  	GerarXml
  Else
  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "BSANS001", _
                                    "GerarAnexo", _
                                    "SIP - Sistema de informações de Produtos - Gerar", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "ANS_SIP_COMPETENCIA", _
                                    "SITUACAO", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vmensagemretorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
  		GerarXml
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vmensagemretorno, "I")
  	 End If

  End If

  If Not WebMode Then
	RefreshNodesWithTable("ANS_SIP_COMPETENCIA")
  End If
End Sub


Public Sub BOTAOGERARPLANILHA_OnClick()
	Dim Obj, qAnexo As Object
	Set qAnexo = NewQuery

	qAnexo.Active = False
	qAnexo.Clear
	qAnexo.Add("SELECT ANEXO, R_RELATORIOS.NOME RELATORIOANEXOI FROM ANS_SIP_ANEXO ")
	qAnexo.Add("JOIN R_RELATORIOS ON R_RELATORIOS.HANDLE = ANS_SIP_ANEXO.RELATORIOANEXOI ")
	qAnexo.Add("WHERE ANS_SIP_ANEXO.HANDLE = :ANEXO")
	qAnexo.ParamByName("ANEXO").AsInteger = CurrentQuery.FieldByName("ANEXO").AsInteger
	qAnexo.Active = True

	If qAnexo.FieldByName("ANEXO").AsString <> "1" Then

		If CurrentQuery.State <> 1 Then
  			bsShowMessage("Os parâmetros não podem estar em edição", "I")
			Exit Sub
		End If

		Set Obj = CreateBennerObject("BSANS001.uRotinas")
		Obj.CriarPlanilha(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
		Set Obj = Nothing
	Else
	  	Dim relatorio As CSReportPrinter
		Dim qhRelatorio As Object
		Set qhRelatorio = NewQuery

		qhRelatorio.Add(" SELECT HANDLE")
		qhRelatorio.Add(" FROM R_RELATORIOS")
		qhRelatorio.Add("  WHERE NOME = :RNOME ")
		qhRelatorio.ParamByName("RNOME").AsString = qAnexo.FieldByName("RELATORIOANEXOI").AsString
		qhRelatorio.Active = True

		Set relatorio = NewReport(qhRelatorio.FieldByName("HANDLE").AsInteger)
		relatorio.CanFilter = False
		relatorio.Preview

		Set relatorio = Nothing
		Set qhRelatorio = Nothing
	End If
	Set qAnexo = Nothing
End Sub

Public Sub BOTAOGERARXML_OnClick()
  GerarXml
End Sub

Public Sub GerarXml()
	Dim Obj As Object
	Dim vsMensagem As String
	Set Obj = CreateBennerObject("Benner.Saude.ANS.SIPRotinas")
  	vsMensagem = Obj.criaXML(CurrentQuery.FieldByName("HANDLE").AsInteger,CurrentSystem)
	bsShowMessage(vsMensagem,"I")
  	Set Obj = Nothing
End Sub

Public Sub BOTAOINCONSISTENCIAS_OnClick()
  Dim Obj As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If
  If VisibleMode Then
  	Set Obj = CreateBennerObject("BSINTERFACE0055.VerificarInconsistencias")
  	Obj.VerificarInconsistencias(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set Interface = Nothing
  Else
  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "BSANS001", _
                                    "VerificarInconsistencias", _
                                    "SIP - Sistema de informações de Produtos - Verificar Inconsistências", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "ANS_SIP_COMPETENCIA", _
                                    "SITUACAO", _
                                    "", _
                                    "", _
                                    "C", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If
  End If

  Set Obj = Nothing
End Sub



Public Sub TABLE_AfterPost()
	Dim Obj, qAnexo As Object
	Set qAnexo = NewQuery

	qAnexo.Active = False
	qAnexo.Clear
	qAnexo.Add("SELECT ANEXO FROM ANS_SIP_ANEXO WHERE HANDLE = :ANEXO")
	qAnexo.ParamByName("ANEXO").AsInteger = CurrentQuery.FieldByName("ANEXO").AsInteger
	qAnexo.Active = True

	If qAnexo.FieldByName("ANEXO").AsString = "1" Then
		BOTAOGERARPLANILHA.Caption = "Gerar Relatório"
	Else
		BOTAOGERARPLANILHA.Caption = "Gerar Planilha"
	End If

	Set qAnexo = Nothing
End Sub

Public Sub TABLE_AfterScroll()
    CONVENIO.ReadOnly = Not (CurrentQuery.FieldByName("SITUACAO").AsString = "1" Or CurrentQuery.FieldByName("SITUACAO").AsString = "")
	If (CurrentQuery.State = 1) And Not(CurrentQuery.FieldByName("HANDLE").IsNull) Then
		BOTAOGERAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString = "1")
		BOTAOCANCELAR.Enabled = ((CurrentQuery.FieldByName("SITUACAO").AsString = "5") Or (CurrentQuery.FieldByName("SITUACAO").AsString = "4"))
		BOTAOGERARPLANILHA.Enabled = BOTAOCANCELAR.Enabled
	End If

	If CurrentQuery.State <> 3 Then
		Dim qAnexo As Object
		Set qAnexo = NewQuery

		qAnexo.Active = False
		qAnexo.Clear
		qAnexo.Add("SELECT ANEXO FROM ANS_SIP_ANEXO WHERE HANDLE = :ANEXO")
		qAnexo.ParamByName("ANEXO").AsInteger = CurrentQuery.FieldByName("ANEXO").AsInteger
		qAnexo.Active = True

		If qAnexo.FieldByName("ANEXO").AsString = "1" Then
			BOTAOGERARPLANILHA.Caption = "Gerar Relatório"
		Else
			BOTAOGERARPLANILHA.Caption = "Gerar Planilha"
		End If

		Set qAnexo = Nothing
	End If
End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOGERAR"
			BOTAOGERAR_OnClick
		Case "BOTAOGERARXML"
			BOTAOGERARXML_OnClick
		Case "BOTAOINCONSISTENCIAS"
			BOTAOINCONSISTENCIAS_OnClick
	End Select

End Sub
