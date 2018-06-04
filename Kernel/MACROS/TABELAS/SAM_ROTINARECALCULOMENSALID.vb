'HASH: F3878ED67BD1D4EE654EEC79D21416A5

'SAM_ROTINARECALCULOMENSALID

'#Uses "*ProcuraBeneficiarioAtivo"
'#Uses "*bsShowMessage"

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraBeneficiarioAtivo(True, ServerDate, BENEFICIARIO.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
End Sub

Public Sub BOTAOAGENDAR_OnClick()
  Dim qr As Object
  Dim qr1 As Object
  Dim vSituacaoProcessamento As String
  Dim vSituacaoFaturamento As String
  Dim vTabela As String
  Set qr = NewQuery
  Set qr1 = NewQuery
  vTabela = "SAM_ROTINARECALCULOMENSALID"
  qr.Clear
  qr.Add("SELECT SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO FROM " + vTabela + " WHERE HANDLE = :pHANDLE")
  qr.ParamByName("pHandle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qr.Active = True

  vSituacaoProcessamento = qr.FieldByName("SITUACAOPROCESSAMENTO").AsString
  vSituacaoFaturamento = qr.FieldByName("SITUACAOFATURAMENTO").AsString

  If vSituacaoProcessamento <> "3" And vSituacaoFaturamento <> "3" Then
    If vSituacaoProcessamento = "1" Or (vSituacaoProcessamento = "5" And vSituacaoFaturamento = "1") Then
      If (vSituacaoProcessamento = "5" And vSituacaoFaturamento = "1") Then
        If CurrentQuery.FieldByName("DATAVENCIMENTO").IsNull And CurrentQuery.FieldByName("DATACONTABIL").IsNull Then
		  	bsShowMessage("A data de vencimento e a data contábil deve ser informada", "I")
		  	Exit Sub
		Else
		  	If CurrentQuery.FieldByName("DATAVENCIMENTO").IsNull Then
			  bsShowMessage("A data de vencimento deve ser informada", "I")
			  DATAVENCIMENTO.SetFocus
			  Exit Sub
			End If

			If CurrentQuery.FieldByName("DATACONTABIL").IsNull Then
			  bsShowMessage("A data contábil deve ser informada", "I")
			  DATACONTABIL.SetFocus
			  Exit Sub
			End If
		End If
      End If
      If bsShowMessage("Confirme o agendamento da rotina", "Q") = vbYes Then '(6=yes, 7=não)
        qr1.Clear
        qr1.Add("UPDATE " + vTabela + " SET SITUACAOPROCESSAMENTO = :pSituacao, SITUACAOFATURAMENTO = :pSituacao WHERE HANDLE = :pHANDLE")
        qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qr1.ParamByName("pSituacao").AsString = "3"
        qr1.ExecSQL
      End If
    Else
      bsShowMessage("Rotina já foi faturada.", "E")
    End If
  Else
    If bsShowMessage("Rotina já está agendada. Para retirar o agendamento pressione 'SIM'", "Q") = vbYes Then
      qr1.Clear
      qr1.Add("UPDATE " + vTabela + " SET SITUACAOPROCESSAMENTO = :pSituacaoP, SITUACAOFATURAMENTO = :pSituacaoF WHERE HANDLE = :pHANDLE")
      qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      If (CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull) Then
        qr1.ParamByName("pSituacaoP").AsString = "1"
        qr1.ParamByName("pSituacaoF").AsString = "1"
      Else
        qr1.ParamByName("pSituacaoP").AsString = "5"
        qr1.ParamByName("pSituacaoF").AsString = "1"
      End If
      qr1.ExecSQL
    End If
  End If
  Set qr = Nothing
  Set qr1 = Nothing

  If VisibleMode Then
  	SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If

End Sub

Public Sub BOTAOCANCELAR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "E")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "8" Or _
     CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "6" Or _
     CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "8" Or _
     CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "6" Then
    bsShowMessage("A rotina esta aguardando cancelamento", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "2" Or _
     CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "4" Or _
     CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "2" Or _
     CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "4" Then
    bsShowMessage("Rotina em processamento não pode ser cancelada", "I")
    Exit Sub
  End If

  Dim Obj As Object
  Dim vsMensagemErro As String
  Dim viRetorno As Long

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "5" Then

	If CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "1" Then

		If VisibleMode Then

			Set Obj = CreateBennerObject("BSINTERFACE0065.RotinaRecalculo")
    		Obj.CancelarProcessamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    		viRetorno = 0

		Else

			Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                 "SamRecalcMensal", _
                                 "RotinaProcessamentoRecalculo_Cancelar", _
                                 "Cancelamento de processamento (Recalculo Mensalidade) - Rotina: " + CurrentQuery.FieldByName("HANDLE").AsString, _
                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                 "SAM_ROTINARECALCULOMENSALID", _
                                 "SITUACAOPROCESSAMENTO", _
                                 "", _
                                 "", _
                                 "C", _
                                 False, _
                                 vsMensagemErro, _
                                 Null)

            If viRetorno = 0 Then

				bsShowMessage("A rotina foi enviada para execução no servidor", "I")

  			Else

  				bsShowMessage("Erro ao enviar rotina para o servidor" + vsMensagemErro, "E")

  			End If


		End If


  	ElseIf CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "5" Then

  		If VisibleMode Then

			Set Obj = CreateBennerObject("BSINTERFACE0065.RotinaRecalculo")
	    	Obj.CancelarFaturamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    	Else

			Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                 "SamRecalcMensal", _
                                 "RotinaFaturamentoRecalculo_Cancelar", _
                                 "Cancelamento de faturamento (Recalculo Mensalidade) - Rotina: " + CurrentQuery.FieldByName("HANDLE").AsString, _
                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                 "SAM_ROTINARECALCULOMENSALID", _
                                 "SITUACAOFATURAMENTO", _
                                 "", _
                                 "", _
                                 "C", _
                                 False, _
                                 vsMensagemErro, _
                                 Null)

            If viRetorno = 0 Then

				bsShowMessage("A rotina foi enviada para execução no servidor", "I")

  			Else

  				bsShowMessage("Erro ao enviar rotina para o servidor. " + vsMensagemErro, "E")

  			End If

    	End If

    End If

  Else

    bsShowMessage("A rotina não pôde ser cancelada", "E")

  End If

  Set Obj = Nothing

  WriteAudit("C", HandleOfTable("SAM_ROTINARECALCULOMENSALID"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Recálculo de Mensalidades - Cancelamento de faturamento")
  RefreshNodesWithTable("SAM_ROTINARECALCULOMENSALID")
End Sub

Public Sub BOTAOFATURAR_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString <> "5" Then
    bsShowMessage("A rotina não foi processada", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "5" Then
    bsShowMessage("A rotina já foi faturada", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAVENCIMENTO").IsNull And CurrentQuery.FieldByName("DATACONTABIL").IsNull Then
  	bsShowMessage("A data de vencimento e a data contábil deve ser informada", "I")
  	Exit Sub
  Else
  	If CurrentQuery.FieldByName("DATAVENCIMENTO").IsNull Then
	  bsShowMessage("A data de vencimento deve ser informada", "I")
	  DATAVENCIMENTO.SetFocus
	  Exit Sub
	End If

	If CurrentQuery.FieldByName("DATACONTABIL").IsNull Then
	  bsShowMessage("A data contábil deve ser informada", "I")
	  DATACONTABIL.SetFocus
	  Exit Sub
	End If
  End If

  Dim Obj As Object
  Dim vsMensagemErro As String
  Dim viRetorno As Long

  If VisibleMode Then

	Set Obj = CreateBennerObject("BSINTERFACE0065.RotinaRecalculo")
	Obj.Faturar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Else

	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                         "SamRecalcMensal", _
                         "RotinaFaturamentoRecalculo_Processar", _
                         "Faturamento (Recalculo Mensalidade) - Rotina: " + CurrentQuery.FieldByName("HANDLE").AsString, _
                         CurrentQuery.FieldByName("HANDLE").AsInteger, _
                         "SAM_ROTINARECALCULOMENSALID", _
                         "SITUACAOFATURAMENTO", _
                         "", _
                         "", _
                         "P", _
                         False, _
                         vsMensagemErro, _
                         Null)

    If viRetorno = 0 Then

		bsShowMessage("A rotina foi enviada para execução no servidor", "I")

	Else

		bsShowMessage("Erro ao enviar rotina para o servidor" + vsMensagemErro, "E")

	End If

  End If

  Set Obj = Nothing

  WriteAudit("F", HandleOfTable("SAM_ROTINARECALCULOMENSALID"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Recálculo de Mensalidades - Faturamento da Rotina")

  RefreshNodesWithTable("SAM_ROTINARECALCULOMENSALID")
End Sub

Public Sub BOTAOFATURARROTINAS_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "1" Then
    bsShowMessage("A rotina não foi processada", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "5" And CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "5" Then
    bsShowMessage("A rotina já foi faturada", "I")
    Exit Sub
  End If

  Dim Obj As Object
  Set Obj = CreateBennerObject("SamRecalcMensal.Rotinas")
  Obj.FaturarRecalculos(CurrentSystem)
  Set Obj = Nothing

  RefreshNodesWithTable("SAM_ROTINARECALCULOMENSALID")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "2" Then
    bsShowMessage("A rotina esta aguardando processamento", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "4" Then
    bsShowMessage("A rotina já esta sendo processada", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "5" Then
    bsShowMessage("A rotina já foi processada", "I")
    Exit Sub
  End If

  Dim Obj As Object
  Dim vsMensagemErro As String
  Dim viRetorno As Long

  If VisibleMode Then

	Set Obj = CreateBennerObject("BSINTERFACE0065.RotinaRecalculo")
	Obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Else

	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                         "SamRecalcMensal", _
                         "RotinaProcessamentoRecalculo_Processar", _
                         "Processamento (Recalculo Mensalidade) - Rotina: " + CurrentQuery.FieldByName("HANDLE").AsString, _
                         CurrentQuery.FieldByName("HANDLE").AsInteger, _
                         "SAM_ROTINARECALCULOMENSALID", _
                         "SITUACAOPROCESSAMENTO", _
                         "", _
                         "", _
                         "P", _
                         False, _
                         vsMensagemErro, _
                         Null)

    If viRetorno = 0 Then

		bsShowMessage("A rotina foi enviada para execução no servidor", "I")

	Else

		bsShowMessage("Erro ao enviar rotina para o servidor" + vsMensagemErro, "E")

	End If

  End If

  Set Obj = Nothing

  WriteAudit("P", HandleOfTable("SAM_ROTINARECALCULOMENSALID"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Recálculo de Mensalidades - Processamento da Rotina")

  RefreshNodesWithTable("SAM_ROTINARECALCULOMENSALID")
End Sub

Public Sub BOTAOPROCESSARROTINAS_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "5" And CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "1" Then
    bsShowMessage("A rotina já foi processada", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "5" And CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "5" Then
    bsShowMessage("A rotina já foi faturada", "I")
    Exit Sub
  End If

  Dim Obj As Object
  Set Obj = CreateBennerObject("SamRecalcMensal.Rotinas")
  Obj.ProcessarRecalculos(CurrentSystem)
  Set Obj = Nothing

  RefreshNodesWithTable("SAM_ROTINARECALCULOMENSALID")
End Sub

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vOrdem As Integer

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  'vCriterio =vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  If IsNumeric(CONTRATO.LocateText) Then
    vOrdem = 1
  Else
    vOrdem = 2
  End If

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, vOrdem, vCampos, vCriterio, "Contratos", True, CONTRATO.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATO").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub CONTRATOFINAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vOrdem As Integer

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  'vCriterio =vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  If IsNumeric(CONTRATOFINAL.LocateText) Then
    vOrdem = 1
  Else
    vOrdem = 2
  End If

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, vOrdem, vCampos, vCriterio, "Contratos", True, CONTRATOFINAL.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOFINAL").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub CONTRATOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vOrdem As Integer

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  'vCriterio =vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  If IsNumeric(CONTRATOINICIAL.LocateText) Then
    vOrdem = 1
  Else
    vOrdem = 2
  End If

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, vOrdem, vCampos, vCriterio, "Contratos", True, CONTRATOINICIAL.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOINICIAL").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub FAMILIAFINAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vContrato As Integer

  vContrato = CurrentQuery.FieldByName("CONTRATO").Value
  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "FAMILIA|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND CONTRATO = " + CStr(vContrato)
  vCampos = "Nº da Familia|Data Adesão"

  vHandle = interface.Exec(CurrentSystem, "SAM_FAMILIA", vColunas, 1, vCampos, vCriterio, "Familias", True, FAMILIAFINAL.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("FAMILIAFINAL").Value = vHandle
  End If

  Set interface = Nothing
End Sub


Public Sub FAMILIAINICIAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vContrato As Integer

  vContrato = CurrentQuery.FieldByName("CONTRATO").Value

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "FAMILIA|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND CONTRATO = " + CStr(vContrato)
  vCampos = "Nº da Familia|Data Adesão"

  vHandle = interface.Exec(CurrentSystem, "SAM_FAMILIA", vColunas, 1, vCampos, vCriterio, "Familias", True, FAMILIAINICIAL.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("FAMILIAINICIAL").Value = vHandle
  End If

  Set interface = Nothing
End Sub


Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString = "1" Or CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").IsNull Then
    DESCRICAO.ReadOnly = False
    DATAROTINA.ReadOnly = False
    COMPETENCIAINICIAL.ReadOnly = False
    COMPETENCIAFINAL.ReadOnly = False
    GRUPOCONTRATO.ReadOnly = False
    CONTRATOINICIAL.ReadOnly = False
    CONTRATOFINAL.ReadOnly = False
    CONTRATO.ReadOnly = False
    FAMILIAINICIAL.ReadOnly = False
    FAMILIAFINAL.ReadOnly = False
    BENEFICIARIO.ReadOnly = False
  Else
    DESCRICAO.ReadOnly = True
    DATAROTINA.ReadOnly = True
    COMPETENCIAINICIAL.ReadOnly = True
    COMPETENCIAFINAL.ReadOnly = True
    GRUPOCONTRATO.ReadOnly = True
    CONTRATOINICIAL.ReadOnly = True
    CONTRATOFINAL.ReadOnly = True
    CONTRATO.ReadOnly = True
    FAMILIAINICIAL.ReadOnly = True
    FAMILIAFINAL.ReadOnly = True
    BENEFICIARIO.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim QAux As Object
	Set QAux = NewQuery
	QAux.Add("UPDATE SAM_CONTRATO_AUTOGESTAO                  ")
	QAux.Add("   SET HANDLEROTINARECALC = NULL                ")
	QAux.Add(" WHERE HANDLEROTINARECALC = :HANDLEROTINARECALC ")
	QAux.ParamByName("HANDLEROTINARECALC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	QAux.ExecSQL
	Set QAux = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If(Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And _
     (CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime <CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)Then
  bsShowMessage("A Competência Final , se informada, deve ser maior ou igual a inicial", "I")
  CanContinue = False
Else
  CanContinue = True
End If
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser
  CurrentQuery.FieldByName("DATAINCLUSAO").Value = ServerNow
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
		  BOTAOCANCELAR_OnClick
		Case "BOTAOPROCESSAR"
		  BOTAOPROCESSAR_OnClick
		Case "BOTAOFATURAR"
		  BOTAOFATURAR_OnClick
		Case "BOTAOAGENDAR"
		  BOTAOAGENDAR_OnClick
	End Select

End Sub
