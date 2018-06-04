'HASH: C9C3E4BC6FCA9337F52ADCE88F98616A
'SAM_ROTINASUSPENSAO
'#Uses "*bsShowMessage"

Public Sub HabilitaCampos()
	DATAINICIALSUSPENSAO.ReadOnly        = CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString = "S"
	DATAINICIALSUSPENSAO.ReadOnly        = CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString = "S"
	DATAFINALSUSPENSAO.ReadOnly	         = CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString = "S"
	COMPETENCIAINICIALSUSPENSAO.ReadOnly = CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString = "S"
	COMPETENCIAFINALSUSPENSAO.ReadOnly   = CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString = "S"
	QTDFATURASATRASO.ReadOnly			 = CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString = "S"
	MOTIVOSUSPENSAO.ReadOnly			 = CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString = "S"
End Sub


Public Function ChecaDataFechamento()As Boolean
  Dim qFechamento As Object
  Set qFechamento = NewQuery
  Dim vMesComp As Integer
  Dim vAnoComp As Integer
  Dim vMesComp1 As Integer
  Dim vAnoComp1 As Integer
  Dim vMesFechamento As Integer
  Dim vAnoFechamento As Integer

  ChecaDataFechamento = True


  qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  qFechamento.Active = True

  vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").AsDateTime)
  vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").AsDateTime)

  vMesFechamento = DatePart("m", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
  vAnoFechamento = DatePart("yyyy", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)



  If CurrentQuery.FieldByName("DATAINICIALSUSPENSAO").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
    bsShowMessage("Não é possível suspender contratos com data inicial de suspensão inferior a data de fechamento", "E")
    ChecaDataFechamento = False
    Exit Function
  End If

  If Not CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").IsNull Then
    If(vAnoComp <vAnoFechamento)Or _
       (vAnoComp = vAnoFechamento And vMesComp <vMesFechamento)Then
    bsShowMessage("A competência inicial não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
    ChecaDataFechamento = False
    Exit Function
  End If
End If


If Not CurrentQuery.FieldByName("DATAFINALSUSPENSAO").IsNull Then
  If CurrentQuery.FieldByName("DATAFINALSUSPENSAO").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
    bsShowMessage("Não é possível suspender contratos com data final de suspensão inferior a data de fechamento", "E")
    ChecaDataFechamento = False
    Exit Function
  End If
End If

If Not CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").IsNull Then
  vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").AsDateTime)
  vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").AsDateTime)
  If(vAnoComp <vAnoFechamento)Or _
     (vAnoComp = vAnoFechamento And vMesComp <vMesFechamento)Then
  bsShowMessage("A competência final não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
  ChecaDataFechamento = False
  Exit Function
End If
End If

End Function


Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
    bsShowMessage("Rotina não foi gerada!", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
    bsShowMessage("Rotina já foi processada!", "I")
    Exit Sub
  End If

  Dim viRetorno As Long
  Dim vsMensagemErro As String
  Dim dllBSServerExec As Object
  Dim dllInterface As Object

  If (VisibleMode) Then
	Set dllInterface = CreateBennerObject("BSINTERFACE0067.RotinasSuspensao")
	dllInterface.Cancelar(CurrentQuery.FieldByName("HANDLE").AsInteger)

    RefreshNodesWithTable("SAM_ROTINASUSPENSAO")
  ElseIf (WebMode) Then
  	If bsShowMessage("Confirma o cancelamento da rotina?", "Q") = vbNo Then
      Exit Sub
  	End If

    Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
											  	 "SamRotinaSuspensao", _
											  	 "CancelamentoSuspensao", _
											     "Cancelamento de Suspensão", _
											     CurrentQuery.FieldByName("HANDLE").AsInteger, _
											     "SAM_ROTINASUSPENSAO", _
											     "SITUACAOGERACAO", _
											     "", _
											     "", _
											     "C", _
											     False, _
											     vsMensagemErro, _
											     Null)
	If (viRetorno = 0) Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set dllInterface = Nothing
  Set dllBSServerExec = Nothing
End Sub

Public Sub BOTAOGERAR_OnClick()
  If Not CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
    bsShowMessage("Rotina já foi gerada!", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString="N" Then
    If Not ChecaDataFechamento Then
      Exit Sub
    End If
  End If

  Dim viRetorno As Long
  Dim vsMensagemErro As String
  Dim dllBSServerExec As Object
  Dim dllInterface As Object

  If (VisibleMode) Then
	Set dllInterface = CreateBennerObject("BSINTERFACE0067.RotinasSuspensao")
	dllInterface.Gerar(CurrentQuery.FieldByName("HANDLE").AsInteger)

	RefreshNodesWithTable("SAM_ROTINASUSPENSAO")
  ElseIf (WebMode) Then
    Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
											  	 "SamRotinaSuspensao", _
											  	 "GerarSuspensao", _
											     "Geração de lista de Suspensão", _
											     CurrentQuery.FieldByName("HANDLE").AsInteger, _
											     "SAM_ROTINASUSPENSAO", _
											     "SITUACAOGERACAO", _
											     "", _
											     "", _
											     "P", _
											     False, _
											     vsMensagemErro, _
											     Null)
	If (viRetorno = 0) Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set dllInterface = Nothing
  Set dllBSServerExec = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.FieldByName("USUARIOGERACAO").IsNull Then
    bsShowMessage("Rotina não foi gerada!", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
    bsShowMessage("Rotina já foi processada!", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString="N" Then
    If Not ChecaDataFechamento Then
      Exit Sub
    End If
  End If

  Dim viRetorno As Long
  Dim vsMensagemErro As String
  Dim dllBSServerExec As Object
  Dim dllInterface As Object

  If (VisibleMode) Then
	Set dllInterface = CreateBennerObject("BSINTERFACE0067.RotinasSuspensao")
	dllInterface.Processar(CurrentQuery.FieldByName("HANDLE").AsInteger)

	RefreshNodesWithTable("SAM_ROTINASUSPENSAO")
  ElseIf (WebMode) Then
    Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
											  	 "SamRotinaSuspensao", _
											  	 "ProcessarSuspensao", _
											     "Processamento de Suspensão", _
											     CurrentQuery.FieldByName("HANDLE").AsInteger, _
											     "SAM_ROTINASUSPENSAO", _
											     "SITUACAOPROCESSAMENTO", _
											     "", _
											     "", _
											     "P", _
											     False, _
											     vsMensagemErro, _
											     Null)
	If (viRetorno = 0) Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set dllInterface = Nothing
  Set dllBSServerExec = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	If (VisibleMode) Then
		HabilitaCampos
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qFechamento As Object
  Dim q1 As Object
  Set q1 = NewQuery
  Set qFechamento = NewQuery
  Dim vMesComp As Integer
  Dim vAnoComp As Integer
  Dim vMesComp1 As Integer
  Dim vAnoComp1 As Integer
  Dim vMesFechamento As Integer
  Dim vAnoFechamento As Integer

  If (CurrentQuery.FieldByName("QTDFATURASATRASO").IsNull) And (CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString="N") Then
    bsShowMessage("Qtd. de faturas em atraso é obrigatório!", "E")
    CanContinue = False
    Exit Sub
  End If


  If(Not CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").IsNull)And(CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").IsNull)Then
    CanContinue = False
    bsShowMessage("A competência final não pode ser preenchida quando a competência inicial for nula.", "E")
    Exit Sub
  End If

q1.Active = False
q1.Clear
q1.Add("SELECT SUSPENDEFATURAMENTO FROM SAM_MOTIVOSUSPENSAO WHERE HANDLE = :HANDLE")
q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MOTIVOSUSPENSAO").AsInteger
q1.Active = True

If (CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString="N") And _
   (CurrentQuery.FieldByName("MOTIVOSUSPENSAO").IsNull) Then
   bsShowMessage("Motivo de suspensão é obrigatório.", "E")
   CanContinue = False
   Exit Sub
End If

If (CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString="N") And (q1.FieldByName("SUSPENDEFATURAMENTO").AsString = "S") And (CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").IsNull) Then
  bsShowMessage("Competência inicial é obrigatório.", "E")
  CanContinue = False
  Exit Sub
ElseIf ((q1.FieldByName("SUSPENDEFATURAMENTO").AsString <>"S") Or (CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString<>"N"))  And (Not CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").IsNull)Then
  bsShowMessage("Competência inicial deve ser nula.", "E")
  CanContinue = False
  Exit Sub
End If


CanContinue = True

qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
qFechamento.Active = True

vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").AsDateTime)
vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").AsDateTime)

vMesFechamento = DatePart("m", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
vAnoFechamento = DatePart("yyyy", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

If CurrentQuery.State = 3 Then
  If (CurrentQuery.FieldByName("UTILIZARPARAMETROCONTRATO").AsString="N") And (CurrentQuery.FieldByName("DATAINICIALSUSPENSAO").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime) Then
    bsShowMessage("Não é possível suspender contratos com data inicial de suspensão inferior a data de fechamento", "E")
    CanContinue = False
  End If
  If Not CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").IsNull Then
    If(vAnoComp <vAnoFechamento)Or _
       (vAnoComp = vAnoFechamento And vMesComp <vMesFechamento)Then
       bsShowMessage("A competência inicial não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
       CanContinue = False
     End If
  End If
  If Not CanContinue Then
    Exit Sub
  End If
End If

If Not CurrentQuery.FieldByName("DATAFINALSUSPENSAO").IsNull Then
  If CurrentQuery.FieldByName("DATAFINALSUSPENSAO").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
    bsShowMessage("Não é possível suspender contratos com data final de suspensão inferior a data de fechamento", "E")
    CanContinue = False
  End If
End If

If Not CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").IsNull Then
  vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").AsDateTime)
  vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").AsDateTime)
  If(vAnoComp <vAnoFechamento)Or _
     (vAnoComp = vAnoFechamento And vMesComp <vMesFechamento)Then
  bsShowMessage("A competência final não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
  CanContinue = False
End If
End If

If(Not CurrentQuery.FieldByName("DATAFINALSUSPENSAO").IsNull)And _
   (CurrentQuery.FieldByName("DATAFINALSUSPENSAO").AsDateTime <CurrentQuery.FieldByName("DATAINICIALSUSPENSAO").AsDateTime)Then
bsShowMessage("A Data final de suspensão, se informada, deve ser maior ou igual a inicial", "E")
CanContinue = False
End If

If(Not CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").IsNull)And _
   (CurrentQuery.FieldByName("COMPETENCIAFINALSUSPENSAO").AsDateTime <CurrentQuery.FieldByName("COMPETENCIAINICIALSUSPENSAO").AsDateTime)Then
bsShowMessage("A Competência final, se informada, deve ser maior ou igual a inicial", "E")
CanContinue = False
End If

End Sub


Public Sub UTILIZARPARAMETROCONTRATO_OnChange()
    CurrentQuery.UpdateRecord
	HabilitaCampos
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOPROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	ElseIf CommandID = "BOTAOCANCELAR" Then
		BOTAOCANCELAR_OnClick
	ElseIf CommandID = "BOTAOGERAR" Then
		BOTAOGERAR_OnClick
	End If
End Sub
