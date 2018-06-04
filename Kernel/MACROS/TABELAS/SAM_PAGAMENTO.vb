'HASH: F3105810B1547C1D637FA9B4529D7B57
'Macro: SAM_PAGAMENTO
'#Uses "*bsShowMessage"

' Mauricio Ibelli -09/08/2001 -sms2978 -Data de pagamento por tipo de prestador
Dim vgdataanterior As Date

Public Sub BOTAOABRIR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Registro esta em edição.", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull Then
    bsShowMessage("Calendário geral Aberto.", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    bsShowMessage("Calendário geral Já Processado.", "I")
    Exit Sub
  End If

  Dim Qu As Object
  Set Qu = NewQuery

  If Not InTransaction Then StartTransaction

  Qu.Add("UPDATE SAM_PAGAMENTO SET DATAFECHAMENTO = :DATAFECHAMENTO, USUARIOFECHAMENTO = :USUARIOFECHAMENTO WHERE HANDLE = :HANDLE")
  Qu.ParamByName("DATAFECHAMENTO").DataType = ftDateTime
  Qu.ParamByName("DATAFECHAMENTO").Clear
  Qu.ParamByName("USUARIOFECHAMENTO").DataType = ftInteger
  Qu.ParamByName("USUARIOFECHAMENTO").Clear
  Qu.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Qu.ExecSQL

  If InTransaction Then Commit

  CurrentQuery.Active = False
  CurrentQuery.Active = True

  'Balani SMS 48043 01/08/2005
  Set Qu = Nothing
End Sub

Public Sub BOTAOATUALIZARPAGAMENTO_OnClick()

 On Error GoTo Erro

 	Dim callEntity As CSEntityCall
 	Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPagamento, Benner.Saude.Entidades", "AtualizarPagamentos")
 	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
 	callEntity.Execute
 	Set callEntity =  Nothing
 	bsShowMessage("Processo enviado para execução no servidor", "I")
 	RefreshNodesWithTable("SAM_PAGAMENTO")

 Exit Sub

  Erro:
 bsShowMessage("Problema ao Fechar Pagamentos: " + Err.Description, "I")


End Sub

Public Sub BOTAOFECHAR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Registro esta em edição.", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull Then
    bsShowMessage("Calendário geral já Fechado.", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    bsShowMessage("Calendário geral já Processado.", "I")
    Exit Sub
  End If

  Dim Qu As Object
  Set Qu = NewQuery

  If Not InTransaction Then StartTransaction

  Qu.Add("UPDATE SAM_PAGAMENTO SET DATAFECHAMENTO = :DATAFECHAMENTO, USUARIOFECHAMENTO = :USUARIOFECHAMENTO WHERE HANDLE = :HANDLE")
  Qu.ParamByName("DATAFECHAMENTO").AsDateTime = ServerNow
  Qu.ParamByName("USUARIOFECHAMENTO").AsInteger = CurrentUser
  Qu.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  Qu.ExecSQL

  If InTransaction Then Commit

  CurrentQuery.Active = False
  CurrentQuery.Active = True
  'Balani SMS 48043 01/08/2005
  Set Qu = Nothing

End Sub

Public Sub BOTAOFECHARPAGAMENTOS_OnClick()

  On Error GoTo Erro

  Dim callEntity As CSEntityCall
  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPagamento, Benner.Saude.Entidades", "FecharPagamentos")
  callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
  callEntity.Execute
  Set callEntity =  Nothing
  bsShowMessage("Processo enviado para execução no servidor", "I")
  RefreshNodesWithTable("SAM_PAGAMENTO")
  Exit Sub

  Erro:
	bsShowMessage("Problema ao Fechar Pagamentos: " + Err.Description, "I")

End Sub

Public Sub BOTAOGERARPAGAMENTOS_OnClick()
  On Error GoTo Erro

  Dim callEntity As CSEntityCall
  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPagamento, Benner.Saude.Entidades", "GerarPagamentos")
  callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
  callEntity.Execute
  Set callEntity =  Nothing
  bsShowMessage("Processo enviado para execução no servidor", "I")
  RefreshNodesWithTable("SAM_PAGAMENTO")
  Exit Sub

  Erro:
	bsShowMessage("Problema ao Gerar Pagamentos: " + Err.Description, "I")

End Sub

Public Sub BOTAOEXPORTARPREVIAPAGAMENTO_OnClick()
  On Error GoTo Erro

  Dim callEntity As CSEntityCall
  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPagamento, Benner.Saude.Entidades", "ExportarPreviaPagamento")
  callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
  callEntity.Execute
  Set callEntity =  Nothing
  bsShowMessage("Processo enviado para execução no servidor", "I")
  RefreshNodesWithTable("SAM_PAGAMENTO")
  Exit Sub

  Erro:
	bsShowMessage("Problema ao Exportar Prévia de Pagamento: " + Err.Description, "I")

End Sub


Public Sub BOTAOLIBERARTETOALCADA_OnClick()
  Dim Interface As Object

  Set Interface = CreateBennerObject("LiberacaoTetoAlcadaPagamento.LiberacaoTetoAlcadaPagamento")
  Interface.Exec(0, CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)
  Set Interface = Nothing

End Sub

Public Sub BOTAOREPROCESSARPAGAMENTOS_OnClick()

  On Error GoTo Erro

  Dim callEntity As CSEntityCall
  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPagamento, Benner.Saude.Entidades", "ReprocessarPagamentos")
  callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
  callEntity.Execute
  Set callEntity =  Nothing
  bsShowMessage("Processo enviado para execução no servidor", "I")
  RefreshNodesWithTable("SAM_PAGAMENTO")
  Exit Sub

  Erro:
	bsShowMessage("Problema ao Reprocessar Pagamentos: " + Err.Description, "I")


End Sub

Public Sub BOTAOEXCLUIRPAGAMENTOS_OnClick()
  On Error GoTo Erro

  Dim callEntity As CSEntityCall
  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPagamento, Benner.Saude.Entidades", "ExcluirPagamentos")
  callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
  callEntity.Execute
  Set callEntity =  Nothing
  bsShowMessage("Processo enviado para execução no servidor", "I")
  RefreshNodesWithTable("SAM_PAGAMENTO")
  Exit Sub

  Erro:
	bsShowMessage("Problema ao Excluir Pagamentos: " + Err.Description, "I")

End Sub

Public Sub TABLE_AfterScroll()
  If ((CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull) And (CurrentQuery.FieldByName("USUARIOFECHAMENTO").IsNull)) Then
    BOTAOGERARPAGAMENTOS.Visible = True
    BOTAOFECHARPAGAMENTOS.Visible = True
    BOTAOEXPORTARPREVIAPAGAMENTO.Visible = True
    BOTAOREPROCESSARPAGAMENTOS.Visible = True
    BOTAOEXCLUIRPAGAMENTOS.Visible = True
    BOTAOLIBERARTETOALCADA.Enabled = True
  Else
    BOTAOGERARPAGAMENTOS.Visible = False
    BOTAOFECHARPAGAMENTOS.Visible = False
    BOTAOEXPORTARPREVIAPAGAMENTO.Visible = False
    BOTAOREPROCESSARPAGAMENTOS.Visible = False
    BOTAOEXCLUIRPAGAMENTOS.Visible = False
    BOTAOLIBERARTETOALCADA.Enabled = False
  End If


  If Not (CurrentQuery.FieldByName("USUARIOFECHAMENTO").IsNull And CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull) Then
  	BOTAOATUALIZARPAGAMENTO.Visible = False
  Else
    BOTAOATUALIZARPAGAMENTO.Visible = True
  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull Then
  Else
    CanContinue = False
    bsShowMessage("Data de pagamento fechada. Não permitido exclusão", "E")
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If CurrentQuery.State = 1 Then
    If Not(CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull)Then
      CanContinue = False
      bsShowMessage("Calendário já fechado", "E")
      Exit Sub
    End If
  End If

  vgdataanterior = CurrentQuery.FieldByName("datapagamento").AsDateTime
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  If CurrentQuery.State = 2 Then
    If Not(CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull)Then
      CanContinue = False
      Exit Sub
    End If

    If vgdataanterior <>CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime Then
      bsShowMessage("Os PEG's com as datas de pagamento " + CStr(vgdataanterior) + " não serão pagos.", "I")
    End If
  End If

  Set Interface = CreateBennerObject("samcalendariopgto.ROTINAS")
  'Interface.INICIALIZAR(CurrentSystem) Balani SMS 48043 02/08/2005
  If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime <>Interface.DIAUTILANTERIOR(CurrentSystem, CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)Then
    bsShowMessage("Entre com um dia útil para a Data de Pagamento", "E")
    DATAPAGAMENTO.SetFocus
    'Interface.FINALIZAR Balani SMS 48043 02/08/2005
    Set Interface = Nothing
    CanContinue = False
    Exit Sub
  End If

  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_PAGAMENTO WHERE DATAPAGAMENTO = :DATA and handle <> :handle")
  Q.ParamByName("DATA").Value = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime
  Q.ParamByName("handle").Value = CurrentQuery.FieldByName("handle").AsInteger
  Q.Active = True
  If Not Q.EOF Then
    bsShowMessage("Data de pagamento já cadastrada.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "BOTAOABRIR" Then
    BOTAOABRIR_OnClick
  ElseIf CommandID = "BOTAOFECHAR" Then
    BOTAOFECHAR_OnClick
  ElseIf CommandID = "BOTAOGERARPAGAMENTOS" Then
    BOTAOGERARPAGAMENTOS_OnClick
  ElseIf CommandID = "BOTAOFECHARPAGAMENTOS" Then
    BOTAOFECHARPAGAMENTOS_OnClick
  ElseIf CommandID = "BOTAOEXPORTARPREVIAPAGAMENTO" Then
    BOTAOEXPORTARPREVIAPAGAMENTO_OnClick
  ElseIf CommandID = "BOTAOREPROCESSARPAGAMENTOS" Then
    BOTAOREPROCESSARPAGAMENTOS_OnClick
  ElseIf CommandID = "BOTAOEXCLUIRPAGAMENTOS" Then
    BOTAOEXCLUIRPAGAMENTOS_OnClick
  ElseIf CommandID = "BOTAOLIBERARTETOALCADA" Then
  	BOTAOLIBERARTETOALCADA_OnClick
  ElseIf CommandID = "BOTAOATUALIZARPAGAMENTO" Then
  	BOTAOATUALIZARPAGAMENTO_OnClick
  End If
End Sub
