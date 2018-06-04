'HASH: 773D4442155D19F695D7F1D28D27466C
 
'#Uses "*bsShowMessage"

Public Sub BOTAOIMPORTAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
	bsShowMessage("Os parâmetros não podem estar em edição", "I")
	Exit Sub
  End If

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0053.ImportacaoExtrato")
    Obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
    Dim vsMensagemErro As String
    Dim viRetorno      As Long
    Dim qSelecao       As Object

    Set qSelecao = NewQuery

    qSelecao.Add("SELECT IMP.SEQUENCIA,")
    qSelecao.Add("       TES.DESCRICAO TESOURARIA")
    qSelecao.Add("FROM SFN_IMPORTACAOEXTRATOBANCARIO IMP")
    qSelecao.Add("JOIN SFN_TESOURARIA                TES ON TES.HANDLE = IMP.TESOURARIA")
    qSelecao.Add("WHERE IMP.HANDLE = :HIMPORTACAOEXTRATO")
    qSelecao.ParamByName("HIMPORTACAOEXTRATO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSelecao.Active = True

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSFIN008", _
                                     "ImportacaoExtrato_Processar", _
                                     "Importação de Extrato Bancário -" + _
                                       " Tesouraria: "  + qSelecao.FieldByName("TESOURARIA").AsString + _
                                       " Sequência: "   + qSelecao.FieldByName("SEQUENCIA").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_IMPORTACAOEXTRATOBANCARIO", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagemErro, _
                                     Null)
    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If

    Set qSelecao = Nothing
  End If

  Set Obj = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  DATAINICIAL.ReadOnly = False
  DATAFINAL.ReadOnly   = False
  FORMATO.ReadOnly     = False
  ARQUIVO.ReadOnly     = False

  Dim qSelecao As Object

  Set qSelecao = NewQuery

  qSelecao.Add("SELECT BNC.FORMATOPADRAOIMPORTACAOEXTRATO")
  qSelecao.Add("FROM SFN_TESOURARIA TES")
  qSelecao.Add("JOIN SFN_BANCO      BNC ON BNC.HANDLE = TES.BANCO")
  qSelecao.Add("WHERE TES.HANDLE =:HTESOURARIA")
  qSelecao.ParamByName("HTESOURARIA").AsInteger = CurrentQuery.FieldByName("TESOURARIA").AsInteger
  qSelecao.Active = True

  If Not qSelecao.FieldByName("FORMATOPADRAOIMPORTACAOEXTRATO").IsNull Then
    CurrentQuery.FieldByName("FORMATO").AsInteger = qSelecao.FieldByName("FORMATOPADRAOIMPORTACAOEXTRATO").AsInteger
  End If

  Set qSelecao = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.State = 1 Then
    If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
      DATAINICIAL.ReadOnly = False
      DATAFINAL.ReadOnly   = False
      FORMATO.ReadOnly     = False
      ARQUIVO.ReadOnly     = False
    Else
      DATAINICIAL.ReadOnly = True
      DATAFINAL.ReadOnly   = True
      FORMATO.ReadOnly     = True
      ARQUIVO.ReadOnly     = True
    End If
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If Not (CurrentQuery.FieldByName("SITUACAO").AsString = "1" And _
          CurrentQuery.FieldByName("SITUACAO").AsString = "5") Then
    If bsShowMessage("Se o registro de importação for excluído os lançamentos vinculados a ele também serão excluídos! Deseja continuar?", "Q") = vbYes Then
      Dim vvBSFin008     As Object
      Dim vsMensagemErro As String

      Set vvBSFin008 = CreateBennerObject("BSFin008.Rotinas")

      If vvBSFin008.ExcluirLancamentosExtratoImportado(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemErro) Then
        CancContinue = False
        BSShowMessage(vsMensagemErro, "E")
      End If

      Set vvBSFin008 = Nothing
    Else
      If VisibleMode Then
        'Pelo Runner é necessário forçar o Abort da exclusão
        'No WES esse tratamento é feito automaticamente pela resposta negativa ao questionamento acima
        CancContinue = False
      End If
    End If
  Else
    CancContinue = False
    bsShowMessage("Somente registros de importação com situação de processamento em aberto ou processados podem ser excluídos", "E")
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
    CanContinue = False
    bsShowMessage("Apenas registros de importação em aberto podem ser alterados!", "E")
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "BOTAOIMPORTAR" Then
    BOTAOIMPORTAR_OnClick
  End If
End Sub
