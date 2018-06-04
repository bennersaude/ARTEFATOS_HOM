'HASH: C9D60194A36A42B6ACFD960DEEF5D9B7
'Macro: SAM_PARAMETROSATENDIMENTO

'#Uses "*ProcuraEvento"
'#uses "*bsShowMessage"

Dim gVisualiza As Boolean

Public Sub BOTAOATUALIZADIARIO_OnClick()

  '*********************************************************************************
  '****** ALTERAÇÃO PARA POS - IMPORTANDO AS GUIAS MOVIMENTO,DE ACERTO *************
  '*********************************************************************************
  Dim Interface As Object
  Set Interface = CreateBennerObject("BSATE003.Rotinas")

  Interface.ExecAutoriz(CurrentSystem)

  Set Interface = Nothing

  RefreshNodesWithTable("POS_LOTE")

  '******************* FIM DA ALTERAÇÃO ********************************************
End Sub

Public Sub EVENTOPUERICULTURA_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False

  vHandle = ProcuraEvento(True, EVENTOPUERICULTURA.Text)

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOPUERICULTURA").Value = vHandle
  End If
End Sub

Public Sub EVENTOREEMBOLSOMATMED_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False

  vHandle = ProcuraEvento(True, EVENTOREEMBOLSOMATMED.Text)

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOREEMBOLSOMATMED").Value = vHandle
  End If
End Sub

Public Sub FILTRARGRAUSVALIDOS_OnChange()
  If Not (CurrentQuery.FieldByName("FILTRARGRAUSVALIDOS").AsString = "N") Then
    CurrentQuery.FieldByName("FILTRARGRAUSVALIDOSNADIGITACAO").AsString = "N"
  End If
End Sub

Public Sub FILTRARGRAUSVALIDOSNADIGITACAO_OnChange()
  If (CurrentQuery.FieldByName("FILTRARGRAUSVALIDOS").AsString = "N") Then
    CurrentQuery.FieldByName("FILTRARGRAUSVALIDOSNADIGITACAO").AsString = "N"
  End If
End Sub

Public Sub TABLE_AfterScroll()
 AUTOTIPOAUTORIZPADRAO.LocalWhere = "(DATAINICIAL <= " + CurrentSystem.SQLDate(CurrentSystem.ServerDate) + ") AND (DATAFINAL Is Null)" 'SMS 81867 - Débora Rebello - 21/05/2007
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  gVisualiza = CurrentQuery.FieldByName("VISUALIZARHISTATENDCENTRAL").AsString = "S"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vMsg As String
  vMsg = ""
  If CurrentQuery.FieldByName("USAPOS").AsString = "S" Then
    If CurrentQuery.FieldByName("USUARIOPOS").IsNull Then
      vMsg = vMsg + "Usuario padrão do autorizador externo é obrigatório" + Chr(13)
    End If

    If CurrentQuery.FieldByName("TIPOAUTORIZACAOPOS").IsNull Then
      vMsg = vMsg + "Tipo de autorização padrão do autorizador externo é obrigatório" + Chr(13)
    End If

    If vMsg <> "" Then
      CanContinue = False
      bsShowMessage(vMsg, "E")
    End If
  End If

  If TABTIPOAUTEXTERNO.PageIndex = 1 Then
    CurrentQuery.FieldByName("USAPOS").Value = "S"
  Else
    CurrentQuery.FieldByName("USAPOS").Value = "N"
  End If

  If CurrentQuery.FieldByName("MOSTRARMENSAGEMENDERECO").AsString = "S" Then
    If CurrentQuery.FieldByName("INTERVALOENTRECONFIRMACAO").IsNull Then
      bsShowMessage("É necessário configurar o intervalo entre confirmação de endereços", "E")
      CanContinue = False

    End If
    If CurrentQuery.FieldByName("MOSTRARENDERECOCOML").AsString = "N" And _
                                 CurrentQuery.FieldByName("MOSTRARENDERECORESL").AsString = "N" And _
                                 CurrentQuery.FieldByName("MOSTRARENDERECOCOR").AsString = "N" Then
      bsShowMessage("É necessário marcar pelo menos um endereço para ser exibido", "E")
      CanContinue = False
    End If
  End If

  Dim qParamAntes As BPesquisa
  Set qParamAntes = NewQuery

  qParamAntes.Active = False
  qParamAntes.Clear
  qParamAntes.Add("SELECT FORNECIMENTOMEDICAMENTO FROM SAM_PARAMETROSATENDIMENTO")
  qParamAntes.Active = True

  If qParamAntes.FieldByName("FORNECIMENTOMEDICAMENTO").AsString = "N" And _
     qParamAntes.FieldByName("FORNECIMENTOMEDICAMENTO").AsString <> CurrentQuery.FieldByName("FORNECIMENTOMEDICAMENTO").AsString Then
    Dim qEvento As BPesquisa
    Set qEvento = NewQuery

    qEvento.Clear
    qEvento.Add("SELECT COUNT(*) REGISTROS                                    ")
    qEvento.Add("  FROM SAM_TGE E                                             ")
    qEvento.Add("  JOIN SAM_TGE_COMPLEMENTAR C ON C.EVENTO = E.HANDLE         ")
    qEvento.Add("  JOIN SAM_TGE EG ON C.EVENTOAGERAR = EG.HANDLE              ")
    qEvento.Add(" WHERE (E.TABTIPOEVENTO = '4' OR EG.TABTIPOEVENTO = '4') AND ")
    qEvento.Add("       C.EVENTOAGERAR <> C.EVENTO                            ")
    qEvento.Active = True

    If qEvento.FieldByName("REGISTROS").AsInteger > 0 Then
      CanContinue = False
      Set qEvento = Nothing
      Set qParamAntes = Nothing
      bsShowMessage("Não é possível utilizar 'Cadastro de Fornecimento'!"+ Chr(13) + _
                    "Existe(m) Evento(s) do tipo 'Medicamento' como ou com 'Evento Complementar'!","E")
 	  'CurrentQuery.FieldByName("FORNECIMENTOMEDICAMENTO").AsString = "N"
      Exit Sub
    End If

    Set qEvento = Nothing
  End If

  Set qParamAntes = Nothing
End Sub

Public Sub TIPODOCUMENTOAUTORIZACAO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_TIPODOCUMENTO")
  ShowPopup = False
  TIPODOCUMENTOAUTORIZACAO.LocalWhere = " TIPODOCUMENTOADM = 'S'"
  ShowPopup = True
End Sub

Public Sub PROCESSAR_OnClick()
  If CurrentQuery.FieldByName("GRAUPADRAOPOS").IsNull Then
    bsShowMessage("O grau padrão é obrigatório!", "E")
  Else
    Dim Obj As Object
    Set Obj = CreateBennerObject("SamPOS.Processar")

    Obj.Inicializar

    Set Obj = Nothing
  End If
End Sub

Public Sub VISUALIZARHISTATENDCENTRAL_OnChange()
  gVisualiza = Not gVisualiza
  HISTATENDAOINICIARCENTRAL.Visible = gVisualiza
End Sub
