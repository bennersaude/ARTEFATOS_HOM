'HASH: DB4EBEFBD716D962E6C7F2A0AB5B759B
'#uses "*bsShowMessage"'

Option Explicit

Public Sub BOTAOIMPRIMIRANEXO_OnClick()
  WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Início")
  Dim qParametrosAtendimento As BPesquisa
  Dim HandleRelatorio As Long

  Set qParametrosAtendimento = NewQuery

  qParametrosAtendimento.Add("SELECT TABUTILIZACONCEITOPROTOCOLO,")
  qParametrosAtendimento.Add("       PROTRELATORIOOPME,")
  qParametrosAtendimento.Add("       PROTRELATORIOINTERMEDIACAOOPME")
  qParametrosAtendimento.Add("FROM SAM_PARAMETROSATENDIMENTO")

  qParametrosAtendimento.Active = True

  HandleRelatorio = 0

  If CurrentQuery.FieldByName("INTERMEDIACAOCOMPRA").AsString = "S" Then
    WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Intermediação de Compra")
    If (qParametrosAtendimento.FieldByName("TABUTILIZACONCEITOPROTOCOLO").AsInteger = 2) And _
       (Not qParametrosAtendimento.FieldByName("PROTRELATORIOINTERMEDIACAOOPME").IsNull) Then
      WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Relatório de Intermediação de Compra")
      HandleRelatorio = qParametrosAtendimento.FieldByName("PROTRELATORIOINTERMEDIACAOOPME").AsInteger
    End If
    Set qParametrosAtendimento = Nothing
  Else
    If (qParametrosAtendimento.FieldByName("TABUTILIZACONCEITOPROTOCOLO").AsInteger = 2) And _
       (Not qParametrosAtendimento.FieldByName("PROTRELATORIOOPME").IsNull) Then
      WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Relatório de OPME")
      HandleRelatorio = qParametrosAtendimento.FieldByName("PROTRELATORIOOPME").AsInteger
    End If
  End If

  Set qParametrosAtendimento = Nothing

  If HandleRelatorio = 0 Then
    WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Buscando TISS014")
    Dim qRelatorio As BPesquisa
    Set qRelatorio = NewQuery

    qRelatorio.Active = False
    qRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'TISS014'")
    qRelatorio.Active = True

    HandleRelatorio = qRelatorio.FieldByName("HANDLE").AsInteger

    Set qRelatorio = Nothing
  End If

  If HandleRelatorio > 0 Then
    WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Relatório Encontrado: " + CStr(HandleRelatorio))
	Dim qProcedimento As BPesquisa
  	Set qProcedimento = NewQuery

	qProcedimento.Add("SELECT 1 FROM SAM_AUTORIZ_ANEXOOPME_PROC WHERE ANEXOOPME = :ANEXOOPME")
  	qProcedimento.ParamByName("ANEXOOPME").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  	qProcedimento.Active = True

	If qProcedimento.EOF Then
	  WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Sem procedimentos")
	  bsShowMessage("O anexo não pode ser impresso pois não possui procedimentos cadastrados!", "I")
	Else
	  WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Imprimindo relatório")
	  Dim rep As CSReportPrinter
      Set rep = NewReport(HandleRelatorio)
      SessionVar("ANEXOOPME") = CurrentQuery.FieldByName("HANDLE").AsString
      rep.Preview
      Set rep = Nothing
  	End If

    Set qProcedimento = Nothing
  Else
    WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Nenhum relatório encontrado")
    bsShowMessage("Nenhum relatório para impressão de anexo de solicitação de OPME foi encontrado.", "E")
  End If
  WriteBDebugMessage("SAM_AUTORIZ_ANEXOOPME.BOTAOIMPRIMIRANEXO_OnClick - Fim")
End Sub

Public Sub TABLE_AfterInsert()
  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
    Dim qAutorizacao As BPesquisa
    Set qAutorizacao = NewQuery

  	If Not CurrentQuery.FieldByName("WEBAUTORIZ").IsNull Then
	  qAutorizacao.Add("SELECT BENEFICIARIO, ATENDIMENTORECEMNATO FROM WEB_AUTORIZ WHERE HANDLE = :HANDLE")
      qAutorizacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("WEBAUTORIZ").AsInteger
   	  qAutorizacao.Active = True
  	Else
   	  qAutorizacao.Add("SELECT AUTORIZACAO, BENEFICIARIO, ATENDIMENTORECEMNATO FROM SAM_AUTORIZ WHERE HANDLE = :HANDLE")
   	  qAutorizacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
   	  qAutorizacao.Active = True
      CurrentQuery.FieldByName("NUMEROGUIAREFERENCIADA").AsString = qAutorizacao.FieldByName("AUTORIZACAO").AsString
  	End If

    CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = qAutorizacao.FieldByName("BENEFICIARIO").AsInteger
    CurrentQuery.FieldByName("ATENDIMENTORN").AsString = qAutorizacao.FieldByName("ATENDIMENTORECEMNATO").AsString
    CurrentQuery.FieldByName("DATASOLICITACAO").AsDateTime = ServerDate

    Set qAutorizacao = Nothing
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If Not WebMode Then
    If SessionVar("HANDLEAUTORIZACAOANEXO") <> "" Then
      NUMEROGUIAREFERENCIADA.ReadOnly = True
      BENEFICIARIO.ReadOnly = True
      NUMEROPROTOCOLORECEBIMENTO.Visible = False
    End If
  End If
End Sub

Public Sub TABLE_NewRecord()
  If Not WebMode Then

    If SessionVar("HANDLEAUTORIZACAOANEXO") <> "" Then
      CurrentQuery.FieldByName("AUTORIZACAO").AsInteger = CLng(SessionVar("HANDLEAUTORIZACAOANEXO"))
      Dim qaux As Object
      Set qaux = NewQuery
      qaux.Clear
      qaux.Add("SELECT AUTORIZACAO, BENEFICIARIO FROM SAM_AUTORIZ WHERE HANDLE = :HANDLE")
      qaux.ParamByName("HANDLE").Value = SessionVar("HANDLEAUTORIZACAOANEXO")
      qaux.Active = True

      CurrentQuery.FieldByName("NUMEROGUIAREFERENCIADA").AsInteger = qaux.FieldByName("AUTORIZACAO").AsInteger
      CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = qaux.FieldByName("BENEFICIARIO").AsInteger

      Set qaux = Nothing
    End If

  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

  If (CommandID = "BOTAOIMPRIMIRANEXO") Then
    BOTAOIMPRIMIRANEXO_OnClick
  End If

End Sub
