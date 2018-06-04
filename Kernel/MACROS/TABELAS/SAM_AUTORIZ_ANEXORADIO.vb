'HASH: C860AC972F31E4D10B1BAA9A41E558FF
'#uses "*bsShowMessage"'
Option Explicit

Public Sub BOTAOIMPRIMIRANEXO_OnClick()
    Dim qAux As BPesquisa
    Dim vHandleRelatorio As Integer

	Set qAux = NewQuery
	qAux.Clear
	qAux.Add("SELECT HANDLE			    ")
	qAux.Add("  FROM R_RELATORIOS       ")
	qAux.Add(" WHERE CODIGO = 'TISS013' ")
    qAux.Active = True
	qAux.First

    vHandleRelatorio = qAux.FieldByName("HANDLE").AsInteger
    Set qAux = Nothing

	If vHandleRelatorio > 0 Then
        Dim qProcedimento As BPesquisa
        Set qProcedimento = NewQuery

  		qProcedimento.Add("SELECT 1 FROM SAM_AUTORIZ_ANEXORADIO_PROC WHERE ANEXORADIOTERAPIA = :ANEXORADIO")
  		qProcedimento.ParamByName("ANEXORADIO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  		qProcedimento.Active = True

  		If qProcedimento.EOF Then
		  bsShowMessage("O anexo não pode ser impresso pois não possui procedimentos cadastrados!", "I")
  		  Set qProcedimento = Nothing
		  Exit Sub
  		Else
		  Dim rep As CSReportPrinter
		  Set rep = NewReport(vHandleRelatorio)
		  SessionVar("ANEXORADIO") = CurrentQuery.FieldByName("HANDLE").AsString
		  rep.Preview
		  Set rep = Nothing
		End If
		Set qProcedimento = Nothing
	Else
  	   bsShowMessage("Nenhum relatório para impressão de anexo de solicitação de radioterapia foi encontrado.", "E")
  	End If
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
	Select Case CommandID
		Case  "BOTAOIMPRIMIRANEXO"
			BOTAOIMPRIMIRANEXO_OnClick
	End Select
End Sub
