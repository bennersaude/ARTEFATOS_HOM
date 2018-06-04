'HASH: 5B208D09F03CABF79E2CFF110851267A
'Macro: SAM_CALENDARIOREEMBOLSO
'#Uses "*bsShowMessage"

Public Sub BOTAOABRIR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Registro esta em edição.", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull Then
    bsShowMessage("Calendário Aberto.", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    bsShowMessage("Calendário Já Processado.", "I")
    Exit Sub
  End If

  If VisibleMode Then
  	If MsgBox("Confirma abertura do calendário?.", vbYesNo, "Calendário de Reembolso") = vbNo Then
	    Exit Sub
  	End If
  End If

  Dim Qu As Object
  Set Qu = NewQuery

  If Not InTransaction Then StartTransaction

  Qu.Add("UPDATE SAM_CALENDARIOREEMBOLSO SET DATAFECHAMENTO = :DATAFECHAMENTO, USUARIOFECHAMENTO = :USUARIOFECHAMENTO WHERE HANDLE = :HANDLE")
  Qu.ParamByName("DATAFECHAMENTO").DataType = ftDateTime
  Qu.ParamByName("DATAFECHAMENTO").Clear
  Qu.ParamByName("USUARIOFECHAMENTO").DataType = ftInteger
  Qu.ParamByName("USUARIOFECHAMENTO").Clear
  Qu.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Qu.ExecSQL

  If InTransaction Then Commit

  CurrentQuery.Active = False
  CurrentQuery.Active = True

End Sub

Public Sub BOTAOFECHAR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Registro esta em edição.", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull Then
    bsShowMessage("Calendário já Fechado.", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    bsShowMessage("Calendário já Processado.", "I")
    Exit Sub
  End If

  'If   CurrentQuery.FieldByName("DATAFINAL").AsDateTime <ServerDate Then
  '  If MsgBox("Data final menor que data do fechamento - Continuar?.",vbYesNo,"Calendário de Reembolso")=vbNo Then
  '    Exit Sub
  '  End If
  'End If

  'If   CurrentQuery.FieldByName("DATAFINAL").AsDateTime > ServerDate Then
  '  If MsgBox("Data final maior que data do fechamento - Continuar?.",vbYesNo,"Calendário de Reembolso")=vbNo Then
  '    Exit Sub
  '  End If
  'End If

  Dim Qu As Object
  Set Qu = NewQuery

  If Not InTransaction Then StartTransaction

  Qu.Add("UPDATE SAM_CALENDARIOREEMBOLSO SET DATAFECHAMENTO = :DATAFECHAMENTO, USUARIOFECHAMENTO = :USUARIOFECHAMENTO WHERE HANDLE = :HANDLE")
  Qu.ParamByName("DATAFECHAMENTO").AsDateTime = ServerNow
  Qu.ParamByName("USUARIOFECHAMENTO").AsInteger = CurrentUser
  Qu.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  Qu.ExecSQL

  If InTransaction Then Commit

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub TABLE_AfterInsert()
  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT DATAFINAL FROM SAM_CALENDARIOREEMBOLSO ORDER BY DATAFINAL DESC")
  Q.Active = True
  If Not Q.FieldByName("DATAFINAL").IsNull Then
    CurrentQuery.FieldByName("DATAINICIAL").Value = (Q.FieldByName("DATAFINAL").AsDateTime + 1)
    DATAINICIAL.ReadOnly = True
    DATAFINAL.SetFocus
  Else
    DATAINICIAL.ReadOnly = False
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("TABDATAPAGAMENTO").AsString = "1" Then
    BOTAOFECHAR.Enabled = True
    BOTAOABRIR.Enabled = True
  Else
    BOTAOFECHAR.Enabled = False
    BOTAOABRIR.Enabled = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  If CurrentQuery.FieldByName("TABDATAPAGAMENTO").AsString = "1" Then
    Set Interface = CreateBennerObject("samcalendariopgto.ROTINAS")
    Interface.INICIALIZAR(CurrentSystem)
    If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime <>Interface.DIAUTILANTERIOR(CurrentSystem, CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)Then
      bsShowMessage("Entre com um dia útil para a Data de Pagamento", "E")
      DATAPAGAMENTO.SetFocus
      Interface.FINALIZAR
      Set Interface = Nothing
      CanContinue = False
      Exit Sub
    End If

    Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
    Linha = Interface.Vigencia(CurrentSystem, "SAM_CALENDARIOREEMBOLSO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "", "")

    If Linha = "" Then
      CanContinue = True
      If CLng(CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)<= CLng(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)Then
        bsShowMessage("Data de pagamento não pode ser menor ou igual a data final","E")
        CanContinue = False
      End If
    Else
      CanContinue = False
      bsShowMessage(Linha, "E")
    End If
    Set Interface = Nothing

  Else
    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
       bsShowMessage("A data final não pode ser menor que a data inicial","E")
       CanContinue = False
    End If
  	CurrentQuery.FieldByName("DATAPAGAMENTO").Clear
  	CurrentQuery.FieldByName("NUMEROPAGAMENTO").Clear

  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOABRIR"
			BOTAOABRIR_OnClick
		Case "BOTAOFECHAR"
			BOTAOFECHAR_OnClick
	End Select
End Sub
