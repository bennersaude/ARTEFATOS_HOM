'HASH: DFAC538ADBE11938742A4A841846F28C
'Macro: SAM_CALENDARIOGERAL
' Mauricio Ibelli -sms1198 -incluido data de fechamento no calendario

Public Sub BOTAOABRIR_OnClick()

  If CurrentQuery.State <>1 Then
    MsgBox("Registro esta em edição.")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull Then
    MsgBox("Calendário Aberto.")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    MsgBox("Calendário Já Processado.")
    Exit Sub
  End If

  If MsgBox("Confirma abertura do calendário?.", vbYesNo, "Calendário de Pagamento") = vbNo Then
    Exit Sub
  End If

  Dim Qu As Object
  Set Qu = NewQuery
  Qu.Add("UPDATE SAM_CALENDARIOGERAL SET DATAFECHAMENTO = :DATAFECHAMENTO, USUARIOFECHAMENTO = :USUARIOFECHAMENTO WHERE HANDLE = :HANDLE")
  Qu.ParamByName("DATAFECHAMENTO").DataType = ftDateTime
  Qu.ParamByName("DATAFECHAMENTO").Clear
  Qu.ParamByName("USUARIOFECHAMENTO").DataType = ftInteger
  Qu.ParamByName("USUARIOFECHAMENTO").Clear
  Qu.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Qu.ExecSQL

  CurrentQuery.Active = False
  CurrentQuery.Active = True

End Sub

Public Sub BOTAOFECHAR_OnClick()

  If CurrentQuery.State <>1 Then
    MsgBox("Registro esta em edição.")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAFECHAMENTO").IsNull Then
    MsgBox("Calendário já Fechado.")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    MsgBox("Calendário já Processado.")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAFINAL").AsDateTime <ServerDate Then
    If MsgBox("Data final menor que data do fechamento - Continuar?.", vbYesNo, "Calendário de Pagamento") = vbNo Then
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("DATAFINAL").AsDateTime >ServerDate Then
    If MsgBox("Data final maior que data do fechamento - Continuar?.", vbYesNo, "Calendário de Pagamento") = vbNo Then
      Exit Sub
    End If
  End If

  Dim Qu As Object
  Set Qu = NewQuery
  Qu.Add("UPDATE SAM_CALENDARIOGERAL SET DATAFECHAMENTO = :DATAFECHAMENTO, USUARIOFECHAMENTO = :USUARIOFECHAMENTO WHERE HANDLE = :HANDLE")
  Qu.ParamByName("DATAFECHAMENTO").AsDateTime = ServerNow
  Qu.ParamByName("USUARIOFECHAMENTO").AsInteger = CurrentUser
  Qu.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  Qu.ExecSQL

  CurrentQuery.Active = False
  CurrentQuery.Active = True

End Sub

Public Sub TABLE_AfterInsert()

  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT DATAFINAL FROM SAM_CALENDARIOGERAL ORDER BY DATAFINAL DESC")
  Q.Active = True
  If Not Q.FieldByName("DATAFINAL").IsNull Then
    CurrentQuery.FieldByName("DATAINICIAL").Value = (Q.FieldByName("DATAFINAL").AsDateTime + 1)
    DATAINICIAL.ReadOnly = True
    DATAFINAL.SetFocus
  Else
    DATAINICIAL.ReadOnly = False
  End If

End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Exit Sub

  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT DATAFINAL FROM SAM_CALENDARIOGERAL WHERE DATAFECHAMENTO IS NULL")
  Q.Active = True
  If Not Q.EOF Then
    MsgBox("Calendário com processo em Aberto - Inclusão não permitida.")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Set INTERFACE = CreateBennerObject("samcalendariopgto.ROTINAS")
  INTERFACE.INICIALIZAR(CurrentSystem)
  If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime <>INTERFACE.DIAUTILANTERIOR(CurrentSystem, CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)Then
    MsgBox("Entre com um dia útil para a Data de Pagamento")
    DATAPAGAMENTO.SetFocus
    INTERFACE.FINALIZAR
    Set INTERFACE = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Linha = Interface.Vigencia(CurrentSystem, "SAM_CALENDARIOGERAL", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "", "")

  If Linha = "" Then
    CanContinue = True

    If CLng(CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)<= CLng(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)Then
      MsgBox "Data de pagamento nao pode ser menor ou igual a data final"
      Cancontinue = False
    End If
  Else
    CanContinue = False
    MsgBox(Linha)
  End If
  Set Interface = Nothing
End Sub

