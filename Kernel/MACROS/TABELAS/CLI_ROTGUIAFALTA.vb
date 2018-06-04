'HASH: AB7AC4D386112DA65BA9784D9621E93F

'CLI_ROTGERAFALTA

Public Sub BOTAOCANCELAR_OnClick()
  If Not CurrentQuery.FieldByName("DATAGERACAO").IsNull Then
    If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
      MsgBox("Rotina já processada!")
      Exit Sub
    End If
    If MsgBox("Confirma o cancelamento da rotina ?", vbYesNo) = vbYes Then
      Dim cancela As Object
      Set cancela = NewQuery
      cancela.Clear
      cancela.Add("DELETE FROM CLI_ROTGUIAFALTAAGENDA WHERE ROTGUIAFALTA = :ROTGUIAFALTA")
      cancela.ParamByName("ROTGUIAFALTA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      cancela.ExecSQL

      cancela.Clear
      cancela.Add("UPDATE CLI_ROTGUIAFALTA SET DATACANCELAMENTO = :DATACANCELAMENTO,")
      cancela.Add("USUARIOCANCELAMENTO = :USUARIOCANCELAMENTO")
      cancela.Add("WHERE HANDLE = :HANDLE")
      cancela.ParamByName("DATACANCELAMENTO").AsDateTime = ServerNow
      cancela.ParamByName("USUARIOCANCELAMENTO").AsInteger = CurrentUser
      cancela.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("handle").AsInteger
      cancela.ExecSQL
      RefreshNodesWithTable("CLI_ROTGUIAFALTA")
      Set cancela = Nothing
    End If
  Else
    MsgBox("A rotina não foi processada!")
  End If
End Sub

Public Sub BOTAOCONFIRMARTUDO_OnClick()
  If Not InTransaction Then StartTransaction
  Dim CONFIRMA As Object
  Set CONFIRMA = NewQuery
  CONFIRMA.Clear
  CONFIRMA.Add("UPDATE CLI_ROTGUIAFALTAAGENDA SET SITUACAO = 'C', USUARIO = :USUARIO, DATA = :DATA WHERE HANDLE = :ROTGUIAFALTAAGENDA")
  Dim agenda As Object
  Set agenda = NewQuery
  agenda.Clear
  agenda.Add("SELECT * FROM CLI_ROTGUIAFALTAAGENDA WHERE ROTGUIAFALTA = :ROTGUIAFALTA AND SITUACAO = 'P'")
  agenda.ParamByName("ROTGUIAFALTA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  agenda.Active = True
  While Not agenda.EOF
    CONFIRMA.ParamByName("ROTGUIAFALTAAGENDA").AsInteger = agenda.FieldByName("HANDLE").AsInteger
    CONFIRMA.ParamByName("USUARIO").AsInteger = CurrentUser
    CONFIRMA.ParamByName("DATA").AsDateTime = ServerNow
    CONFIRMA.ExecSQL
    agenda.Next
  Wend
  RefreshNodesWithTable("CLI_ROTGUIAFALTAAGENDA")
  Set agenda = Nothing
  Set CONFIRMA = Nothing
  If InTransaction Then Commit
End Sub

Public Sub BOTAOGERAR_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("O registro está em edição! Por favor confirme ou cancele as alterações")
    Exit Sub
  End If
  If Not CurrentQuery.FieldByName("DATAGERACAO").IsNull And CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    MsgBox("Rotina já gerada!")
    Exit Sub
  End If
  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    MsgBox("Rotina já processada!")
    Exit Sub
  End If
  Dim interface As Object
  Set interface = CreateBennerObject("CliGeraGuia.GeraGuia")
  interface.Falta(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger, _
                  CurrentQuery.FieldByName("clinica").AsInteger)
  RefreshNodesWithTable("CLI_ROTGUIAFALTA")
  Set interface = Nothing
End Sub


Public Sub BOTAOPROCESSAR_OnClick()
  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    MsgBox("Rotina já processada!")
    Exit Sub
  End If
  Dim VERIFICA As Object
  Set VERIFICA = NewQuery
  VERIFICA.Clear
  VERIFICA.Add("SELECT COUNT(HANDLE) TOTAL FROM CLI_ROTGUIAFALTAAGENDA WHERE ROTGUIAFALTA = :ROTGUIAFALTA AND SITUACAO = 'P'")
  VERIFICA.ParamByName("ROTGUIAFALTA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  VERIFICA.Active = True
  If VERIFICA.FieldByName("TOTAL").AsInteger = 0 Then
    Dim interface As Object
    Set interface = CreateBennerObject("CliGeraGuia.GeraGuia")
    interface.GuiaFalta(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger)
    RefreshNodesWithTable("CLI_ROTGUIAFALTA")
    Set interface = Nothing
  Else
    MsgBox("Ainda existem registros pendentes!")
  End If
  Set VERIFICA = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qBuscaRotina As Object
  Set qBuscaRotina = NewQuery
  qBuscaRotina.Clear
  qBuscaRotina.Add("SELECT NUMERO FROM CLI_ROTGUIAFALTA  ")
  qBuscaRotina.Add(" WHERE DATAPROCESSAMENTO IS NULL     ")
  qBuscaRotina.Add("   AND CLINICA = :CLINICA            ")
  qBuscaRotina.Add("   AND COMPETENCIA = :COMPETENCIA    ")
  qBuscaRotina.Add("   AND HANDLE <> :HANDLE             ")
  qBuscaRotina.ParamByName("CLINICA").AsInteger = CurrentQuery.FieldByName("CLINICA").AsInteger
  qBuscaRotina.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
  qBuscaRotina.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qBuscaRotina.Active = True

  If Not qBuscaRotina.EOF Then
    MsgBox("A rotina " + qBuscaRotina.FieldByName("NUMERO").AsString + " desta mesma competência e clínica não foi processada!")
    CanContinue = False
  End If

  Set qBuscaRotina = Nothing
End Sub
