'HASH: EC371F1A962AE6CDA3399840468EDBFD

'CLI_ROTINAGRATIFICACAO


Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("É necessário gravar o registro!")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    MsgBox("A rotina não foi processada!")
    Exit Sub
  End If

  If MsgBox("Confirma o cancelamento da rotina?", vbYesNo) = vbYes Then
    Dim CANCELA As Object
    Set CANCELA = NewQuery
    CANCELA.Add("DELETE FROM CLI_ROTINAGRATIFICACAONOTA")
    CANCELA.Add("WHERE EXISTS (Select 1 FROM CLI_ROTINAGRATIFICACAOCLINICA C,")
    CANCELA.Add("                            CLI_ROTINAGRATIFICACAORECURSO R")
    CANCELA.Add("               WHERE R.ROTINAGRATIFICACAOCLINICA = C.HANDLE")
    CANCELA.Add("                 And R.HANDLE = CLI_ROTINAGRATIFICACAONOTA.GRATIFICACAORECURSO")
    CANCELA.Add("                 And C.ROTINAGRATIFICACAO = :ROTINA)")
    CANCELA.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    CANCELA.ExecSQL

    CANCELA.Clear
    CANCELA.Add("UPDATE CLI_ROTINAGRATIFICACAORECURSO Set MEDIAGERAL = Null, INDICEABSENTEISMO = Null")
    CANCELA.Add(" WHERE EXISTS (Select 1 FROM CLI_ROTINAGRATIFICACAOCLINICA C,")
    CANCELA.Add("                             CLI_ROTINAGRATIFICACAORECURSO R")
    CANCELA.Add("                WHERE C.HANDLE = R.ROTINAGRATIFICACAOCLINICA")
    CANCELA.Add("                  And R.HANDLE = CLI_ROTINAGRATIFICACAORECURSO.HANDLE")
    CANCELA.Add("                  And C.ROTINAGRATIFICACAO = :ROTINA)")
    CANCELA.ParamByName("ROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    CANCELA.ExecSQL

    CANCELA.Clear
    CANCELA.Add("UPDATE CLI_ROTINAGRATIFICACAO")
    CANCELA.Add("   SET DATAPROCESSAMENTO = NULL,")
    CANCELA.Add("       USUARIOPROCESSAMENTO = NULL,")
    CANCELA.Add("       DATACANCELAMENTO = :DATA,")
    CANCELA.Add("       USUARIOCANCELAMENTO = :USUARIO")
    CANCELA.Add(" WHERE HANDLE = :ROTINA")
    CANCELA.ParamByName("DATA").Value = ServerNow
    CANCELA.ParamByName("USUARIO").Value = CurrentUser
    CANCELA.ParamByName("ROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    CANCELA.ExecSQL
    Set CANCELA = Nothing

    RefreshNodesWithTable("CLI_ROTINAGRATIFICACAO")
  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("É necessário gravar o registro!")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    MsgBox("A rotina já foi processada!")
    Exit Sub
  End If

  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("BSCli003.Rotinas")
  AGENDA.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
  Set AGENDA = Nothing
End Sub

