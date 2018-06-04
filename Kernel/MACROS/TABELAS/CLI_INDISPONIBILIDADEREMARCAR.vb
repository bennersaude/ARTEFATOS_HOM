'HASH: B1B8F27C5DCDB4C83981BF978123B257
'CLI_INDISPONIBILIDADEREMARCAR


Public Sub BOTAODESMARCAR_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("É necessário gravar o registro!")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    If CurrentQuery.FieldByName("MOTIVODESMARQUE").IsNull Then
      MsgBox("Não foi informado o motivo de desmarque!")
    Else
      Dim BSCli001dll As Object
      Set BSCli001dll = CreateBennerObject("BSCli001.Rotinas")
      BSCli001dll.DesmarcaAgendaIndisponibilidade(CurrentSystem, _
                                                  CurrentQuery.FieldByName("AGENDA").AsInteger, _
                                                  CurrentQuery.FieldByName("MOTIVODESMARQUE").AsInteger, _
                                                  CurrentQuery.FieldByName("HANDLE").AsInteger)
      RefreshNodesWithTable("CLI_INDISPONIBILIDADEREMARCAR")
      Set BSCli001dll = Nothing
    End If
  Else
    MsgBox("Para desmarcar a situação deve estar pendente!")
  End If
End Sub

Public Sub BOTAOREMARCAR_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("É necessário gravar o registro!")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    Dim BSCli001dll As Object
    Set BSCli001dll = CreateBennerObject("BSCli001.Rotinas")
    BSCli001dll.RemarcaAgendaIndisponibilidade(CurrentSystem, _
                                               CurrentQuery.FieldByName("AGENDA").AsInteger, _
                                               CurrentQuery.FieldByName("HANDLE").AsInteger)
    RefreshNodesWithTable("CLI_INDISPONIBILIDADEREMARCAR")
    Set BSCli001dll = Nothing
  Else
    MsgBox("Para remarcar a situação deve estar pendente!")
  End If
End Sub

