'HASH: 3B5F2402DD28780FE3EE8C56B38187C8

'Macro: SFN_ROTINASALDOGERENCIALREC

Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.State <> 1 Then

    MsgBox("O registro não pode estar em edição")
    Exit Sub

  End If

  Dim vObj As Object

  Set vObj = CreateBennerObject("BSFin003.SaldoGerencialRecebimento")
  vObj.Cancelar(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentSystem)
  Set vObj = Nothing

  RefreshNodesWithTable("SFN_ROTINASALDOGERENCIALREC")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()

  If CurrentQuery.State <> 1 Then

    MsgBox("O registro não pode estar em edição")
    Exit Sub

  End If

  Dim vObj As Object

  Set vObj = CreateBennerObject("BSFin003.SaldoGerencialRecebimento")
  vObj.Processar(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentSystem)
  Set vObj = Nothing

  RefreshNodesWithTable("SFN_ROTINASALDOGERENCIALREC")
End Sub

