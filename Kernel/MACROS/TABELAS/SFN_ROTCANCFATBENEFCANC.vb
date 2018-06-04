'HASH: 601140520DD7DEB7C5147DC21092D18F
'Macro: SFN_ROTCANCFATBENEFCANC
'#Uses "*UltimoDiaCompetencia"

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <> 1 Then

    MsgBox("O registro não pode estar em edição")
    Exit Sub

  End If

  Dim vObj As Object

  Set vObj = CreateBennerObject("BSFin004.CancelamentoFatura")
  vObj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set vObj = Nothing

  RefreshNodesWithTable("SFN_ROTCANCFATBENEFCANC")
End Sub

Public Sub BOTAOVERIFICAR_OnClick()
  If CurrentQuery.State <> 1 Then

    MsgBox("O registro não pode estar em edição")
    Exit Sub

  End If

  Dim vObj As Object

  Set vObj = CreateBennerObject("BSFin004.CancelamentoFatura")
  vObj.Verificar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set vObj = Nothing

  RefreshNodesWithTable("SFN_ROTCANCFATBENEFCANC")
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime > CurrentQuery.FieldByName("VENCIMENTOFINAL").AsDateTime Then
    MsgBox("Vencimento final é inferior ao vencimento inicial!")
    CanContinue = False
    Exit Sub
  End If
  Dim vdDataContabil As Date
  vdDataContabil = UltimoDiaCompetencia(CurrentQuery.FieldByName("MESCONTABIL").AsDateTime)
  If vdDataContabil < CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime Then
    CanContinue = False
    MsgBox("Último dia do mês contábil não pode ser inferior à data de vencimento inicial")
  Else
    If CurrentQuery.FieldByName("VENCIMENTOFINAL").AsDateTime > vdDataContabil Then
      CurrentQuery.FieldByName("VENCIMENTOFINAL").AsDateTime = vdDataContabil
      If CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime > vdDataContabil Then
        CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime = vdDataContabil
      End If
    End If
  End If
End Sub

