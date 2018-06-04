'HASH: 159826A1C494DED995621794DEE8004E
'Macro: SAM_ROTINACONVENIORECIP

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <> 1 Then
    If VisibleMode Then
      MsgBox("O registro não pode estar em edição!")
    Else
      CancelDescription = "O registro não pode estar em edição!"
    End If
  Else
    Dim vsRetornoRotina As String

    If CurrentQuery.FieldByName("TABTIPO").AsInteger = 1 Then 'Rotina de renovação
      Dim dllBSBen017_RenovacaoConvenioReciprocidade As Object
      Set dllBSBen017_RenovacaoConvenioReciprocidade = CreateBennerObject("BSBen017.RenovacaoConvenioReciprocidade")

      vsRetornoRotina= dllBSBen017_RenovacaoConvenioReciprocidade.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

      Set dllBSBen017_RenovacaoConvenioReciprocidade = Nothing

      If vsRetornoRotina <> "" Then
        If VisibleMode Then
          MsgBox(vsRetornoRotina)
        Else
          InfoDescription = vsRetornoRotina
        End If
      End If
    Else 'Rotina de cancelamento
      Dim dllBSBen017_CancelamentoConvenioReciprocidade As Object
      Set dllBSBen017_CancelamentoConvenioReciprocidade = CreateBennerObject("BSBen017.CancelamentoConvenioReciprocidade")

      vsRetornoRotina= dllBSBen017_CancelamentoConvenioReciprocidade.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

      Set dllBSBen017_CancelamentoConvenioReciprocidade = Nothing

      If vsRetornoRotina <> "" Then
        If VisibleMode Then
          MsgBox(vsRetornoRotina)
        Else
          InfoDescription = vsRetornoRotina
        End If
      End If
    End If
  End If
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.State <> 1 Then
    If VisibleMode Then
      MsgBox("O registro não pode estar em edição!")
    Else
      CancelDescription = "O registro não pode estar em edição!"
    End If
  Else
    Dim vsRetornoRotina As String

    If CurrentQuery.FieldByName("TABTIPO").AsInteger = 1 Then 'Rotina de renovação
      Dim dllBSBen017_RenovacaoConvenioReciprocidade As Object
      Set dllBSBen017_RenovacaoConvenioReciprocidade = CreateBennerObject("BSBen017.RenovacaoConvenioReciprocidade")

      vsRetornoRotina= dllBSBen017_RenovacaoConvenioReciprocidade.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

      Set dllBSBen017_RenovacaoConvenioReciprocidade = Nothing

      If vsRetornoRotina <> "" Then
        If VisibleMode Then
          MsgBox(vsRetornoRotina)
        Else
          InfoDescription = vsRetornoRotina
        End If
      End If
    Else 'Rotina de cancelamento
      Dim dllBSBen017_CancelamentoConvenioReciprocidade As Object
      Set dllBSBen017_CancelamentoConvenioReciprocidade = CreateBennerObject("BSBen017.CancelamentoConvenioReciprocidade")

      vsRetornoRotina= dllBSBen017_CancelamentoConvenioReciprocidade.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

      Set dllBSBen017_CancelamentoConvenioReciprocidade = Nothing

      If vsRetornoRotina <> "" Then
        If VisibleMode Then
          MsgBox(vsRetornoRotina)
        Else
          InfoDescription = vsRetornoRotina
        End If
      End If
    End If
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    CanContinue = False
    If VisibleMode Then
      MsgBox("A rotina já está processada. Alteração não permitida!")
    Else
      CancelDescription = "A rotina já está processada. Alteração não permitida!"
    End If
  End If
End Sub

Public Sub TABLE_NewRecord()
  If NodeInternalCode = 1 Then 'Código interno da carga da rotina de renovação
    CurrentQuery.FieldByName("TABTIPO").AsInteger = 1
  Else
    CurrentQuery.FieldByName("TABTIPO").AsInteger = 2
  End If
End Sub
