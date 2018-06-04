'HASH: A3A8BDA63F27B94E7A3CA0D95E8A344C


Public Sub BOTAOCALCULARVALORCOTA_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If



  Set Obj = CreateBennerObject("SAMFaturamento.Rateio")
  Obj.CalcularValorCota(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  CurrentQuery.Active = False
  CurrentQuery.Active = True
  Set Obj = Nothing

  WriteAudit("C", HandleOfTable("SFN_COMPETFIN_GRUPO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Custo Operacional (Rateio) - Cálculo do Valor da Cota")
End Sub

Public Sub BOTAOTOTALIZARGASTOS_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If



  Set Obj = CreateBennerObject("SAMFaturamento.Rateio")
  Obj.TotalizarGastos(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  CurrentQuery.Active = False
  CurrentQuery.Active = True
  Set Obj = Nothing

  WriteAudit("T", HandleOfTable("SFN_COMPETFIN_GRUPO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Custo Operacional (Rateio) - Totalização dos Gastos")
End Sub

