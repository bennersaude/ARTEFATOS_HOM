'HASH: F8157D9D87AA67F564C5390915D05418
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim qLancQuebra As Object

  Set qLancQuebra = NewQuery

  qLancQuebra.Add("SELECT COUNT(*) QTDREGISTROS")
  qLancQuebra.Add("FROM SFN_LANCAMENTOEXTRATO_QUEBRA")
  qLancQuebra.Add("WHERE LANCAMENTOEXTRATO =:HLANCAMENTOEXTRATO")
  qLancQuebra.ParamByName("HLANCAMENTOEXTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qLancQuebra.Active = True

  If qLancQuebra.FieldByName("QTDREGISTROS").AsInteger > 0 Then
    CanContinue = False
    bsShowMessage("Lançamento possui registros de quebra! Exclua primeiros os lançamentos de quebra.", "E")
  End If

  Set qLancQuebra = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("LANCAMENTOCONCILIADO").AsString = "S" Then
    If CurrentQuery.FieldByName("CATEGORIA").IsNull Then
      CanContinue = False
      bsShowMessage("É obrigatório preenchimento do campo ""Categoria""!", "E")
    End If

    Dim qLancQuebra As Object

    Set qLancQuebra = NewQuery

    qLancQuebra.Add("SELECT COUNT(HANDLE) QTDREGISTROS,")
    qLancQuebra.Add("       Abs(SUM(Case WHEN NATUREZA = 'D' THEN -VALOR ELSE VALOR END)) VALOR,")
    qLancQuebra.Add("       CASE WHEN (SUM(CASE WHEN NATUREZA = 'D' THEN -VALOR ELSE VALOR END)) < 0 THEN 'D' ELSE 'C' END NATUREZA")
    qLancQuebra.Add("FROM SFN_LANCAMENTOEXTRATO_QUEBRA")
    qLancQuebra.Add("WHERE LANCAMENTOEXTRATO =:HLANCAMENTOEXTRATO")
    qLancQuebra.ParamByName("HLANCAMENTOEXTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qLancQuebra.Active = True

    If qLancQuebra.FieldByName("QTDREGISTROS").AsInteger > 0 And _
       (CurrentQuery.FieldByName("NATUREZA").AsString <> qLancQuebra.FieldByName("NATUREZA").AsString Or _
        CurrentQuery.FieldByName("VALOR").AsFloat     <> qLancQuebra.FieldByName("VALOR").AsFloat) Then
      CanContinue = False
      bsShowMessage("Soma dos lançamentos de quebra está diferente do valor deste lançamento! Não será permitido marcar este lançamento como ""Conciliado"".", "E")
    End If

    Set qLancQuebra = Nothing
  End If
End Sub
