'HASH: 637BE96EEF8A000A49278CC0CC7F2701
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If LancamentoConciliado Then
    CanContinue = False
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If LancamentoConciliado Then
    CanContinue = False
  End If

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If LancamentoConciliado Then
    CanContinue = False
  End If

  If CurrentQuery.FieldByName("VALOR").AsFloat < 0 Then
    CanContinue = False
    bsShowMessage("Valor mínimo para o campo ""Valor"" é 0,01!", "E")
  End If
End Sub


' checar se o lançamento está conciliado, caso esteja barrar exibindo mensagem para que o lançamento seja desconsiliado liberando assim alteração/exclusão/inclusão de ítens
Public Function LancamentoConciliado As Boolean
	LancamentoConciliado = False

	Dim qLancConciliado As Object
	Set qLancConciliado = NewQuery
	qLancConciliado.Clear
	qLancConciliado.Add("SELECT LANCAMENTOCONCILIADO FROM SFN_LANCAMENTOEXTRATO WHERE HANDLE = :LANCAMENTOEXTRATO")
	qLancConciliado.ParamByName("LANCAMENTOEXTRATO").AsInteger = CurrentQuery.FieldByName("LANCAMENTOEXTRATO").AsInteger
	qLancConciliado.Active = True
	If qLancConciliado.FieldByName("LANCAMENTOCONCILIADO").AsString = "S" Then
	  LancamentoConciliado = True
	      bsShowMessage("Lançamento do extrato já está conciliado! " + Chr(10) + Chr(13) + "Para realizar qualquer modificação deve ser desmarcado a conciliação no lançamento do extrato", "E")
    End If
	Set qLancConciliado = Nothing
End Function

