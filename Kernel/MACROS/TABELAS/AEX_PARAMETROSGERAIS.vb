'HASH: 3B661353032D2C5BBC61AA1625D7F6D8
' atualizada em 10/08/2007
'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)

If (CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString = "R" Or CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString = "O")  And _
    CurrentQuery.FieldByName("AUTORIZACAOEXECUTOR").AsString = "N"    And _
    CurrentQuery.FieldByName("AUTORIZACAOSOLICITANTE").AsString = "N" And _
    CurrentQuery.FieldByName("AUTORIZACAORECEBEDOR").AsString = "N"   And _
    CurrentQuery.FieldByName("AUTORIZACAOLOCALEXEC").AsString = "N"   Then

   bsShowMessage("Ação autorização exige pelo menos um prestador!", "E")
   CanContinue = False
   Exit Sub

End If

If CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString = "R" And _
   CurrentQuery.FieldByName("MOTIVONEGACAO").IsNull  Then

   bsShowMessage("Ação autorização de restriçao exige um motivo de negação!","E")
   CanContinue = False
   Exit Sub

End If

If  CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString <> "N"       And _
    CurrentQuery.FieldByName("PAGAMENTOEXECUTOR").AsString = "N"    And _
    CurrentQuery.FieldByName("PAGAMENTORECEBEDOR").AsString = "N" And _
    CurrentQuery.FieldByName("PAGAMENTOLOCALEXEC").AsString = "N"   Then

   bsShowMessage("Ação pagamento exige pelo menos um prestador!", "E")
   CanContinue = False
   Exit Sub

End If

If CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString = "R" And _
   CurrentQuery.FieldByName("MOTIVOGLOSA").IsNull  Then

   bsShowMessage("Ação pagamento de restriçao exige um motivo de glosa!", "E")
   CanContinue = False
   Exit Sub

End If

End Sub
