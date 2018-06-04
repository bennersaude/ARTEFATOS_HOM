'HASH: A714947605F099B9040BD04FA135ABDB


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  CurrentQuery.UpdateRecord

  If (CurrentQuery.FieldByName("COBRACORRECAO").AsString = "S" And CurrentQuery.FieldByName("SUSPENDEFATURAMENTO").AsString = "N") Then
    MsgBox("Não é possível Cobrar Correção se o faturamento não estiver suspenso, desmarque a opção Cobrar Correção!")
    CanContinue = False
  End If


End Sub

