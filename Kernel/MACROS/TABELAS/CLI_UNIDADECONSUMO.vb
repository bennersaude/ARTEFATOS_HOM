'HASH: 964D397B5AF8B9ECA4FD6BEE83C7D925
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("TIPO").AsString <> "4" And _
     CurrentQuery.FieldByName("TIPO").AsString <> "5" And _
     Not CurrentQuery.FieldByName("RELACAOIMC").IsNull Then

     CanContinue = False
     BSShowMessage("O campo 'Relação unidade/cálculo IMC' só deve ser informado para peso e altura!","E")

  End If

End Sub
