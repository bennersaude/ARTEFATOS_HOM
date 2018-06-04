'HASH: 7F110375804C0329454C59D1F34AAE9D
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("FAIXAINICIO").AsFloat > CurrentQuery.FieldByName("FAIXAFIM").AsFloat Then
    bsShowMessage("O início da faixa deve ser menor ou igual ao fim da faixa!", "E")
    CanContinue = False
  End If
End Sub

