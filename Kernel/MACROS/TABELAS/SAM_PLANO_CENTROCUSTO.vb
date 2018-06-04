'HASH: 5358BC78396635FB5CA40226B1B96F6E
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("LOCALFATURAMENTO").IsNull Then
    bsShowMessage("O campo 'Local de Faturamento' deve ser preenchido.", "E")
    CanContinue = False
  Else
    CanContinue = True
  End If


End Sub

