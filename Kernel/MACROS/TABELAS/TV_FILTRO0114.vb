'HASH: 609867E6302D574E6A104C81A54661E8
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If(CurrentQuery.FieldByName("PESSOA").AsInteger >0)And(CurrentQuery.FieldByName("BENEFICIARIO").AsInteger >0)Then
 		bsShowMessage("Não deve ser selecionado uma PESSOA e um Beneficiário ao mesmo tempo!", "E")
		CanContinue =False
	End If

End Sub
