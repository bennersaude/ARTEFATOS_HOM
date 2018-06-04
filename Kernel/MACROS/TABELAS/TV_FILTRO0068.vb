'HASH: 00501828599BD51140BC732C36E7BD5E

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull) And (Not CurrentQuery.FieldByName("PESSOA").IsNull) Then
		bsShowMessage("Escolha um Beneficiário ou uma Pessoa", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
