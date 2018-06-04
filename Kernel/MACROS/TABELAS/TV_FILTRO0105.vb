'HASH: 23472D49E140098A6733FBCA002A0CAF

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If (Not CurrentQuery.FieldByName("PRESTADOR").IsNull) And (Not CurrentQuery.FieldByName("PESSOA").IsNull) Then
		bsShowMessage("Escolha - Prestador ou Pessoa !", "E")
		CanContinue = False
	End If

End Sub
