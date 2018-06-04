'HASH: 42FD476273CDEABEF1DFC6258C964554
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull) And (Not CurrentQuery.FieldByName("PESSOA").IsNull) And (Not CurrentQuery.FieldByName("PRESTADOR").IsNull) Then
		bsShowMessage("Selecione somente um responsável da conta financeira", "E")
		CanContinue = False
		Exit Sub
	ElseIf (Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull) And (Not CurrentQuery.FieldByName("PESSOA").IsNull) Then
		bsShowMessage("Selecione somente um responsável da conta financeira", "E")
		CanContinue = False
		Exit Sub
	ElseIf (Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull) And (Not CurrentQuery.FieldByName("PRESTADOR").IsNull) Then
		bsShowMessage("Selecione somente um responsável da conta financeira", "E")
		CanPrint = False
		Exit Sub
	ElseIf (Not CurrentQuery.FieldByName("PESSOA").IsNull) And (Not CurrentQuery.FieldByName("PRESTADOR").IsNull) Then
		bsShowMessage("Selecione somente um responsável da conta financeira", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
