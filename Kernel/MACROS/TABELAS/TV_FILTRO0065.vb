'HASH: 6374D218F2576295C0C65B78A9757490
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("NUMEROINICIAL").AsInteger < 0) Or _
	   (CurrentQuery.FieldByName("NUMEROINICIAL").AsInteger > CurrentQuery.FieldByName("NUMEROFINAL").AsInteger) Then
		bsShowMessage("Conta financeira inicial inválida", "E")
		CanContinue = False
		Exit Sub
	ElseIf CurrentQuery.FieldByName("NUMEROFINAL").AsInteger < 0 Then
		bsShowMessage("Conta financeira final inválida", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
