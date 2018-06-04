'HASH: 3FB534D105FF7B017CB06B5C42294BB4
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime) Then
		bsShowMessage("A Data Inicial deve ser menor que a Data Final", "E")
		CanContinue = False
		Exit Sub
	End If

	If (CurrentQuery.FieldByName("CHECKGERAL1").AsString = "N") And (CurrentQuery.FieldByName("CHECKGERAL2").AsString = "N") Then
		bsShowMessage("Escolha uma opção: Física/Jurídica", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
