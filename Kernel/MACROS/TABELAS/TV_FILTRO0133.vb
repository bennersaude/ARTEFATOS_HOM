'HASH: 64850BB862B70792AEB242FFBA9758DB
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime) Then
		bsShowMessage("A Data Inicial deve ser menor que a Data Final", "E")
		CanContinue = False
		Exit Sub
	End If

	If (CurrentQuery.FieldByName("CHECKGERAL1").AsString = "N") And (CurrentQuery.FieldByName("CHECKGERAL2").AsString = "N") _
	  And (CurrentQuery.FieldByName("CHECKGERAL3").AsString = "N") Then
		bsShowMessage("Escolha uma opção: Física/Jurídica/Cooperativa", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
