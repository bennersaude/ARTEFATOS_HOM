'HASH: D7BD004FD98C1A655BA1AF9FF7F7B839
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("CHECKGERAL1").AsString = "N" Or CurrentQuery.FieldByName("CHECKGERAL1").AsString = Null) _
	 And (CurrentQuery.FieldByName("CHECKGERAL2").AsString = "N" Or CurrentQuery.FieldByName("CHECKGERAL2").AsString = Null) _
	 And(CurrentQuery.FieldByName("CHECKGERAL3").AsString = "N" Or CurrentQuery.FieldByName("CHECKGERAL3").AsString = Null) Then
		bsShowMessage("Deve ser informado o Tipo Responsável!", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
