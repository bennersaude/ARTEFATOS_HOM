'HASH: 4F8D2519FD0EABE18909B926AC1BDD90
'Macro: SFN_PESSOA_CONTROLESOCIETARIO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If ((Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And  (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime)) Then
		bsShowMessage("A Data Inicial não pode ser superior a Data Final.", "E")
		CanContinue = False
	End If
End Sub
