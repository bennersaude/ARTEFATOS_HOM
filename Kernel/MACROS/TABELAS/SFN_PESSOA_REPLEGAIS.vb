'HASH: 3A01B019ADFDE4A527B6134BAB843709
'Macro: SFN_PESSOA_REPLEGAIS
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If ((Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And  (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime)) Then
		bsShowMessage("A Data Inicial não pode ser superior a Data Final.", "E")
		CanContinue = False
	End If
End Sub
