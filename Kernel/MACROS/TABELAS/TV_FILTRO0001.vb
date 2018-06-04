'HASH: 141B1C92BCBBDA674AF12C639238B980
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime) Then
		bsShowMessage("A Data Inicial deve ser menor que a Data Final", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
