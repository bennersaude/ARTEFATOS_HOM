﻿'HASH: 7E7ECBB3A2DFBD96C9B2F9C1865FE7EA
 

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime) Then
		bsShowMessage("A Data Inicial deve ser menor que a Data Final", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
