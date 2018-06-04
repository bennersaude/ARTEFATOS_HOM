'HASH: DAC4DE3F46556782066795E9E2E7347D
 

Public Sub TABLE_AfterInsert()
'SMS 90283 - Ricardo Rocha - Adequacao WEB - 16/01/2008
	If (RecordHandleOfTable("SAM_LIVRO_ROTINAEMISSAO") > 0) Then
		CurrentQuery.FieldByName("ROTINAEMISSAO").AsInteger = RecordHandleOfTable("SAM_LIVRO_ROTINAEMISSAO")
		ROTINAEMISSAO.ReadOnly = True
	Else
		ROTINAEMISSAO.ReadOnly = False
	End If
End Sub
