'HASH: FD4F3D25F9EA541CE1106B866053D522

Public Sub TABLE_AfterInsert()
'SMS 90283 - Ricardo Rocha - Adequacao WEB - 16/01/2008
	If RecordHandleOfTable("SAM_LIVRO_ROTINAEMISSAO") > 0 Then
		CurrentQuery.FieldByName("ROTINAEMISSAO").AsInteger = RecordHandleOfTable("SAM_LIVRO_ROTINAEMISSAO")
		ROTINAEMISSAO.ReadOnly = True
	Else
		ROTINAEMISSAO.ReadOnly = False
	End If
End Sub

