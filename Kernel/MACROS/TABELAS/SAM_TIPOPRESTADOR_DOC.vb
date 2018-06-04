'HASH: 59F6F565F0C18C15597983F0CD978BB9

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		TIPODOCUMENTO.WebLocalWhere = "A.HANDLE IN (SELECT HANDLE			 " + _
									               "  FROM SAM_TIPODOCUMENTO " + _
									               " WHERE TIPODOCUMENTOPRESTADOR = 'S' )"
	ElseIf VisibleMode Then
		TIPODOCUMENTO.LocalWhere = "HANDLE IN (SELECT HANDLE 				" + _
									            "  FROM SAM_TIPODOCUMENTO 	" + _
									            " WHERE TIPODOCUMENTOPRESTADOR = 'S' )"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		TIPODOCUMENTO.WebLocalWhere = "A.HANDLE IN (SELECT HANDLE			 " + _
									               "  FROM SAM_TIPODOCUMENTO " + _
									               " WHERE TIPODOCUMENTOPRESTADOR = 'S' )"
	ElseIf VisibleMode Then
		TIPODOCUMENTO.LocalWhere = "HANDLE IN (SELECT HANDLE 				" + _
									            "  FROM SAM_TIPODOCUMENTO 	" + _
									            " WHERE TIPODOCUMENTOPRESTADOR = 'S' )"
	End If
End Sub
