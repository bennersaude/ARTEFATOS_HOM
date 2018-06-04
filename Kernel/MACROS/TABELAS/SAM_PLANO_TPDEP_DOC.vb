'HASH: 210CB4B562768ED29E7952CFC2455728


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
 	If WebMode Then
 	  TIPODOCUMENTO.WebLocalWhere = " TIPODOCUMENTOBENEFICIARIO = 'S'"
 	ElseIf VisibleMode Then
	  TIPODOCUMENTO.LocalWhere = " TIPODOCUMENTOBENEFICIARIO = 'S'"
    End If
End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
 	If WebMode Then
 	  TIPODOCUMENTO.WebLocalWhere = " TIPODOCUMENTOBENEFICIARIO = 'S'"
 	ElseIf VisibleMode Then
	  TIPODOCUMENTO.LocalWhere = " TIPODOCUMENTOBENEFICIARIO = 'S'"
    End If
End Sub
