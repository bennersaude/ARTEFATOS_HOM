'HASH: 7DC115790966FDBA63FC259D27877B98


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		CONTRATOINDICACAO.WebLocalWhere = "PERMITEINDICAR = 'S'"
	ElseIf VisibleMode Then
		CONTRATOINDICACAO.LocalWhere = "PERMITEINDICAR = 'S'"
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		CONTRATOINDICACAO.WebLocalWhere = "PERMITEINDICAR = 'S'"
	ElseIf VisibleMode Then
		CONTRATOINDICACAO.LocalWhere = "PERMITEINDICAR = 'S'"
	End If
End Sub
