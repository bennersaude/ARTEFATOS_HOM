'HASH: 42FE121F6B01AC8515C40A7257A9D56A
'#Uses "*bsShowMessage"
'#Uses "*VerificaPermissaoEdicaoTriagem"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
	RecordReadOnly = Not CanContinue
End Sub
