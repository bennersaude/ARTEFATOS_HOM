'HASH: B61E81A056FB48C0C2D573E67C6CFA96
'Macro: SAM_REAJUSTEMOD_CTR

'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 2, CurrentQuery.FieldByName("HANDLE").AsInteger
	SAM_REAJUSTEMOD_Excluir 8, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub


