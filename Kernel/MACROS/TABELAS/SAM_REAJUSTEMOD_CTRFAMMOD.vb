'HASH: C4CD5A466266729C9A30BA845F1E219D
'Macro: SAM_REAJUSTEMOD_CTRFAMMOD

'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 4, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub
