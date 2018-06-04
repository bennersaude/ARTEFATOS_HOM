'HASH: BBA478FDAE386D49E0211D5E3869934E

'Macro: SAM_REAJUSTEMOD_CTRFAMMODPRCFX


'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 7, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub
