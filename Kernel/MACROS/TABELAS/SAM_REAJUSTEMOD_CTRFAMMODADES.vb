'HASH: 5779296AD72BF2F1A2DA56F5626B8B41
'Macro: SAM_REAJUSTEMOD_CTRFAMMODADES

'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 5, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub
