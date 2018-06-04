'HASH: E1623788C978D0C3CBB0202E7F1B5518
 

'Macro: SAM_REAJUSTEMOD_CTRMODADES


'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 10, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub
