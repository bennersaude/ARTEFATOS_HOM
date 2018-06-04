'HASH: 11C02D235A2F268B211F0052B586AC65
 

'Macro: SAM_REAJUSTEMOD_CTRMOD


'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 9, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub
