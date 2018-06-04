'HASH: B57D5B5B5A70FA3DC0C323AC1C470201
 


'Macro: SAM_REAJUSTEMOD_CTRMODPRCFX


'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 11, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub
