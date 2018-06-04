'HASH: 21016B61B7F1EB9269BF2FC0C0ADEB22
 

'Macro: SAM_REAJUSTEMOD_CTRMODPRC


'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 10, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub
