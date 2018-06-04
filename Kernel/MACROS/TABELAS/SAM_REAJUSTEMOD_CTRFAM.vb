'HASH: 52704D91BBDC1A9B540D5F52286D069A
'Macro: SAM_REAJUSTEMOD_CTRFAM

'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 3, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub
