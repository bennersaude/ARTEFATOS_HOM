'HASH: 5C32DAC482E1682EEC4B9F3EC770C40B
'Macro: SAM_REAJUSTEMOD

'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	SAM_REAJUSTEMOD_Excluir 0, CurrentQuery.FieldByName("HANDLE").AsInteger
End Sub



