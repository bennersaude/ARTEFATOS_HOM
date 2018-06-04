'HASH: 946DDA077352F16E274D87860256355D
'Macro: SAM_REAJUSTEMOD_CTRFAMMODPRC

'#Uses "*SAM_REAJUSTEMOD_Excluir"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 6, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub


