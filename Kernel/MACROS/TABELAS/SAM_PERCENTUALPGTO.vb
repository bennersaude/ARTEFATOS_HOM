'HASH: 371BE8EB98AAE6E2191ACC005D8DF237
'Macro: SAM_PERCENTUALPGTO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("PERCENTUALPGTODEMAIS").AsInteger >0 Then
		If CurrentQuery.FieldByName("INCIDENCIAMINIMA").AsInteger < 2 Then
			bsShowMessage("Incidência Mínima deve ser maior que (1) !", "E")
			INCIDENCIAMINIMA.SetFocus
			CanContinue = False
		End If
	End If
End Sub


