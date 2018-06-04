'HASH: 6E2CD59FF2D668F8CB24B96A7D470E14
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim SQLTipoFat As Object

Set SQLTipoFat = NewQuery
	SQLTipoFat.Clear
	SQLTipoFat.Add("SELECT HANDLE, CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CStr(CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger))
	SQLTipoFat.Active = True

	If SQLTipoFat.FieldByName("CODIGO").AsInteger <> 660 Then
		bsShowMessage("O Tipo de Faturamento deve ser recolhimento de Contribuições Federais!", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
