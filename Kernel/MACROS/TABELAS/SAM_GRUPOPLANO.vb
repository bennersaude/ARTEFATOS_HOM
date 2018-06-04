'HASH: E3BA16D52CA0A94B381F0DA118E902E7

'#Uses "*bsShowMessage"

'SMS - 91312
Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim CODIGO As Object
	Set CODIGO = NewQuery

	CODIGO.Add("SELECT HANDLE")
	CODIGO.Add("FROM SAM_GRUPOPLANO")
	CODIGO.Add("WHERE CODIGO = :CODIGO")
	CODIGO.Add("AND HANDLE <> :HANDLE")
	CODIGO.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
	CODIGO.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	CODIGO.Active = True

	If CODIGO.FieldByName("HANDLE").AsInteger > 0 Then
		bsShowMessage("Código já cadastrado!", "E")
		Set CODIGO = Nothing
		CanContinue = False
		Exit Sub
	End If

	Set CODIGO = Nothing
End Sub
