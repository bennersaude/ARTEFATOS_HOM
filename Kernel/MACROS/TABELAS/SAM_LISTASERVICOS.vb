'HASH: D21A768F9605D55B9D7B99B78A5F262B
'MACRO = SAM_LISTASERVICOS
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim QueryCodigo As Object
	Dim vsCodigo As String
	Dim viCaracter As Integer
	Dim vsCaracter As Variant

	vsCodigo = CurrentQuery.FieldByName("CODIGO").AsString

	For viCaracter = 1 To Len(vsCodigo)
		vsCaracter = Mid(vsCodigo, viCaracter, 1)
		If vsCaracter <> Chr(46) Then
			If Not (IsNumeric(vsCaracter)) Then
				bsShowMessage("Campo código não foi preenchido corretamente. Verifique!", "E")
				CanContinue = False
				Exit Sub
			End If
		End If
	Next

	Set QueryCodigo = NewQuery

	QueryCodigo.Clear
	QueryCodigo.Active = False
	QueryCodigo.Add("SELECT HANDLE ")
	QueryCodigo.Add("  FROM SAM_LISTASERVICOS ")
	QueryCodigo.Add(" WHERE CODIGO = :pCodigo ")

	If CurrentQuery.State = 2 Then   'edição
		QueryCodigo.Add(" AND HANDLE <> :pHandleAtual")
		QueryCodigo.ParamByName("pHandleAtual").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	End If

	QueryCodigo.ParamByName("pCodigo").AsString = CurrentQuery.FieldByName("CODIGO").AsString
	QueryCodigo.Active = True

	If Not QueryCodigo.EOF Then
		bsShowMessage("Já existe um registro com este código!", "E")
		CanContinue = False
		Set QueryCodigo = Nothing
		Exit Sub
	End If

	Set QueryCodigo = Nothing

End Sub
