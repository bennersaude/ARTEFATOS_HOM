'HASH: CA1E49E680AA0CB83E49AD1E98524706
'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If CurrentQuery.FieldByName("VALIDACAOPROCEDIMENTO").AsString = "N" And	CurrentQuery.FieldByName("PERMITEEVENTOSDUPLICADOS").AsString = "S" Then
		bsShowMessage("'Permitir Eventos Duplicados' só poderá ser marcado se a validação de consistências de existência de procedimento no sistema estiver ativa (Flag Procedimento marcado)!", "E")
		CanContinue = False
	End If

	If CurrentQuery.FieldByName("VALIDACAOPROCEDIMENTO").AsString = "N" And	CurrentQuery.FieldByName("VALIDACAOPROCVIGENTE").AsString = "S" Then
		bsShowMessage("'Procedimento Vigente' só poderá ser marcado se a validação de consistências de existência de procedimento no sistema estiver ativa (Flag Procedimento marcado)!", "E")
		CanContinue = False
	End If

End Sub
