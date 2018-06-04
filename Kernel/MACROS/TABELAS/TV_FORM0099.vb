'HASH: 25C582A484728A1CF130BD3D1DF8E98C
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("HANDLE").AsString = SessionVar("REAJUSTEPARAM")
End Sub

Public Sub TABLE_AfterPost()
		Dim Interface As Object
		Set Interface = CreateBennerObject("BSBEN003.GerarModulo")

		Dim vContainer As CSDContainer
		Set vContainer = NewContainer

		vContainer.GetFieldsFromQuery(CurrentQuery.TQuery)
		vContainer.LoadAllFromQuery(CurrentQuery.TQuery)

		Interface.Exec(CurrentSystem,vContainer)

		Set Interface = Nothing
		Set vContainer = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("EMPRESARIAL").AsString = "N") And (CurrentQuery.FieldByName("FAMILIAR").AsString = "N") And (CurrentQuery.FieldByName("INDIVIDUAL").AsString = "N") Then
		bsShowMessage("Nenhum tipo de contrato selecionado !","E")
		CanContinue = False
	End If
End Sub

