'HASH: 72D0A969F4717B27AC7660E7710B1A05
'#uses "*BsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("CONTAFINANCEIRA").AsString = Solver(RecordHandleOfTable("SAM_BENEFICIARIO"),"SAM_BENEFICIARIO","NOME")
	CurrentQuery.FieldByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_CONTAFIN")
End Sub

Public Sub TABLE_AfterPost()
	Dim vContainer As CSDContainer
	Set vContainer = NewContainer

	vContainer.GetFieldsFromQuery(CurrentQuery.TQuery)
	vContainer.LoadAllFromQuery(CurrentQuery.TQuery)

	Dim Interface As Object
	Set Interface = CreateBennerObject("SamAdiantamento.Rotinas")
	Interface.AdiantamentoWeb(CurrentSystem,vContainer)

	BsShowMessage("Fatura criada com sucesso.","I")


	Set Interface = Nothing
	Set vContainer = Nothing
End Sub
