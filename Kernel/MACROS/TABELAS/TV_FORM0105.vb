'HASH: F23EA81B04BBB3F9AEA266A43930C93D
'#Uses "*bsShowMessage"
Public Sub TABLE_AfterPost()

	Dim vContainer As CSDContainer
	Dim vInterface As Object
	Dim msgArray() As String

	Dim TipoImportaFaixaContrato As String
	TipoImportaFaixaContrato = SessionVar("TIPOIMPORTAFAIXA")

	Set vContainer = NewContainer
	vContainer.GetFieldsFromQuery(CurrentQuery.TQuery)
	vContainer.LoadAllFromQuery(CurrentQuery.TQuery)

	If (TipoImportaFaixaContrato = "C") Then

		Set vInterface = CreateBennerObject("SamImportafaixaPrc.Geral")
		vSaida = vInterface.ImportarFaixas(CurrentSystem,  RecordHandleOfTable("SAM_CONTRATO_MODADESAOPRC"), "C", vContainer)
	Else
		Set vInterface = CreateBennerObject("SamImportafaixaPrc.Geral")
		vSaida = vInterface.ImportarFaixas(CurrentSystem,  RecordHandleOfTable("SAM_FAMILIA_MODADESAOPRC"), "F", vContainer)
	End If


	msgArray = Split(vSaida,":")
	If(VisibleMode) Then
		If(msgArray(0) = "erro") Then
			bsShowMessage(msgArray(1), "E")
		Else
			bsShowMessage(msgArray(1), "I")
		End If
	Else
		bsShowMessage(msgArray(1), "I")
	End If
	Set vContainer = Nothing
	Set vInterface = Nothing
End Sub
