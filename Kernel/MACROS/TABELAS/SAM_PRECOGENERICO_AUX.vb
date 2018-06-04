'HASH: B2B827BA51257D71AB9C91878771E7B3
'Macro: SAM_PRECOGENERICO_AUX
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = "AND NUMAUXILIAR = " + CurrentQuery.FieldByName("NUMAUXILIAR").AsString
	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOGENERICO_AUX", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "TABELAPRECO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
End Sub
