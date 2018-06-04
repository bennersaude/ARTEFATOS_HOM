'HASH: 78E86FE9A84D8DE5835A20E2BA107277
'Macro: SAM_PRECOGENERICO_SL
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = "AND PORTESALA = " + CurrentQuery.FieldByName("PORTESALA").AsString
	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOGENERICO_SL", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "TABELAPRECO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
End Sub
