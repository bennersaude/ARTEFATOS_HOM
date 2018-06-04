'HASH: F79859F8A00FAF89427DA60D1A0E69BF
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = "AND MATMEDPRECOTAB =" + CurrentQuery.FieldByName("MATMEDPRECOTAB").AsString

	Linha = Interface.Vigencia(CurrentSystem, "SAM_MATMEDPRECOTAB_VLR", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "MATMED", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
		Exit Sub
	End If
End Sub

