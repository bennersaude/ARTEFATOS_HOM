'HASH: A4867A70759AB97FEF70ABE13B1F0BFE
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	Dim qVersaoTIss As BPesquisa

	Set qVersaoTIss = NewQuery



	qVersaoTIss.Clear
	qVersaoTIss.Active = False
	qVersaoTIss.Add("SELECT MAX(HANDLE) HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'")
	qVersaoTIss.Active = True

	CurrentQuery.FieldByName("VERSAOTISS").Value = qVersaoTIss.FieldByName("HANDLE").Value

	Set qVersaoTIss = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
    Dim DLLPreco As Object
    Dim qCBOS As BPesquisa
    Dim linha As String

    Set qCBOS = NewQuery

    qCBOS.Clear
	qCBOS.Active = False
	qCBOS.Add("SELECT CODIGO, HANDLE FROM TIS_CBOS WHERE HANDLE =:CBOS ")
	qCBOS.ParamByName("CBOS").AsInteger = CurrentQuery.FieldByName("CBOS").AsInteger
	qCBOS.Active = True

    Set DLLPreco = CreateBennerObject("Preco.PegaPreco")

    linha = DLLPreco.ReplicarCBOS(CurrentSystem, qCBOS.FieldByName("CODIGO").AsString, qCBOS.FieldByName("HANDLE").AsInteger)

    If (linha <> "") Then
		bsShowMessage("Preço Não Replicado. " + linha, "I")
	Else
		bsShowMessage("Preço Replicado com Sucesso", "I")
    End If


End Sub
