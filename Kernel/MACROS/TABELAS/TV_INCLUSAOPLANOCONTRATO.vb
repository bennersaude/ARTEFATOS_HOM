'HASH: 49156018BA51D6D3720C7F286066F6CC
'#Uses "*bsShowMessage"
Private Function SNParaByte(psValor As String)
	If (psValor = "S") Then
		SNParaByte = 1
	Else
		SNParaByte = 0
	End If
End Function

Private Sub AdicionarTabelaPrecos(ByVal psTabelaPrecos As String, ByVal psValor As String)
	If (psTabelaPrecos = "") Then psTabelaPrecos = psTabelaPrecos + ", "

	psTabelaPrecos = psTabelaPrecos + psValor
End Sub

Private Function TabelasPreco
	Dim vsTabelasPreco As String

	vsTabelasPreco = ""
	If (CurrentQuery.FieldByName("TABPRECOA").AsString = "S") Then
		If Not (vsTabelasPreco = "") Then vsTabelasPreco = vsTabelasPreco + ", "

		vsTabelasPreco = vsTabelasPreco + "'A'"
	End If

	If (CurrentQuery.FieldByName("TABPRECOB").AsString = "S") Then
		If Not (vsTabelasPreco = "") Then vsTabelasPreco = vsTabelasPreco + ", "

		vsTabelasPreco = vsTabelasPreco + "'B'"
	End If

	If (CurrentQuery.FieldByName("TABPRECOC").AsString = "S") Then
		If Not (vsTabelasPreco = "") Then vsTabelasPreco = vsTabelasPreco + ", "

		vsTabelasPreco = vsTabelasPreco + "'C'"
	End If

	If (CurrentQuery.FieldByName("TABPRECOD").AsString = "S") Then
		If Not (vsTabelasPreco = "") Then vsTabelasPreco = vsTabelasPreco + ", "

		vsTabelasPreco = vsTabelasPreco + "'D'"
	End If

	If (CurrentQuery.FieldByName("TABPRECOE").AsString = "S") Then
		If Not (vsTabelasPreco = "") Then vsTabelasPreco = vsTabelasPreco + ", "

		vsTabelasPreco = vsTabelasPreco + "'E'"
	End If

	If (CurrentQuery.FieldByName("TABPRECOF").AsString = "S") Then
		If Not (vsTabelasPreco = "") Then vsTabelasPreco = vsTabelasPreco + ", "

		vsTabelasPreco = vsTabelasPreco + "'F'"
	End If

	If (CurrentQuery.FieldByName("TABPRECOG").AsString = "S") Then
		If Not (vsTabelasPreco = "") Then vsTabelasPreco = vsTabelasPreco + ", "

		vsTabelasPreco = vsTabelasPreco + "'G'"
	End If

	If (CurrentQuery.FieldByName("TABPRECOPADRAO").AsString = "S") Then
		If Not (vsTabelasPreco = "") Then vsTabelasPreco = vsTabelasPreco + ", "

		vsTabelasPreco = vsTabelasPreco + "'P'"
	End If

	TabelasPreco = vsTabelasPreco
End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  	Dim DLLADICIONARPLANO As Object
  	Dim vsMensagemRetorno As String
  	Dim viRetorno As Integer

	If Not (CurrentQuery.FieldByName("MODULOSSELECIONADOS").AsString = "") Then
	    Set DLLADICIONARPLANO = CreateBennerObject("BSBEN007.AdicaoPlano")
	    DLLADICIONARPLANO.ParametrosPlanos(CurrentQuery.FieldByName("CONTRATO").AsInteger, _
	                                       CurrentQuery.FieldByName("PLANO").AsInteger, _
	                                       CurrentQuery.FieldByName("DATAADESAO").AsDateTime, _
	                                       TabelasPreco(), _
	                                       CurrentQuery.FieldByName("MODULOSSELECIONADOS").AsString)
	    DLLADICIONARPLANO.ParametrosReplicar(SNParaByte(CurrentQuery.FieldByName("CARENCIAS").AsString), _
	                     SNParaByte(CurrentQuery.FieldByName("CENTROCUSTO").AsString), _
	                     SNParaByte(CurrentQuery.FieldByName("FRANQUIAS").AsString), _
	                     SNParaByte(CurrentQuery.FieldByName("LIMITACOES").AsString), _
	                     SNParaByte(CurrentQuery.FieldByName("PF").AsString), _
	                     SNParaByte(CurrentQuery.FieldByName("TABELASCOBRANCA").AsString), _
	                     SNParaByte(CurrentQuery.FieldByName("TIPOSDEPENDENTES").AsString))
	    viRetorno = DLLADICIONARPLANO.Exec(CurrentSystem, vsMensagemRetorno, -1)

	    If (viRetorno > 0) Then
			CanContinue = False
			bsShowMessage(vsMensagemRetorno, "E")
		Else
			CanContinue = True
			bsShowMessage("Plano adicionado com sucesso.", "I")
	    End If
	Else
		CanContinue = False
		bsShowMessage("Selecione ao menos um módulo", "E")
	End If

  	Set DLLADICIONARPLANO = Nothing
End Sub
