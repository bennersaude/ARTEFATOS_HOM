'HASH: A9BC2614C69E4D9FF20BFCEFDDECFBF7
 
'#Uses "*bsShowMessage"

Dim pValor As Double
Dim pValorTotal As Double
Dim pValorJuros As Double
Dim pValorMulta As Double
Dim pValorCorrecao As Double
Dim pValorDesconto As Double
Dim pVencimentoAnterior As Date
Dim pNovoVencimento As Date
Dim pAlteraFatura As Boolean

Public Sub TABLE_AfterInsert()
	Dim qDocumento As BPesquisa
	Set qDocumento = NewQuery

	qDocumento.Clear
	qDocumento.Add("SELECT * FROM SFN_DOCUMENTO WHERE HANDLE = :HANDLE")
	qDocumento.ParamByName("HANDLE").AsInteger = CLng(SessionVar("WebHandleDocumento")) 'Variavel de sessão criada no afterscroll da tabela 'SFN_DOCUMENTO'
	qDocumento.Active = True

	CurrentQuery.FieldByName("VENCIMENTOANTERIOR").AsDateTime = qDocumento.FieldByName("DATAVENCIMENTO").AsDateTime
	CurrentQuery.FieldByName("NOVOVENCIMENTO").AsDateTime = Format(SessionVar("NovoVencimento"),"DD/MM/YYYY")


	pValor = CurrentQuery.FieldByName("VALOR").AsFloat
	pValorTotal = CurrentQuery.FieldByName("VALORTOTAL").AsFloat
	pValorJuros = CurrentQuery.FieldByName("VALORJUROS").AsFloat
	pValorMulta = CurrentQuery.FieldByName("VALORMULTA").AsFloat
	pValorCorrecao = CurrentQuery.FieldByName("VALORCORRECAO").AsFloat
	pValorDesconto = CurrentQuery.FieldByName("VALORDESCONTO").AsFloat


	Dim SAMCONTAFINANC As Object
	Set SAMCONTAFINANC = CreateBennerObject("SamContaFinanceira.Consulta")

	vsMensagem = SAMCONTAFINANC.BxCalcDocumento(CurrentSystem, _
									     		CLng(SessionVar("WebHandleDocumento")), _
										 		CurrentQuery.FieldByName("VENCIMENTOANTERIOR").AsDateTime, _
										 		CurrentQuery.FieldByName("NOVOVENCIMENTO").AsDateTime, _
										 		pValor, _
										 		pValorTotal, _
										 		pValorJuros, _
										 		pValorMulta, _
										 		pValorCorrecao, _
										 		pValorDesconto)


	CurrentQuery.FieldByName("VALOR").AsFloat = pValor
	CurrentQuery.FieldByName("VALORTOTAL").AsFloat = pValorTotal
	CurrentQuery.FieldByName("VALORJUROS").AsFloat = pValorJuros
	CurrentQuery.FieldByName("VALORMULTA").AsFloat = pValorMulta
	CurrentQuery.FieldByName("VALORCORRECAO").AsFloat = pValorCorrecao
	CurrentQuery.FieldByName("VALORDESCONTO").AsFloat = pValorDesconto


	CurrentQuery.FieldByName("VALORTOTAL").AsFloat = (CurrentQuery.FieldByName("VALOR").AsFloat + CurrentQuery.FieldByName("VALORJUROS").AsFloat + CurrentQuery.FieldByName("VALORMULTA").AsFloat + CurrentQuery.FieldByName("VALORCORRECAO").AsFloat) - CurrentQuery.FieldByName("VALORDESCONTO").AsFloat

	qDocumento.Active = False

	Set qDocumento = Nothing
End Sub

Public Sub TABLE_AfterPost()
	Dim vsMensagem As String

	Dim SAMCONTAFINANC As Object
	Set SAMCONTAFINANC = CreateBennerObject("SamContaFinanceira.Consulta")

	vsMensagem = SAMCONTAFINANC.AlterarDataVencimentoDocumento(CurrentSystem, _
									     	CLng(SessionVar("WebHandleDocumento")), _
											pNovoVencimento, _
											pVencimentoAnterior, _
											pValor, _
											pValorTotal, _
											pValorJuros, _
											pValorMulta, _
											pValorCorrecao, _
											pValorDesconto, _
										 	pAlteraFatura)

	If (vsMensagem <> "") Then
		bsShowMessage(vsMensagem, "I")
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
    pVencimentoAnterior = CurrentQuery.FieldByName("VENCIMENTOANTERIOR").AsDateTime
    pNovoVencimento = CurrentQuery.FieldByName("NOVOVENCIMENTO").AsDateTime
	pValor = CurrentQuery.FieldByName("VALOR").AsFloat
	pValorTotal = CurrentQuery.FieldByName("VALORTOTAL").AsFloat
	pValorJuros = CurrentQuery.FieldByName("VALORJUROS").AsFloat
	pValorMulta = CurrentQuery.FieldByName("VALORMULTA").AsFloat
	pValorCorrecao = CurrentQuery.FieldByName("VALORCORRECAO").AsFloat
	pValorDesconto = CurrentQuery.FieldByName("VALORDESCONTO").AsFloat
	pAlteraFatura = CurrentQuery.FieldByName("ALTERARFATURASDOCUMENTO").AsBoolean
End Sub
