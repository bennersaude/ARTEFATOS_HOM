'HASH: F17F0B0A40E77C64AEAC6063259BCDA1
Public Sub Main
	Dim Rotina As Long
	Rotina = CLng(ServiceVar("Rotina"))

    Dim interfaceDelphi As Object
    Set interfaceDelphi = CreateBennerObject("SAMCONTAFINANCEIRA.PARCELAMENTO")

Dim vtotal As Double
        	Dim vJuros As Double
	Select Case Rotina
		'consulta de Faturas
		Case 1
			ServiceResult = interfaceDelphi.ConsultarFaturas(CurrentSystem,CLng(ServiceVar("ContaFinanceira")),CLng(ServiceVar("TipoParcelamento")))

        'Previa do parcelamento
        Case 2

            ServiceResult = interfaceDelphi.CalcularParcelas(CurrentSystem,CStr(ServiceVar("Faturas")), CLng(ServiceVar("ContaFinanceira")), CLng(ServiceVar("TipoParcelamento")), CDate(ServiceVar("DataVencimento")), CLng(ServiceVar("DiasPrazo")), CBool(ServiceVar("IncluirJurosMulta")), CLng(ServiceVar("NumeroParcelas")), vtotal, vJuros)

			ServiceVar("ValorTotal") = CCur(vtotal)
			ServiceVar("JurosTotal") = CCur(vJuros)

		'Faturar Competencias Aberta
		Case 3
			ServiceResult = interfaceDelphi.FaturarCompetenciasAbertas(CurrentSystem,CLng(ServiceVar("ContaFinanceira")),CBool(ServiceVar("FaturarCanceladas")))

		Case 4
			Dim docs As String
			ServiceResult = interfaceDelphi.ProcessarParcelamento(CurrentSystem,CLng(ServiceVar("ContaFinanceira")),CLng(ServiceVar("TipoDocumento")),CLng(ServiceVar("TipoParcelamento")),CStr(ServiceVar("Faturas")), CLng(ServiceVar("NumeroParcelas")), CBool(ServiceVar("IncluirJurosMulta")),CDate(ServiceVar("DataVencimento")),CLng(ServiceVar("DiasPrazo")),CLng(ServiceVar("ValorJuros")),CLng(ServiceVar("ValorMulta")), docs)
			ServiceVar("DocsGerados") = CStr(docs)
	End Select
	Set interface = Nothing
End Sub
