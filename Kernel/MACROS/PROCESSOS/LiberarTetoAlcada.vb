'HASH: C807BA2A5D4F8672733C4F782C2B5D23

Public Sub Main
	Dim processo As Long
	processo = CLng(ServiceVar("processo"))

	Dim interface As CSBusinessComponent

	Select Case processo
		Case 1

            Dim dataPagamento As Date
			Dim status As String
			Dim pretador As Long
			Dim tipoBloqueio As Integer

			dataPagamento = CDate(ServiceVar("dataPagamento"))
			status = CStr(ServiceVar("status"))
			prestador = CLng(ServiceVar("prestador"))
			tipoBloqueio = CInt(ServiceVar("tipoBloqueio"))

            Set interface = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.Pagamentos.SamAgrupadorPagamentoBLL, Benner.Saude.ProcessamentoContas.Business")
            interface.AddParameter(pdtDateTime, dataPagamento)
            interface.AddParameter(pdtInteger, prestador)
            interface.AddParameter(pdtInteger, tipoBloqueio)
            interface.AddParameter(pdtString, status)
            ServiceResult = interface.Execute("ConsultarAgrupadorPagamentoParaLiberacaoMediaAlcada")

            Set interface = Nothing

        Case 2
			Dim tipoLiberacao As Integer
			Dim registros As String

			tipoLiberacao = CInt(ServiceVar("tipoLiberacao"))
			registros = CStr(ServiceVar("registros"))

            Set interface = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.Pagamentos.SamAgrupadorPagamentoBLL, Benner.Saude.ProcessamentoContas.Business")
            interface.AddParameter(pdtString, registros)
            interface.AddParameter(pdtInteger, tipoLiberacao)
            ServiceResult = interface.Execute("LiberarPagamentoBloqueado")

            Set interface = Nothing

	End Select

End Sub
