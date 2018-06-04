'HASH: 336170700E0D67403580B1BE7CB37B48
Public Sub Main
	Dim result As String
	Dim precoTotal As String
	Dim valorPF As String
	Dim mensagem As String
	Dim retorno As Integer

	Set interface =CreateBennerObject("PRECO.CONSULTAPRECOEVENTO")
	retorno = interface.Exec(CurrentSystem, CDate(ServiceVar("P_DATA")), CLng(ServiceVar("P_GRAU")), CLng(ServiceVar("P_EVENTO")), CLng(ServiceVar("P_TABXTHM")), _
                     CLng(ServiceVar("P_ACOMODACAO")), CLng(ServiceVar("P_XTHM")), _
					 CLng(ServiceVar("P_CODPAGAMENTO")), CLng(ServiceVar("P_QUANTIDADE")), _
                     CLng(ServiceVar("P_RECEBEDOR")), CLng(ServiceVar("P_LOCALEXECUCAO")), 0, 0, 0, _
                     CLng(ServiceVar("P_CONVENIO")), 0, 0, CLng(ServiceVar("P_EXECUTOR")), _
                     CLng(ServiceVar("P_LOCALATENDIMENTO")), CLng(ServiceVar("P_CONDICAOATENDIMENTO")), _
                     CLng(ServiceVar("P_REGIMEATENDIMENTO")), CLng(ServiceVar("P_TIPOTRATAMENTO")), _
                     CLng(ServiceVar("P_OBJETIVOTRATAMENTO")), CLng(ServiceVar("P_FINALIDADEATENDIMENTO")), _
                     CLng(ServiceVar("P_CBOS")),CDate(ServiceVar("P_HORAATENDIMENTO")),CBool(ServiceVar("P_HORARIOESPECIAL")), _
                     "1", precoTotal, valorPF, mensagem)

	Set interface = Nothing

	If (retorno > 0) Then
	  result = mensagem
	Else
	  result = "Valor do Evento = R$ " + precoTotal
	End If

	ServiceResult = result
End Sub
