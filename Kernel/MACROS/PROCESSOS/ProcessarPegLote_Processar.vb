'HASH: 122E013D168F179842ED3682E6F45A1F

Public Sub Main
	Dim psPEGs As String
	Dim psMsgRetorno As String
	Dim piResult As Long
	Dim BsServerExec As Object
	Dim vcContainer As CSDContainer

	psPEGs = CStr( ServiceVar("psPEGs") )
	psMsgRetorno = CStr( ServiceVar("psMsgRetorno") )
	piResult = CLng( ServiceVar("piResult") )

	On Error GoTo erro

	psMsgRetorno = "PEG(s) enviado(s) para processamento no servidor"
	piResult = 0

	Set vcContainer = NewContainer
    vcContainer.AddFields("PEGS:STRING")
	vcContainer.Insert
	vcContainer.Field("PEGS").AsString = psPEGs

	Set BsServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
	piResult = BsServerExec.ExecucaoImediata(CurrentSystem, _
		"SAMPEG", _
		"PEGLote_Processar", _
		"Processamento de PEG em Lote - (" + Str(ServerNow) + ")", _
		0, "", "", "", "", "P", False, _
		psMsgRetorno, _
		vcContainer)

    GoTo fim

    erro:
		piResult = 1
		psMsgRetorno = Err.Description

    fim:
		Set BsServerExec = Nothing
		Set vcContainer = Nothing

		ServiceVar("psMsgRetorno") = CStr(psMsgRetorno)
		ServiceVar("piResult") = CLng(piResult)

End Sub
