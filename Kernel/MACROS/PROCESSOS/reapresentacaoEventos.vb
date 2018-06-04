'HASH: 0E8FCB0645AE10157BE9431301BA9551

Public Sub Main

	Dim psMsgRetorno As String
	Dim piResult As Long
	Dim BsServerExec As Object
	Dim vcContainer As CSDContainer
	Dim xmlData As String
	Dim numeropeg As Long
	Dim cartaRemessa As Long
	Dim handlepeg As Long
	Dim handlePorPeg As Long
	Dim agendado As Boolean
	Dim sequencia As Long
	Dim SQL As Object
	Dim handlePegReapresentacao As Long

	On Error GoTo Erro

	xmlData = CStr(ServiceVar("xmlData"))
	psMsgRetorno = CStr(ServiceVar("mensagem"))
	numeropeg = CLng(ServiceVar("numeroPeg"))
	cartaRemessa = CLng(ServiceVar("cartaRemessa"))
	handlePorPeg = CLng(ServiceVar("handlePorPeg"))
	agendado = CBool(ServiceVar("agendado"))
	sequencia = CLng(ServiceVar("sequencia"))
	handlePegReapresentacao = CLng(ServiceVar("handlePeg"))
	ServiceVar("handlePeg") = 0

	Set SQL = NewQuery

    SQL.Clear
    SQL.Add(" UPDATE SAM_PEG                     ")
    SQL.Add("   SET SITUACAOREAPRESENTACAO = '2' ")
    SQL.Add("  WHERE HANDLE = :PEG               ")
    SQL.ParamByName("PEG").AsInteger = handlePegReapresentacao
    SQL.ExecSQL

	If (agendado) Then

		psMsgRetorno = "Reapresentação de Eventos enviada para processamento no servidor."
		piResult = 0

		Set vcContainer = NewContainer
		vcContainer.AddFields("XMLDATA:STRING")
		vcContainer.AddFields("NUMEROPEG:INTEGER")
		vcContainer.AddFields("CARTAREMESSA:INTEGER")
		vcContainer.AddFields("HANDLEPORPEG:INTEGER")
		vcContainer.AddFields("HANDLE:INTEGER")
		vcContainer.AddFields("SEQUENCIA:INTEGER")

		vcContainer.Insert
		vcContainer.Field("XMLDATA").AsString = xmlData
		vcContainer.Field("NUMEROPEG").AsInteger = numeropeg
		vcContainer.Field("CARTAREMESSA").AsInteger = cartaRemessa
		vcContainer.Field("HANDLEPORPEG").AsInteger = handlePorPeg
		vcContainer.Field("HANDLE").AsInteger = handlePorPeg
		vcContainer.Field("SEQUENCIA").AsInteger = sequencia

		Set BsServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
		piResult = BsServerExec.ExecucaoImediata(CurrentSystem, _
			"BSPRO006", _
			"REAPRESENTACAO", _
			"Processo de Reapresentação de Eventos", _
			handlePorPeg, _
			"POR_PEGREAPRESENTADO", _
			"SITUACAOPROCESSAMENTO", _
			"", _
			"", _
			"P", _
			False, _
			psMsgRetorno, _
			vcContainer)


		GoTo Fim

	Else
		Dim Interface As Object
		Set Interface = CreateBennerObject("BSPRO006.REAPRESENTACAO")

		handlepeg = Interface.Processar(CurrentSystem, _
										xmlData, _
										numeropeg, _
										cartaRemessa, _
										0, _
										psMsgRetorno)

		GoTo Fim
	End If

	Erro:
		SQL.Clear
    	SQL.Add(" UPDATE SAM_PEG                     ")
    	SQL.Add("   SET SITUACAOREAPRESENTACAO = '1' ")
    	SQL.Add("  WHERE HANDLE = :PEG               ")
    	SQL.ParamByName("PEG").AsInteger = handlePegReapresentacao
    	SQL.ExecSQL

		piResult = 1
		psMsgRetorno = Err.Description

	Fim:
		Set BsServerExec = Nothing
		Set vcContainer = Nothing
		Set Interface = Nothing

		ServiceVar("agResult") = CLng(piResult)
		ServiceVar("mensagem") = CStr(psMsgRetorno)
		ServiceVar("handlePeg") = CLng(handlepeg)

		Set SQL = Nothing
End Sub
