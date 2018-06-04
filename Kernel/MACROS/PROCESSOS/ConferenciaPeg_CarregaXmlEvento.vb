'HASH: CA13430BD3DBE882753A3A906CB827C3
Public Sub Main
	Dim viHPeg As Long
	Dim viLinhaInicial As Long
	Dim xmlEvento As String
	Dim vsMsgRetorno As String
	Dim vvPegConf As Object

	viHPeg = CLng( ServiceVar("pHPeg") )
	viLinhaInicial = CLng( ServiceVar("pLinhaInicial") )

	On Error GoTo erro

		Set vvPegConf = CreateBennerObject("SAMPEGCONFERENCIA.Rotinas")
		xmlEvento = vvPegConf.CarregaXmlEvento(CurrentSystem, viHPeg, viLinhaInicial)

		ServiceVar("xmlEvento") = CStr(xmlEvento)

		Set vvPegConf = Nothing

		Exit Sub
	erro:
		ServiceVar("msgRetorno") = CStr(Err.Description)
		Set vvPegConf = Nothing
End Sub
