'HASH: D7298211CB3A9E98D07DC1D867B02C3E
Sub Main()
	Dim viPeg As Long
	Dim psTipoInic As String
	Dim xmlInic As String
	Dim vsMsgRetorno As String
	Dim vvPegConf As Object

    vsMsgRetorno = ""

	viPeg = CLng( SessionVar("hpeg") )
	psTipoInic = ServiceVar("psTipoInic")

	On Error GoTo erro

		Set vvPegConf = CreateBennerObject("SAMPEGCONFERENCIA.Rotinas")
		xmlInic = vvPegConf.InitWeb(CurrentSystem, viPeg, psTipoInic)

		ServiceVar("xmlInic") = CStr(xmlInic)
		ServiceVar("vsMsgRetorno") = CStr(vsMsgRetorno)
		ServiceVar("vsUsuario") = CStr(CurrentUser)

		Set vvPegConf = Nothing

		Exit Sub
	erro:
		ServiceVar("vsMsgRetorno") = CStr(Err.Description)
		Set vvPegConf = Nothing
End Sub
