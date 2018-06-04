'HASH: A831802CBE15722E89D947735927A01E
Sub Main()
	Dim piGuiaEvento As Integer

	Dim vvPegConf As Object
	Dim vsMsgRetorno As String

	piHandle = CLng( ServiceVar("piHandle") )
                vsMsgRetorno = ""

	On Error GoTo erro

		Set vvPegConf = CreateBennerObject("SAMPEGCONFERENCIA.Rotinas")
		vsMsgRetorno = vvPegConf.ValorGlosa(CurrentSystem, piGuiaEvento)

		ServiceVar("vsMsgRetorno") = CStr(vsMsgRetorno)
		Set vvPegConf = Nothing

		Exit Sub

	erro:
		ServiceVar("vsMsgRetorno") = CStr(Err.Description)
		Set vvPegConf = Nothing
End Sub
