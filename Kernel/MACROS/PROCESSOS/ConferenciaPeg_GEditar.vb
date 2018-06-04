'HASH: 4C5BA5E9826A715603A6A9FCBC36F6A7
Sub Main()
	Dim piHandle As Integer
	Dim psGlosaDependente As String
	Dim psGlosaRevisada As String
	Dim psComplemento As String
	Dim pfValorReconsid As Double
	Dim pfQtdReconsid As Double

	Dim vvPegConf As Object
	Dim vsMsgRetorno As String

	piHandle = CLng( ServiceVar("piHandle") )
	psGlosaDependente = CStr( ServiceVar("psGlosaDependente") )
	psGlosaRevisada = CStr( ServiceVar("psGlosaRevisada") )
	psComplemento = CStr( ServiceVar("psComplemento") )
	pfValorReconsid = CDbl( ServiceVar("pfValorReconsid") )
	pfQtdReconsid = CDbl( ServiceVar("pfQtdReconsid") )

                vsMsgRetorno = ""


	On Error GoTo erro

		Set vvPegConf = CreateBennerObject("SAMPEGCONFERENCIA.Rotinas")
		vsMsgRetorno = vvPegConf.EditarGlosa(CurrentSystem, piHandle, psGlosaDependente, psGlosaRevisada, psComplemento, pfValorReconsid, pfQtdReconsid)

		ServiceVar("vsMsgRetorno") = CStr(vsMsgRetorno)
		Set vvPegConf = Nothing

		Exit Sub

	erro:
		ServiceVar("vsMsgRetorno") = CStr(Err.Description)
		Set vvPegConf = Nothing
End Sub
