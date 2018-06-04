'HASH: F7F712C71182FB9A67ED0F0055BC06D6
Sub Main()
	Dim piGuiaEvento As Integer
	Dim piMotivoGlosa As Integer
	Dim psGlosaDependente As String
	Dim psRevisada As String
	Dim psComplemento As String
	Dim pfValorReconsid As Double
	Dim pfQtdReconsid As Double
	Dim piNovoHandleGlosa As Long

	Dim vsMsgRetorno As String
	Dim vvPegConf As Object

	piGuiaEvento = CLng( ServiceVar("piGuiaEvento") )
	piMotivoGlosa = CLng( ServiceVar("piMotivoGlosa") )
	psGlosaDependente = CStr( ServiceVar("psGlosaDependente") )
	psRevisada = CStr( ServiceVar("psRevisada") )
	psComplemento = CStr( ServiceVar("psComplemento") )
	pfValorReconsid = CDbl( ServiceVar("pfValorReconsid") )
	pfQtdReconsid = CDbl( ServiceVar("pfQtdReconsid") )

               vsMsgRetorno = ""


	On Error GoTo erro

		Set vvPegConf = CreateBennerObject("SAMPEGCONFERENCIA.Rotinas")
		vsMsgRetorno = vvPegConf.AdicionarGlosa(CurrentSystem, piGuiaEvento, piMotivoGlosa, psGlosaDependente, psRevisada, psComplemento, pfValorReconsid, pfQtdReconsid, piNovoHandleGlosa)

		ServiceVar("piNovoHandleGlosa") = CStr(piNovoHandleGlosa)
        ServiceVar("vsMsgRetorno") = CStr(vsMsgRetorno)

		Set vvPegConf = Nothing

		Exit Sub

	erro:
		ServiceVar("vsMsgRetorno") = CStr(Err.Description)
		Set vvPegConf = Nothing
End Sub
