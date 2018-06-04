'HASH: 9132FA760962AB69153A4DCB6C6CD70F
Sub Main()
	Dim piGlosa As Integer
	Dim vsDescricaoTiss As String
	Dim vvPegConf As Object

	Dim vsMsgRetorno As String

               vsDescricaoTiss = ""
                vsMsgRetorno = ""

	piGuiaEvento = CLng( ServiceVar("piGlosa") )


	On Error GoTo erro

		Set vvPegConf = CreateBennerObject("SAMPEGCONFERENCIA.Rotinas")
		vsDescricaoTiss = vvPegConf.BuscarNegacaoGlosaTISS(CurrentSystem, piGlosa)

		ServiceVar("vsDescricaoTiss") = CStr(vsDescricaoTiss)

		Set vvPegConf = Nothing

		Exit Sub

	erro:
		ServiceVar("vsMsgRetorno") = CStr(Err.Description)
		Set vvPegConf = Nothing
End Sub
