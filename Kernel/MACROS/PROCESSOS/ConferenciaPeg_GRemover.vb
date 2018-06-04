'HASH: C337397A5CF057D12A61C0494B03638D
Sub Main()
	Dim piHandle As Integer

	Dim vvPegConf As Object
	Dim vsMsgRetorno As String

	piHandle = CLng( ServiceVar("piHandle") )
                vsMsgRetorno = ""

	On Error GoTo erro

		Set vvPegConf = CreateBennerObject("SAMPEGCONFERENCIA.Rotinas")
		vsMsgRetorno = vvPegConf.RemoverGlosa(CurrentSystem, piHandle)

		ServiceVar("vsMsgRetorno") = CStr(vsMsgRetorno)
		Set vvPegConf = Nothing

		Exit Sub

	erro:
		ServiceVar("vsMsgRetorno") = CStr(Err.Description)
		Set vvPegConf = Nothing
End Sub
