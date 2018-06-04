'HASH: 0948106B25E290EAB474F0E65546004D
Sub Main
	Dim xmlFaturaAvulsa As String
	Dim statusExecucao As Long
	Dim mensagemRetorno As String
	Dim SfnFaturaDll As Object
	
	xmlFaturaAvulsa = CStr( ServiceVar("xmlFaturaAvulsa") )
	statusExecucao = CLng( ServiceVar("statusExecucao") )
	mensagemRetorno = CStr( ServiceVar("mensagemRetorno") )
	
	On Error GoTo erro
		Set SfnFaturaDll = CreateBennerObject("SfnFatura.Rotinas")
		statusExecucao = SfnFaturaDll.FaturaAvulsaWeb(CurrentSystem, xmlFaturaAvulsa, mensagemRetorno)
		Set SfnFaturaDll =  Nothing
	
		GoTo fim
	
	erro:
		statusExecucao = 1
		mensagemRetorno = "Erro no WebService: " + Err.Description
	
	fim:
		ServiceVar("statusExecucao") = CLng(statusExecucao)
		ServiceVar("mensagemRetorno") = CStr(mensagemRetorno)
End Sub
