'HASH: C992C2C603E5DFE5F4777EC57719DDC3

Public Sub Main

	Dim HandleTabela As Long
	Dim HandleRegistro As Long

	HandleTabela = CLng(ServiceVar("HANDLETABELA"))
	HandleRegistro = CLng(ServiceVar("HANDLEREGISTRO"))

	WriteAudit("I", HandleTabela, HandleRegistro, "Auditoria do Gera Lista")

End Sub
