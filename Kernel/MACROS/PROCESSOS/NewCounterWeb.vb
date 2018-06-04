'HASH: 22F9F423D4ACE1D23468220FDFDC76BB

Public Sub Main

	Dim NomeTabela As String
	Dim Contador As Long

	NomeTabela = CStr(ServiceVar("NOMETABELA"))
	NewCounter(NomeTabela, 0, 1, Contador)

	ServiceResult = CLng(Contador)

End Sub
