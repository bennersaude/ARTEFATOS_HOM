'HASH: EE86C6423DFAA7AC6E988FC61DDF93CE

Public Sub Main
	' Codifique aqui o método principal
	Dim NomeTabela As String
	Dim handle As Long

	NomeTabela = CStr(ServiceVar("NOMETABELA"))
	handle = NewHandle(NomeTabela)

	ServiceResult = CLng(handle)

End Sub
