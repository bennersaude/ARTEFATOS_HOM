'HASH: EB846AE1A874ACE52FAC495F7D1FB8D6

Public Sub Main

	SessionVar("ASSUNTO") = ServiceVar("pAssunto")
	SessionVar("EMAILREMETENTE") = ServiceVar("pEmailRemetente")
	SessionVar("EMAILRECEBEDOR") = ServiceVar("pEmailDestinatario")
	SessionVar("MENSAGEM") = ServiceVar("pMensagem")
End Sub
