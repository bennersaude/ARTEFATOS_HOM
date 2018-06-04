'HASH: D2894B84631F9D941089E3F64E5FD540

Public Sub Main
	Dim vsSenha As String

	vsSenha = UCase(CStr(ServiceVar("Senha")))

	ServiceResult = PasswordEncode(vsSenha)

End Sub
