'HASH: 371E84F3BB872D60AE1D791F2837906B

Public Sub Main

Dim vsSenha As String
vsSenha = ServiceVar("SENHA")
ServiceResult = PasswordEncode(UCase(vsSenha))

End Sub
