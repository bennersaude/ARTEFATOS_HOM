'HASH: 42B9FD5BEA10101F5856147786964CEF
 '#Uses "*bsShowMessage"

 Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("ESTADO").AsString = "" And CurrentQuery.FieldByName("REGIAOSAUDE").AsString = "" Then

    bsShowMessage("Necessário informar o Estado ou Região Saúde!", "I")
    CanContinue = False
  End If

End Sub
