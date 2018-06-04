'HASH: DE0CB67317D0E1F6E84E147AE1D388AD
'#Uses "*bsShowMessage

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim ModificarCodigoFolha As CSBusinessComponent
  Dim mensagem As String


  Set ModificarCodigoFolha = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.SfnFaturaBLL, Benner.Saude.Financeiro.Business")

  ModificarCodigoFolha.AddParameter(pdtInteger, CLng(SessionVar("HFATURA")))
  ModificarCodigoFolha.AddParameter(pdtInteger, CurrentQuery.FieldByName("NOVOCODIGOFOLHA").AsInteger)

  mensagem = ModificarCodigoFolha.Execute("ModificarCodigoFolha")

  If mensagem <> "" Then
    bsshowmessage(mensagem, "I")
  End If

  Set ModificarCodigoFolha = Nothing
End Sub
