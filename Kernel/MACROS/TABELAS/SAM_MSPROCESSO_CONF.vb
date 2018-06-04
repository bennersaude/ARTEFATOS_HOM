'HASH: 36FC5574343F62F5B9C9EB5D7CC018F2
'Macro: SAM_MSPROCESSO_CONF
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString = "3" Then
  	bsShowMessage("Não é permitido alterar a Situação para 'Enviado para ANS'!", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
