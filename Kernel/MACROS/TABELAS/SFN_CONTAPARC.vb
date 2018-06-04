'HASH: 2E0BD00E0DD1B78A24AD388DDFAD23FC
'Macro: SFN_CONTAPARC

Public Sub BOTAOATUALIZAR_OnClick()
  Dim Interface As Object
End Sub

Public Sub TABINSTRUCAO_OnChanging(AllowChange As Boolean)
  AllowChange = false
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  CurrentQuery.FieldByName("DATAINSTRUCAO").Value = ServerDate()

End Sub


