'HASH: C788959B90C0E926C8451FC0954C6933


Public Sub TABLE_AfterDelete()
  reinicia
End Sub

Public Sub TABLE_AfterPost()
reinicia
End Sub

Sub reinicia
  MsgBox "O sistema deve ser reiniciado para a alteração ter efeito!"
End Sub
