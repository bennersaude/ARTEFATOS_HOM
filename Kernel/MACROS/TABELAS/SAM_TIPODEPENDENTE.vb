'HASH: 883F720002D9AF67181FD54F1BAC47FB

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("GRUPODEPENDENTE").AsString <> "T" And CurrentQuery.FieldByName("AUTONOMO").AsString = "S" Then
    CanContinue = False
    MsgBox "O campo Autônomo somente pode ser marcado se o tipo do dependente for Titular"
  End If
End Sub

