'HASH: 589C3AB0D3A571198DF92F77203A9DA9
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime) And (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    MsgBox "A data inicial da turma não pode ser superior à data final!"
    CanContinue = False
    Exit Sub
  End If
End Sub
