'HASH: 7C8C3AD4D6533F214966EA3769C796D4

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If ((SQL.FieldByName("DATAINICIAL").IsNull) Or ((Not SQL.FieldByName("DATAINICIAL").IsNull) And (Not SQL.FieldByName("DATAFINAL").IsNull))) Then
    MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
    CanContinue = False
    Exit Sub
  End If
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If ((SQL.FieldByName("DATAINICIAL").IsNull) Or ((Not SQL.FieldByName("DATAINICIAL").IsNull) And (Not SQL.FieldByName("DATAFINAL").IsNull))) Then
    MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
    CanContinue = False
    Exit Sub
  End If

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If ((SQL.FieldByName("DATAINICIAL").IsNull) Or ((Not SQL.FieldByName("DATAINICIAL").IsNull) And (Not SQL.FieldByName("DATAFINAL").IsNull))) Then
    MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
    CanContinue = False
    Exit Sub
  End If
End Sub

