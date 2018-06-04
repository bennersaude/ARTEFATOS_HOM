'HASH: 8F22305692840128B3743656CBCBD5CB

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("LONGE").IsNull And CurrentQuery.FieldByName("PERTO").IsNull Then
    MsgBox "Deve ser informado pelo meno um tipo de exame para olho direito"
    CanContinue = False
  End If
  If CurrentQuery.FieldByName("LONGEE").IsNull And CurrentQuery.FieldByName("PERTOE").IsNull Then
    MsgBox "Deve ser informado pelo meno um tipo de exame para olho esquerdo "
    CanContinue = False
  End If

End Sub


Public Sub TABLE_NewRecord()
  If VisibleMode Then
    CurrentQuery.FieldByName("FATOGERADOR").Value = 6
  End If
End Sub


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

