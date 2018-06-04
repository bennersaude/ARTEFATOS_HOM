'HASH: 326E2594180ACD8C8944B1094CD58D32


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT HANDLE")
  SQL.Add("  FROM GER_ITEM_REGIMEATENDIMENTO")
  SQL.Add(" WHERE ITEM = :ITEM")
  SQL.Add("   AND HANDLE <> :HANDLE")
  SQL.Add("   AND REGIMEATENDIMENTO = :REGIMEATENDIMENTO")
  SQL.ParamByName("ITEM").AsInteger = RecordHandleOfTable("GER_ITEM")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("REGIMEATENDIMENTO").AsInteger = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    MsgBox("Este regime de atendimento já está cadastrado para este item.")
    CanContinue = False
  End If
  SQL.Active = False
  Set SQL = Nothing
End Sub

