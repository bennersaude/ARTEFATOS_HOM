'HASH: D162B26119536F85947812FA26424858

Public Sub TABLE_AfterEdit()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterInsert()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  Dim query As Object

  Set query = NewQuery

  query.Clear
  query.Add("SELECT TABFISICAJURIDICA FROM SFN_PESSOA WHERE HANDLE = :HPESSOA")
  query.ParamByName("HPESSOA").AsInteger = CurrentQuery.FieldByName("PESSOA").AsInteger
  query.Active = True

  If query.FieldByName("TABFISICAJURIDICA").AsInteger = 1 Then
    TABCONTRIBUICOESFEDERAIS.ReadOnly = True
  Else
    TABCONTRIBUICOESFEDERAIS.ReadOnly = False
  End If

  Set query = Nothing
End Sub

