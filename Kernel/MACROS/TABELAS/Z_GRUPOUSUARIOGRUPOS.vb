'HASH: 1DD8B99071A662B3AD2AB59F5D609498


Public Sub TABLE_AfterInsert()
  Dim Q As Object
  Set Q = NewQuery
  'Q.Add("SELECT GRUPO FROM Z_GRUPOUSUARIOS WHERE HANDLE = " + CStr(RecordHandleOfTable("Z_GRUPOUSUARIOS")))

  Q.Add("SELECT GRUPO FROM Z_GRUPOUSUARIOS WHERE HANDLE = " + CStr(CurrentQuery.FieldByName("USUARIO").AsInteger))

  Q.Active = True
  CurrentQuery.FieldByName("GRUPO").AsInteger = Q.FieldByName("GRUPO").AsInteger
  Q.Active = False
  Set Q = Nothing
End Sub

