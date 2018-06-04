'HASH: 2C49C31870E5F178F717A2264FBF5A22
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim SQL As Object
Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT COLUNA FROM GER_MODELO_COLUNA WHERE COLUNA =:HCOLUNA AND HANDLE<>:HANDLEREGISTRO AND MODELO=:HMODELO")
SQL.ParamByName("HCOLUNA").AsString = CurrentQuery.FieldByName("COLUNA").AsString
SQL.ParamByName("HANDLEREGISTRO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ParamByName("HMODELO").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger
SQL.Active = True

If SQL.FieldByName("COLUNA").AsString <> "" Then
  MsgBox("Coluna já existente!")
  CanContinue = False
  Set SQL = Nothing
  Exit Sub
End If

Set SQL = Nothing

End Sub
