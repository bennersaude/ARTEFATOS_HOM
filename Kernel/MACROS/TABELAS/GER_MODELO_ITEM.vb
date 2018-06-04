'HASH: 0917E197CDBB20C0A21BCD1F5E8F161D
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim SQL As Object
Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT ITEM FROM GER_MODELO_ITEM WHERE ITEM =:HITEM AND HANDLE<>:HANDLEREGISTRO AND MODELO=:HMODELO")
SQL.ParamByName("HITEM").AsString = CurrentQuery.FieldByName("ITEM").AsString
SQL.ParamByName("HANDLEREGISTRO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ParamByName("HMODELO").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger

SQL.Active = True

If SQL.FieldByName("ITEM").AsString <> "" Then
  MsgBox("Item já existente!")
  CanContinue = False
  Set SQL = Nothing
  Exit Sub
End If

Set SQL = Nothing


End Sub
