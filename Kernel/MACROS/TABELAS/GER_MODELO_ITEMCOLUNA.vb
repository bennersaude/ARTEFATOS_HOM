'HASH: 7DAF1441529F8C74A4A1755B351D7FE0
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim SQL As Object
Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT COLUNA FROM GER_MODELO_ITEMCOLUNA WHERE COLUNA =:HCOLUNA AND HANDLE<>:HANDLEREGISTRO AND MODELO=:HMODELO AND MODELOITEM=:HMODELOITEM")
SQL.ParamByName("HCOLUNA").AsString = CurrentQuery.FieldByName("COLUNA").AsString
SQL.ParamByName("HANDLEREGISTRO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ParamByName("HMODELO").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger
SQL.ParamByName("HMODELOITEM").AsInteger = CurrentQuery.FieldByName("MODELOITEM").AsInteger

SQL.Active = True

If SQL.FieldByName("COLUNA").AsString <> "" Then
  MsgBox("Item já existente!")
  CanContinue = False
  Set SQL = Nothing
  Exit Sub
End If

Set SQL = Nothing



End Sub
