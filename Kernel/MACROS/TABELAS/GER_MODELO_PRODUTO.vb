'HASH: 2C42008518833EFBAA77830E08B32C83
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim SQL As Object
Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT PRODUTO FROM GER_MODELO_PRODUTO WHERE PRODUTO =:HPRODUTO AND HANDLE<>:HANDLEREGISTRO AND MODELO=:HMODELO")
SQL.ParamByName("HPRODUTO").AsString = CurrentQuery.FieldByName("PRODUTO").AsString
SQL.ParamByName("HANDLEREGISTRO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ParamByName("HMODELO").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger
SQL.Active = True

If SQL.FieldByName("PRODUTO").AsString <> "" Then
  MsgBox("Produto já existente!")
  CanContinue = False
  Set SQL = Nothing
  Exit Sub
End If

Set SQL = Nothing

End Sub
