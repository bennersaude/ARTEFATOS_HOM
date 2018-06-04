'HASH: C8CB0D44CF7F6F0F33D6502781C6ED3B

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim SQL As Object
Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT GRUPOCCUSTO FROM GER_MODELO_GRUPOCCUSTO WHERE GRUPOCCUSTO =:HGRUPOCCUSTO And HANDLE<>:HANDLEREGISTRO AND MODELO=:HMODELO")
SQL.ParamByName("HGRUPOCCUSTO").AsString = CurrentQuery.FieldByName("GRUPOCCUSTO").AsString
SQL.ParamByName("HANDLEREGISTRO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ParamByName("HMODELO").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger

SQL.Active = True

If SQL.FieldByName("GRUPOCCUSTO").AsString <> "" Then
  MsgBox("Grupo de centro de custo já existente!")
  CanContinue = False
  Set SQL = Nothing
  Exit Sub
End If

Set SQL = Nothing


End Sub
