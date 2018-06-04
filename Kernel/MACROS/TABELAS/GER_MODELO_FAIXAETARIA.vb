'HASH: 5B644D0536E89EA6124705120C175AFE
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim SQL As Object
Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT FAIXAETARIA FROM GER_MODELO_FAIXAETARIA WHERE FAIXAETARIA =:HFAIXAETARIA AND HANDLE<>:HANDLEREGISTRO AND MODELO=:HMODELO")
SQL.ParamByName("HFAIXAETARIA").AsString = CurrentQuery.FieldByName("FAIXAETARIA").AsString
SQL.ParamByName("HANDLEREGISTRO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ParamByName("HMODELO").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger

SQL.Active = True

If SQL.FieldByName("FAIXAETARIA").AsString <> "" Then
  MsgBox("Faixa etária já existente!")
  CanContinue = False
  Set SQL = Nothing
  Exit Sub
End If

Set SQL = Nothing


End Sub
