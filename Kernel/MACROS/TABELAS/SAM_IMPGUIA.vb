'HASH: E63EB3773BB94BBFB3CFEF1E0EAAFBF8
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_IMPGUIA WHERE TIPO=:TIPO AND HANDLE<>:HANDLE")
  SQL.ParamByName("TIPO").AsString = CurrentQuery.FieldByName("TIPO").AsString
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Já existe um registro com o mesmo tipo !", "E")
  End If

  Set SQL = Nothing
End Sub

