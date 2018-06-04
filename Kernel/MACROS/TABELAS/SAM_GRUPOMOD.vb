'HASH: 3783568E67A0FD2FB19C2B7D9167027D
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT * FROM SAM_GRUPOMOD WHERE CODIGO = :CODIGO AND HANDLE <> :HANDLE")
  SQL.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Código já existente.", "E")
    CanContinue = False
  End If

  SQL.Active = False
  Set SQL = Nothing

End Sub

