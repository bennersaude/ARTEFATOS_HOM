'HASH: 37EE61B49E71E45038701AB8AD12E62F

'MACRO= sam_autoriz_remocao

Public Sub BOTAORELATORIOREMOCAO_OnClick()

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'AUT037'")
  SQL.Active = True

  ReportPreview(SQL.FieldByName("HANDLE").AsInteger, "", True, True)

  Set SQL = Nothing
End Sub

