'HASH: 58263ACB0C3C9F56BBB871754A54296C
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM SAM_TABELAHE_ITEM WHERE TABELAHE = :TABELAHE AND HANDLE <> :HANDLE AND TIPODIA = :TIPODIA")
  sql.Add("AND :HORA BETWEEN HORAINICIAL AND HORAFINAL")

  sql.Active = False
  sql.ParamByName("TABELAHE").Value = CurrentQuery.FieldByName("TABELAHE").AsInteger
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ParamByName("TIPODIA").Value = CurrentQuery.FieldByName("TIPODIA").AsString
  sql.ParamByName("HORA").Value = CurrentQuery.FieldByName("HORAINICIAL").AsDateTime

  sql.Active = True
  If Not sql.EOF Then
    CanContinue = False
    bsShowMessage("Horário inicial em conflito", "E")
    Exit Sub
    Set sql = Nothing
  End If

  sql.Active = False
  sql.ParamByName("TABELAHE").Value = CurrentQuery.FieldByName("TABELAHE").AsInteger
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ParamByName("TIPODIA").Value = CurrentQuery.FieldByName("TIPODIA").AsString
  sql.ParamByName("HORA").Value = CurrentQuery.FieldByName("HORAFINAL").AsDateTime

  sql.Active = True
  If Not sql.EOF Then
    CanContinue = False
    bsShowMessage("Horário final em conflito", "E")
    Exit Sub
    Set sql = Nothing
  End If


  sql.Clear
  sql.Add("SELECT HANDLE FROM SAM_TABELAHE_ITEM WHERE TABELAHE = :TABELAHE AND HANDLE <> :HANDLE AND TIPODIA = :TIPODIA")
  sql.Add("AND HORAINICIAL BETWEEN :HORAINICIAL AND :HORAFINAL")
  sql.Active = False
  sql.ParamByName("TABELAHE").Value = CurrentQuery.FieldByName("TABELAHE").AsInteger
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ParamByName("TIPODIA").Value = CurrentQuery.FieldByName("TIPODIA").AsString
  sql.ParamByName("HORAINICIAL").Value = CurrentQuery.FieldByName("HORAINICIAL").AsDateTime
  sql.ParamByName("HORAFINAL").Value = CurrentQuery.FieldByName("HORAFINAL").AsDateTime
  sql.Active = True

  sql.Clear
  sql.Add("SELECT HANDLE FROM SAM_TABELAHE_ITEM WHERE TABELAHE = :TABELAHE AND HANDLE <> :HANDLE AND TIPODIA = :TIPODIA")
  sql.Add("AND HORAFINAL BETWEEN :HORAINICIAL AND :HORAFINAL")
  sql.Active = False
  sql.ParamByName("TABELAHE").Value = CurrentQuery.FieldByName("TABELAHE").AsInteger
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ParamByName("TIPODIA").Value = CurrentQuery.FieldByName("TIPODIA").AsString
  sql.ParamByName("HORAINICIAL").Value = CurrentQuery.FieldByName("HORAINICIAL").AsDateTime
  sql.ParamByName("HORAFINAL").Value = CurrentQuery.FieldByName("HORAFINAL").AsDateTime
  sql.Active = True

  If Not sql.EOF Then
    CanContinue = False
    bsShowMessage("Horário final em conflito", "E")
    Exit Sub
    Set sql = Nothing
  End If

  If CurrentQuery.FieldByName("HORAINICIAL").AsDateTime >= CurrentQuery.FieldByName("HORAFINAL").AsDateTime Then
    CanContinue = False
    bsShowMessage("Horário inicial não pode ser posterior ao horário final!", "E")
  End If


End Sub

