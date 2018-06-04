'HASH: 5880B4ED41022BD6AF8BFB9D1D58A9A3
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT 1                         ")
  SQL.Add("  FROM SAM_TGE_MATMEDCLINICA C   ")
  SQL.Add(" WHERE C.CLINICA = :CLINICA      ")
  SQL.Add("   AND C.EVENTO = :EVENTO        ")
  SQL.Add("   AND C.LOTE = :LOTE            ")
  SQL.Add("   AND C.HANDLE <> :HANDLE       ")

  SQL.ParamByName("CLINICA").AsInteger = CurrentQuery.FieldByName("CLINICA").AsInteger
  SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.ParamByName("LOTE").AsString = CurrentQuery.FieldByName("LOTE").AsString
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  SQL.Active = True

  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Já existe este medicamento cadastrado para este lote e clínica.", "E")
    Exit Sub
  End If
  Set SQL = Nothing

End Sub
