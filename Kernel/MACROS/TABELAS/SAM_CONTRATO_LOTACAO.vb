'HASH: 04F13A1DFC9DD9BC60B6D085F7B75139
'Macro da tabela: SAM_CONTRATO_LOTACAO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT COUNT(1) QTDREGISTROS")
  SQL.Add("  FROM SAM_CONTRATO_LOTACAO ")
  SQL.Add(" WHERE CODIGO   = :CODIGO   ")
  SQL.Add("   AND CONTRATO = :CONTRATO ")
  SQL.Add("   AND HANDLE  <> :HANDLE   ")
  SQL.ParamByName("CODIGO" ).AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  SQL.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  SQL.ParamByName("HANDLE" ).AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If (SQL.FieldByName("QTDREGISTROS").AsInteger > 0) Then
    bsShowMessage("Já existe uma lotação cadastrada no contrato com este código.", "E")
    CanContinue = False
    SQL.Active = False
    Set SQL = Nothing
    Exit Sub
  End If

  SQL.Active = False
  Set SQL = Nothing
End Sub

