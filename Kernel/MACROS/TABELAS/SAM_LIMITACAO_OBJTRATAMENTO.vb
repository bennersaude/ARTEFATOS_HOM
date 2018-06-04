'HASH: DC4AE47D4D58973855400B4288478EB3
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_LIMITACAO_OBJTRATAMENTO      ")
Consulta.Add(" WHERE LIMITACAO = :LIMITACAO           ")
Consulta.Add("   AND OBJETIVOTRATAMENTO = :OBJETIVO   ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("LIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
Consulta.ParamByName("OBJETIVO").AsInteger = CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Objetivo de Tratamento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
