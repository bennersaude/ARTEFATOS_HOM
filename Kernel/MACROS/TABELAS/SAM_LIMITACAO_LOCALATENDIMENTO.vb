'HASH: A2C308AE6D842F66AD6309066342F2F3
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_LIMITACAO_LOCALATENDIMENTO   ")
Consulta.Add(" WHERE LIMITACAO = :LIMITACAO           ")
Consulta.Add("   AND LOCALATENDIMENTO = :LOCAL        ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("LIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
Consulta.ParamByName("LOCAL").AsInteger = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Local de atendimento já cadastrada!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
