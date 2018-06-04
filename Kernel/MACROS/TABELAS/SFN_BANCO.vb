'HASH: B0249FFB31072E1468A3CBBF1B2C2FE7
'Macro: SFN_BANCO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As  Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                ")
Consulta.Add("  FROM SFN_BANCO        ")
Consulta.Add(" WHERE CODIGO = :CODIGO ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
Consulta.Active = True

If Consulta.FieldByName("HANDLE").AsInteger > 0 Then
  bsShowMessage("Código já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
