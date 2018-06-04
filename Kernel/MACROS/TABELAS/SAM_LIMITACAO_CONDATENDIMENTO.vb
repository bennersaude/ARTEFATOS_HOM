'HASH: 01382602A3F1D5AF849CF676C52121B6
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_LIMITACAO_CONDATENDIMENTO    ")
Consulta.Add(" WHERE LIMITACAO = :LIMITACAO           ")
Consulta.Add("   AND CONDICAOATENDIMENTO = :CONDICAO  ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("LIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
Consulta.ParamByName("CONDICAO").AsInteger = CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Condição de atendimento já cadastrada!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
