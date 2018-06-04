'HASH: 482244ED9947F915530DEFCE9F2AC38A
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_LIMITACAO_FINATENDIMENTO     ")
Consulta.Add(" WHERE LIMITACAO = :LIMITACAO           ")
Consulta.Add("   AND FINALIDADEATENDIMENTO = :FINALI  ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("LIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
Consulta.ParamByName("FINALI").AsInteger = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Finalidade de atendimento já cadastrada!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
