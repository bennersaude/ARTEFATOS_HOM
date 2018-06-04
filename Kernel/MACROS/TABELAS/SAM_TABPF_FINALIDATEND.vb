'HASH: 8E8D5C50332140862A70C8635E92767B
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_TABPF_FINALIDATEND           ")
Consulta.Add(" WHERE TABPFEVENTO = :TABPFEVENTO       ")
Consulta.Add("   AND FINALIDADEATENDIMENTO = :FINAL   ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("TABPFEVENTO").AsInteger = CurrentQuery.FieldByName("TABPFEVENTO").AsInteger
Consulta.ParamByName("FINAL").AsInteger = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Finalidade de atendimento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
