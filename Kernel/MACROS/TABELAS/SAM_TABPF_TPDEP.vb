'HASH: 810D3771ADB861F1CFF047F9A90A5B7C
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_TABPF_TPDEP                  ")
Consulta.Add(" WHERE TABPFEVENTO = :TABPFEVENTO       ")
Consulta.Add("   AND TIPODEPENDENTE = :DP             ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("TABPFEVENTO").AsInteger = CurrentQuery.FieldByName("TABPFEVENTO").AsInteger
Consulta.ParamByName("DP").AsInteger = CurrentQuery.FieldByName("TIPODEPENDENTE").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Tipo Dependente já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
