'HASH: 50818CA417EC528F0A3E4CCDE842D83A
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_TABPF_OBJTRATAMENTO          ")
Consulta.Add(" WHERE TABPFEVENTO = :TABPFEVENTO       ")
Consulta.Add("   AND OBJETIVOTRATAMENTO = :OBJETIVO   ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("TABPFEVENTO").AsInteger = CurrentQuery.FieldByName("TABPFEVENTO").AsInteger
Consulta.ParamByName("OBJETIVO").AsInteger = CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Objetivo de tratamento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
