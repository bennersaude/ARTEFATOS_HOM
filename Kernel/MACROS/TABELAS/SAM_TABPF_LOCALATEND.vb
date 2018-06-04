'HASH: 7174D2F42351D22FC90E0C05B7AEC732
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_TABPF_LOCALATEND             ")
Consulta.Add(" WHERE TABPFEVENTO = :TABPFEVENTO       ")
Consulta.Add("   AND LOCALATENDIMENTO = :LOCAL        ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("TABPFEVENTO").AsInteger = CurrentQuery.FieldByName("TABPFEVENTO").AsInteger
Consulta.ParamByName("LOCAL").AsInteger = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Local de atendimento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
