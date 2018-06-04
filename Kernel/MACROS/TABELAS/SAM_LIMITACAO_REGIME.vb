'HASH: FE931FF01E878350277F8587371F3E43
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_LIMITACAO_REGIME             ")
Consulta.Add(" WHERE LIMITACAO = :LIMITACAO           ")
Consulta.Add("   AND REGIMEATENDIMENTO = :REGIME      ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("LIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
Consulta.ParamByName("REGIME").AsInteger = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Regime de atendimento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
