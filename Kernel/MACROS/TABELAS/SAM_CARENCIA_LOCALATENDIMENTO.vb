'HASH: 231BDBC10B3D00FEC55BB9702AE1D819
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Active = False
Consulta.Clear
Consulta.Add("SELECT *                            ")
Consulta.Add("  FROM SAM_CARENCIA_LOCALATENDIMENTO")
Consulta.Add(" WHERE LOCALATENDIMENTO = :LOCAL    ")
Consulta.Add("   AND CARENCIA = :CARENCIA         ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE   ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("CARENCIA").AsInteger = CurrentQuery.FieldByName("CARENCIA").AsInteger
Consulta.ParamByName("LOCAL").AsInteger = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
Consulta.Active = True

If Consulta.FieldByName("HANDLE").AsInteger > 0 Then
  bsShowMessage("Local de Atendimento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
