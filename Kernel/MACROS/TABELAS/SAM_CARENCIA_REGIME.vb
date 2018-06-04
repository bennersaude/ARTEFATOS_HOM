'HASH: 936F688C762F3144A717FE3DE6F29ACD
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                             ")
Consulta.Add("  FROM SAM_CARENCIA_REGIME           ")
Consulta.Add(" WHERE REGIMEATENDIMENTO = :REGIME   ")
Consulta.Add("   AND CARENCIA = :CARENCIA          ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE             ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("REGIME").AsInteger = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
Consulta.ParamByName("CARENCIA").AsInteger = CurrentQuery.FieldByName("CARENCIA").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Regime já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If

End Sub
