'HASH: 13727A53A4108E8BB5945F0EFEC08E3E
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Active = False
Consulta.Clear
Consulta.Add("SELECT *                           ")
Consulta.Add("  FROM SAM_CARENCIA_FINATENDIMENTO ")
Consulta.Add(" WHERE FINALIDADEATENDIMENTO = :ATE")
Consulta.Add("   AND CARENCIA = :CARENCIA        ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If

Consulta.ParamByName("CARENCIA").AsInteger = CurrentQuery.FieldByName("CARENCIA").AsInteger
Consulta.ParamByName("ATE").AsInteger = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Finalidade de atendimento ja cadastrada!", "E")
  CanContinue = False
  Exit Sub
End If

End Sub
