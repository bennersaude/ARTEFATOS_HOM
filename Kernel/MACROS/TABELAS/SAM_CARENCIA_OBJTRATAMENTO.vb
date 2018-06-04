'HASH: E033AB5940BFD7DB95226460CF512417
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                               ")
Consulta.Add("  FROM SAM_CARENCIA_OBJTRATAMENTO      ")
Consulta.Add(" WHERE CARENCIA = :CARENCIA            ")
Consulta.Add("   AND OBJETIVOTRATAMENTO = :OBJETIVO  ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE               ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("CARENCIA").AsInteger = CurrentQuery.FieldByName("CARENCIA").AsInteger
Consulta.ParamByName("OBJETIVO").AsInteger = CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Objetivo já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If

End Sub
