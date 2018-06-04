'HASH: AFE09227ED1BC1F92563BBA1A085F720
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                            ")
Consulta.Add("  FROM SAM_CARENCIA_TIPOTRATAMENTO  ")
Consulta.Add(" WHERE CARENCIA = :CARENCIA         ")
Consulta.Add("   AND TIPOTRATAMENTO = :TRATAMENTO ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE            ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("CARENCIA").AsInteger = CurrentQuery.FieldByName("CARENCIA").AsInteger
Consulta.ParamByName("TRATAMENTO").AsInteger = CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Tipo de Tratamento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
