'HASH: 452A8B1311B19C45A856CA70BAB33255
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLTPPREST As Object
  Set SQLTPPREST = NewQuery

  SQLTPPREST.Active = False
  SQLTPPREST.Clear
  SQLTPPREST.Add("SELECT COUNT(1) QTD")
  SQLTPPREST.Add("  FROM POR_ALERTA_TPPRESTADOR")
  SQLTPPREST.Add(" WHERE TIPOPRESTADOR = :TPPREST AND ALERTA = :ALERTA")
  SQLTPPREST.Add("   AND HANDLE <> :HANDLE")
  SQLTPPREST.ParamByName("TPPREST").AsInteger = CurrentQuery.FieldByName("TIPOPRESTADOR").AsInteger
  SQLTPPREST.ParamByName("ALERTA").AsInteger = CurrentQuery.FieldByName("ALERTA").AsInteger
  SQLTPPREST.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLTPPREST.Active = True

  If SQLTPPREST.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para este tipo de prestador. Selecione outro tipo de prestador!","I")
	Set SQLTPPREST = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLTPPREST = Nothing

End Sub
