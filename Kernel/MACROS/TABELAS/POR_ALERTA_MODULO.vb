'HASH: EEE18D534BC5E516E2F86BAF585553B3
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLMODULO As Object
  Set SQLMODULO = NewQuery

  SQLMODULO.Active = False
  SQLMODULO.Clear
  SQLMODULO.Add("SELECT COUNT(1) QTD")
  SQLMODULO.Add("  FROM POR_ALERTA_MODULO")
  SQLMODULO.Add(" WHERE MODULO = :MOD AND ALERTA = :ALERTA")
  SQLMODULO.Add("   AND HANDLE <> :HANDLE")
  SQLMODULO.ParamByName("MOD").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
  SQLMODULO.ParamByName("ALERTA").AsInteger = CurrentQuery.FieldByName("ALERTA").AsInteger
  SQLMODULO.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLMODULO.Active = True

  If SQLMODULO.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para este módulo. Selecione outro módulo!","I")
	Set SQLMODULO = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLMODULO = Nothing

End Sub
