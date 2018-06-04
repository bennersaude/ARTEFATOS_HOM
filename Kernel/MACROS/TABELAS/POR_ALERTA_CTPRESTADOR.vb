'HASH: 6564EC20ABF9ECBE6D295BF031054D25
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLCTPREST As Object
  Set SQLCTPREST = NewQuery

  SQLCTPREST.Active = False
  SQLCTPREST.Clear
  SQLCTPREST.Add("SELECT COUNT(1) QTD")
  SQLCTPREST.Add("  FROM POR_ALERTA_CTPRESTADOR")
  SQLCTPREST.Add(" WHERE CATEGORIAPRESTADOR = :CAT AND ALERTA = :ALERTA")
  SQLCTPREST.Add("   AND HANDLE <> :HANDLE")
  SQLCTPREST.ParamByName("CAT").AsInteger = CurrentQuery.FieldByName("CATEGORIAPRESTADOR").AsInteger
  SQLCTPREST.ParamByName("ALERTA").AsInteger = CurrentQuery.FieldByName("ALERTA").AsInteger
  SQLCTPREST.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLCTPREST.Active = True

  If SQLCTPREST.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para esta categoria de prestador. Selecione outra categoria!","I")
	Set SQLCTPREST = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLCTPREST = Nothing

End Sub
