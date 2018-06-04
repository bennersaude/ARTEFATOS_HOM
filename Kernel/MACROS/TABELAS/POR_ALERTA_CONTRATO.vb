'HASH: EE5A5CE6101FF2C6E14296359FC0DA84
 
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLCONTRATO As Object
  Set SQLCONTRATO = NewQuery

  SQLCONTRATO.Active = False
  SQLCONTRATO.Clear
  SQLCONTRATO.Add("SELECT COUNT(1) QTD")
  SQLCONTRATO.Add("  FROM POR_ALERTA_CONTRATO")
  SQLCONTRATO.Add(" WHERE CONTRATO = :CT AND ALERTA = :ALERTA")
  SQLCONTRATO.Add("   AND HANDLE <> :HANDLE")
  SQLCONTRATO.ParamByName("CT").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  SQLCONTRATO.ParamByName("ALERTA").AsInteger = CurrentQuery.FieldByName("ALERTA").AsInteger
  SQLCONTRATO.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLCONTRATO.Active = True

  If SQLCONTRATO.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para este contrato. Selecione outro contrato!","I")
	Set SQLCONTRATO = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLCONTRATO = Nothing

End Sub
