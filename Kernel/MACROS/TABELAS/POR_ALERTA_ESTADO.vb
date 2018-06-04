'HASH: F1319E8E9194A4B4F06277F59048B720
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLMUNIC As Object
  Set SQLMUNIC = NewQuery

  SQLMUNIC.Active = False
  SQLMUNIC.Clear
  SQLMUNIC.Add("SELECT COUNT(1) QTD")
  SQLMUNIC.Add("  FROM POR_ALERTA_ESTADO_MUN")
  SQLMUNIC.Add(" WHERE ALERTAESTADO = :HANDLE")
  SQLMUNIC.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLMUNIC.Active = True

  If SQLMUNIC.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Existem municípios registrados para este estado. Alteração não permitida.","E")
	Set SQLMUNIC = Nothing
	CanContinue = False
	Exit Sub
  End If

  SQLMUNIC.Active = False
  SQLMUNIC.Clear
  SQLMUNIC.Add("SELECT COUNT(1) QTD")
  SQLMUNIC.Add("  FROM POR_ALERTA_ESTADO")
  SQLMUNIC.Add(" WHERE ESTADO = :ESTADO AND ALERTA = :ALERTA")
  SQLMUNIC.Add("   AND HANDLE <> :HANDLE")
  SQLMUNIC.ParamByName("ESTADO").AsInteger = CurrentQuery.FieldByName("ESTADO").AsInteger
  SQLMUNIC.ParamByName("ALERTA").AsInteger = CurrentQuery.FieldByName("ALERTA").AsInteger
  SQLMUNIC.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLMUNIC.Active = True

  If SQLMUNIC.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para este Estado. Selecione outro Estado!","E")
	Set SQLMUNIC = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLMUNIC = Nothing

End Sub
