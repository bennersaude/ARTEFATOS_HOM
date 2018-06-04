'HASH: 17A789EA1B68F8A7C2FC527B938B62B6
'#Uses "*bsShowMessage"

Public Sub MUNICIPIO_OnPopup(ShowPopup As Boolean)

  If WebMode Then
	MUNICIPIO.WebLocalWhere = "ESTADO = (SELECT ESTADO FROM POR_ALERTA_ESTADO WHERE HANDLE =" + CStr(CurrentQuery.FieldByName("ALERTAESTADO").AsInteger) + ")"
  Else
	MUNICIPIO.LocalWhere = "ESTADO = (SELECT ESTADO FROM POR_ALERTA_ESTADO WHERE HANDLE =" + CStr(CurrentQuery.FieldByName("ALERTAESTADO").AsInteger) + ")"
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLMUNIC As Object
  Set SQLMUNIC = NewQuery

  SQLMUNIC.Active = False
  SQLMUNIC.Clear
  SQLMUNIC.Add("SELECT COUNT(1) QTD")
  SQLMUNIC.Add("  FROM POR_ALERTA_ESTADO_MUN")
  SQLMUNIC.Add(" WHERE MUNICIPIO = :MUNICIPIO AND ALERTAESTADO = :ALERTAESTADO")
  SQLMUNIC.Add("   AND HANDLE <> :HANDLE")
  SQLMUNIC.ParamByName("MUNICIPIO").AsInteger = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
  SQLMUNIC.ParamByName("ALERTAESTADO").AsInteger = CurrentQuery.FieldByName("ALERTAESTADO").AsInteger
  SQLMUNIC.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLMUNIC.Active = True

  If SQLMUNIC.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para este município neste Estado. Selecione outro município!","I")
	Set SQLMUNIC = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLMUNIC = Nothing
End Sub
