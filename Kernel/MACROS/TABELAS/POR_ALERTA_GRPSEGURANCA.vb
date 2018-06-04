'HASH: 9139E7E471E0774FC47E593AF1441408
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLGRPSEG As Object
  Set SQLGRPSEG = NewQuery

  SQLGRPSEG.Active = False
  SQLGRPSEG.Clear
  SQLGRPSEG.Add("SELECT COUNT(1) QTD")
  SQLGRPSEG.Add("  FROM POR_ALERTA_GRPSEGURANCA")
  SQLGRPSEG.Add(" WHERE GRPSEGURANCA = :GRP AND ALERTA = :ALERTA")
  SQLGRPSEG.Add("   AND HANDLE <> :HANDLE")
  SQLGRPSEG.ParamByName("GRP").AsInteger = CurrentQuery.FieldByName("GRPSEGURANCA").AsInteger
  SQLGRPSEG.ParamByName("ALERTA").AsInteger = CurrentQuery.FieldByName("ALERTA").AsInteger
  SQLGRPSEG.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLGRPSEG.Active = True

  If SQLGRPSEG.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para este grupo de segurança. Selecione outro grupo de segurança!","I")
	Set SQLGRPSEG = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLGRPPREST = Nothing

End Sub
