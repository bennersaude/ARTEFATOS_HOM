'HASH: E1F51B2751C11BCE6EB4D33A867B4747
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLPERFIL As Object
  Set SQLPERFIL = NewQuery

  SQLPERFIL.Active = False
  SQLPERFIL.Clear
  SQLPERFIL.Add("SELECT COUNT(1) QTD")
  SQLPERFIL.Add("  FROM POR_ALERTA_PERFILUSUARIO")
  SQLPERFIL.Add(" WHERE PERFILUSUARIO = :PU AND ALERTA = :ALERTA")
  SQLPERFIL.Add("   AND HANDLE <> :HANDLE")
  SQLPERFIL.ParamByName("PU").AsInteger = CurrentQuery.FieldByName("PERFILUSUARIO").AsInteger
  SQLPERFIL.ParamByName("ALERTA").AsInteger = CurrentQuery.FieldByName("ALERTA").AsInteger
  SQLPERFIL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLPERFIL.Active = True

  If SQLPERFIL.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para este perfil de usuário. Selecione outro perfil de usuário!","I")
	Set SQLPERFIL = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLPERFIL = Nothing

End Sub
