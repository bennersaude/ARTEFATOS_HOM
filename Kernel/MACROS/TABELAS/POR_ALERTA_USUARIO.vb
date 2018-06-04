'HASH: BE288A3DFF809CE464079C17BF175A40
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
  NOMEUSUARIO.Visible = False
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQLUSR As Object
  Set SQLUSR = NewQuery

  SQLUSR.Active = False
  SQLUSR.Clear
  SQLUSR.Add("SELECT Z.NOME")
  SQLUSR.Add("  FROM Z_GRUPOUSUARIOS Z")
  SQLUSR.Add("  JOIN POR_USUARIO PU ON PU.USERZGRUPOUSUARIO = Z.HANDLE ")
  SQLUSR.Add(" WHERE PU.HANDLE = :USR")
  SQLUSR.ParamByName("USR").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger
  SQLUSR.Active = True

  If Not SQLUSR.EOF Then
    NOMEUSUARIO.Text = SQLUSR.FieldByName("NOME").AsString
    NOMEUSUARIO.Visible = True
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLUSR As Object
  Set SQLUSR = NewQuery

  SQLUSR.Active = False
  SQLUSR.Clear
  SQLUSR.Add("SELECT COUNT(1) QTD")
  SQLUSR.Add("  FROM POR_ALERTA_USUARIO")
  SQLUSR.Add(" WHERE USUARIO = :USUARIO AND ALERTA = :ALERTA")
  SQLUSR.Add("   AND HANDLE <> :HANDLE")
  SQLUSR.ParamByName("USUARIO").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger
  SQLUSR.ParamByName("ALERTA").AsInteger = CurrentQuery.FieldByName("ALERTA").AsInteger
  SQLUSR.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLUSR.Active = True

  If SQLUSR.FieldByName("QTD").AsInteger > 0 Then
	bsShowMessage("Já existe um registro para este usuário. Selecione outro usuário!","I")
	Set SQLUSR = Nothing
	CanContinue = False
	Exit Sub
  End If

  Set SQLUSR = Nothing

End Sub

