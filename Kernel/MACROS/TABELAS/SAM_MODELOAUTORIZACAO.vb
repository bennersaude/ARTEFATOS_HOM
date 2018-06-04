'HASH: 3A874EA67D8EDFA7A9E2483B76279983
'#uses "*bsShowMessage"

Public Sub BOTAONOVAINTERFACE_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Registro em edição", "E")
    Exit Sub
  End If

  Dim interface As Object
  Set interface = CreateBennerObject("CA044.Rotinas")
  interface.MontarLeiaute(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSQL As Object
  Set qSQL = NewQuery

  qSQL.Add("SELECT HANDLE")
  qSQL.Add("  FROM SAM_MODELOAUTORIZACAO")
  qSQL.Add(" WHERE HANDLE <> :HANDLE")
  qSQL.Add("   AND DESCRICAO = :DESCRICAO")
  qSQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSQL.ParamByName("DESCRICAO").AsString = CurrentQuery.FieldByName("DESCRICAO").AsString
  qSQL.Active = True

  If qSQL.FieldByName("HANDLE").AsInteger > 0 Then
    bsShowMessage("Modelo de autorização já existente com essa descrição, tente outra.", "E")
    CanContinue = False
  End If

  Set qSQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAONOVAINTERFACE"
			BOTAONOVAINTERFACE_OnClick
	End Select
End Sub
