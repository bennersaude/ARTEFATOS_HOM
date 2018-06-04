'HASH: 057DBF98DBC5BF6BF5F4441F62288F16


Public Sub BOTAOGERAR_OnClick()

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim sql As Object

  Set sql = NewQuery

  sql.Clear
  sql.Add("SELECT COUNT(1) QNT          ")
  sql.Add("  FROM CLI_USUARIOMARCACAO M ")
  sql.Add(" WHERE M.USUARIO = :USUARIO  ")
  sql.Add("   AND M.RECURSO = :RECURSO  ")
  sql.ParamByName("USUARIO").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger
  sql.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
  sql.Active = True

  If (sql.FieldByName("QNT").AsInteger > 0) And (CurrentQuery.State = 3) Then
	CanContinue = False
	CancelDescription = "Já existe um vínculo entre o administrador e a agenda informada!"
  End If

  If CurrentQuery.FieldByName("AGENDAPADRAO").AsString = "S" Then
  	If ExisteOutraAgendaPadrao(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("USUARIO").AsInteger) Then
	    CanContinue = False
	    CancelDescription = "Já existe um padrão definido para este usuário! Favor verificar!"
	    Exit Sub
  	End If

  End If

  Set sql = Nothing
End Sub

Public Function ExisteOutraAgendaPadrao(handleAgenda As Long, handleUsuario As Long) As Boolean
  Dim sql As Object

  Set sql = NewQuery

  sql.Clear
  sql.Add("SELECT M.HANDLE              ")
  sql.Add("  FROM CLI_USUARIOMARCACAO M ")
  sql.Add(" WHERE M.USUARIO = :USUARIO  ")
  sql.Add("   AND M.HANDLE <> :HANDLE   ")
  sql.Add("   AND M.AGENDAPADRAO = 'S'  ")
  sql.Add("   AND M.STATUS = 'A'        ")
  sql.ParamByName("USUARIO").AsInteger = handleUsuario
  sql.ParamByName("HANDLE").AsInteger = handleAgenda
  sql.Active = True

  ExisteOutraAgendaPadrao = Not sql.FieldByName("HANDLE").IsNull

  Set sql = Nothing
End Function
