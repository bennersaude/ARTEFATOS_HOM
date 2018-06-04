'HASH: 206D3346687317A1DC2D0E0DD5FD2FD8


Public Sub LIBERARGLOSAS_OnClick()
  If MsgBox("Confirma a liberação de todos as glosas para o grupo ?" ,vbYesNo,"Liberação de Negações") = vbYes Then
	Set SQL1 = NewQuery
	Set SQL2 = NewQuery
	Set SQLUSUARIO = NewQuery

    SQL1.Clear
	SQL1.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA")
	SQL1.Add("WHERE HANDLE NOT IN")
	SQL1.Add("(SELECT MOTIVOGLOSA FROM SAM_GRUPO_MOTIVOGLOSA WHERE GRUPO = :GRUPO)")
	SQL1.ParamByName("GRUPO").Value = RecordHandleOfTable("Z_GRUPOS")
	SQL1.Active=True

	While Not SQL1.EOF
	  SQL2.Clear
	  SQL2.Add("INSERT INTO SAM_GRUPO_MOTIVOGLOSA (HANDLE, GRUPO, MOTIVOGLOSA) VALUES (:HANDLE,:GRUPO,:MOTIVOGLOSA)")
	  SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_GRUPO_MOTIVOGLOSA")
	  SQL2.ParamByName("GRUPO").Value = RecordHandleOfTable("Z_GRUPOS")
	  SQL2.ParamByName("MOTIVOGLOSA").Value = SQL1.FieldByName("HANDLE").AsInteger
	  SQL2.ExecSQL
	  SQL1.Next
	Wend
	RefreshNodesWithTable"SAM_GRUPO_MOTIVOGLOSA"


  SQLUSUARIO.Clear
  SQLUSUARIO.Add("SELECT HANDLE FROM Z_GRUPOUSUARIOS WHERE GRUPO = :GRUPO")
  SQLUSUARIO.ParamByName("GRUPO").Value = RecordHandleOfTable("Z_GRUPOS")
  SQLUSUARIO.Active = True

  While Not SQLUSUARIO.EOF
    Set SQL1 = NewQuery
	Set SQL2 = NewQuery


    SQL1.Clear
    SQL1.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA")
    SQL1.Add("WHERE HANDLE NOT IN")
    SQL1.Add("(SELECT MOTIVOGLOSA FROM SAM_USUARIO_MOTIVOGLOSA WHERE USUARIO = :USUARIO)")
    SQL1.Add("AND HANDLE IN (SELECT MOTIVOGLOSA FROM SAM_GRUPO_MOTIVOGLOSA WHERE GRUPO = :GRUPO)")
    SQL1.ParamByName("USUARIO").Value = sqlusuario.FieldByName("HANDLE").AsInteger
    SQL1.ParamByName("GRUPO").Value = RecordHandleOfTable("Z_GRUPOS")
    SQL1.Active=True

    While Not SQL1.EOF
      SQL2.Clear
      SQL2.Add("INSERT INTO SAM_USUARIO_MOTIVOGLOSA (HANDLE, USUARIO, MOTIVOGLOSA) VALUES (:HANDLE,:USUARIO,:MOTIVOGLOSA)")
      SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVOGLOSA")
      SQL2.ParamByName("USUARIO").Value = sqlusuario.FieldByName("HANDLE").AsInteger
      SQL2.ParamByName("MOTIVOGLOSA").Value = SQL1.FieldByName("HANDLE").AsInteger
      SQL2.ExecSQL
      SQL1.Next
    Wend

   SQLUSUARIO.Next

  Wend

  End If




End Sub
