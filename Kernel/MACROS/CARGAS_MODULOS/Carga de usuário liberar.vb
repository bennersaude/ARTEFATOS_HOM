'HASH: E28608EA364192F852E388DC199BC3FA
 

Public Sub LIBERARNEGACOES_OnClick()
  Dim SQL1 As Object
  Dim SQL2 As Object

If NodeInternalCode=1 Then 
	  If MsgBox("Confirma a liberação de todas as negações para o usuário ?" ,vbYesNo,"Liberação de Negações") = vbYes Then
	    Set SQL1 = NewQuery
	    Set SQL2 = NewQuery
	
	
	    SQL1.Clear
	    SQL1.Add("SELECT HANDLE FROM SAM_MOTIVONEGACAO")
	    SQL1.Add("WHERE HANDLE NOT IN")
	    SQL1.Add("(SELECT MOTIVONEGACAO FROM SAM_USUARIO_MOTIVONEGACAO WHERE USUARIO = :USUARIO)")
	    SQL1.ParamByName("USUARIO").Value = RecordHandleOfTable("Z_GRUPOUSUARIOS")
	    SQL1.Active=True
	
	    While Not SQL1.EOF
	      SQL2.Clear
	      SQL2.Add("INSERT INTO SAM_USUARIO_MOTIVONEGACAO (HANDLE, USUARIO, MOTIVONEGACAO) VALUES (:HANDLE,:USUARIO,:MOTIVONEGACAO)")
	      SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVONEGACAO")
	      SQL2.ParamByName("USUARIO").Value = RecordHandleOfTable("Z_GRUPOUSUARIOS")
	      SQL2.ParamByName("MOTIVONEGACAO").Value = SQL1.FieldByName("HANDLE").AsInteger
	      SQL2.ExecSQL
	      SQL1.Next 
	    Wend
	    RefreshNodesWithTable"SAM_USUARIO_MOTIVONEGACAO" 
	    
	  End If
    
  End If
  
  
  If NodeInternalCode=2 Then
	  If MsgBox("Confirma a liberação de todas as glosas para o usuário ?" ,vbYesNo,"Liberação de Glosas") = vbYes Then
	    Set SQL1 = NewQuery
	    Set SQL2 = NewQuery


	    SQL1.Clear
	    SQL1.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA")
	    SQL1.Add("WHERE HANDLE NOT IN")
	    SQL1.Add("(SELECT MOTIVOGLOSA FROM SAM_USUARIO_MOTIVOGLOSA WHERE USUARIO = :USUARIO)")
	    SQL1.ParamByName("USUARIO").Value = RecordHandleOfTable("Z_GRUPOUSUARIOS")
	    SQL1.Active=True

	    While Not SQL1.EOF
	      SQL2.Clear
	      SQL2.Add("INSERT INTO SAM_USUARIO_MOTIVOGLOSA (HANDLE, USUARIO, MOTIVOGLOSA) VALUES (:HANDLE,:USUARIO,:MOTIVOGLOSA)")
	      SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVOGLOSA")
	      SQL2.ParamByName("USUARIO").Value = RecordHandleOfTable("Z_GRUPOUSUARIOS")
	      SQL2.ParamByName("MOTIVOGLOSA").Value = SQL1.FieldByName("HANDLE").AsInteger
	      SQL2.ExecSQL
	      SQL1.Next
	    Wend
	    RefreshNodesWithTable"SAM_USUARIO_MOTIVOGLOSA"

	  End If
  End If

  Set SQL1 = Nothing
  Set SQL2 = Nothing
End Sub

