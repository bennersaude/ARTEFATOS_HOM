'HASH: 99F61B13600B7F8205AF3EB67A48C610
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
 
 	If Not CurrentQuery.FieldByName("CODIGO").IsNull Then 
 
		Dim Q As BPesquisa 
		Dim JaExiste As Boolean 
 
		Set Q = NewQuery 
		Q.Add("SELECT Z_CODIGO FROM Z_MACROS WHERE Z_CODIGO = :ZCODIGO") 
 
		If CurrentQuery.State = 2 Then 
			Q.Add("AND HANDLE <> :HANDLE") 
			Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
		End If 
 
		Q.ParamByName("ZCODIGO").AsString = TiraAcento(CurrentQuery.FieldByName("CODIGO").AsString, True) 
 
		Q.Active = True 
		JaExiste = Not Q.EOF 
		Q.Active = False 
 
		If JaExiste Then 
			Err.Raise(vbsUserException, , "Já existe uma macro com este código!") 
		End If 
 
	End If 
End Sub 
 
Public Sub TABLE_UpdateRequired() 
  If NodeInternalCode = 255 Then 
    CurrentQuery.FieldByName("TIPO").AsInteger = 5 
  End If 
End Sub 
 
