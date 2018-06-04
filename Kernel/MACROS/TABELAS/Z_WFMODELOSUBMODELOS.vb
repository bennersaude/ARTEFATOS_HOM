'HASH: D922B46A4B118495C04A66AE205F533C
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  Dim q As BPesquisa 
  Set q = NewQuery 
 
  q.Add("SELECT COUNT(1) TOTAL FROM Z_WFMODELOSUBMODELOS A, Z_WFMODELOS B WHERE B.NOME=(SELECT NOME FROM Z_WFMODELOS WHERE HANDLE=:SUBMODELO) AND B.HANDLE=A.SUBMODELO AND A.MODELO=:MODELO") 
  q.ParamByName("SUBMODELO").AsInteger = CurrentQuery.FieldByName("SUBMODELO").AsInteger 
  q.ParamByName("MODELO").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger 
  q.Active = True 
  If Not q.EOF Then 
    If (q.FieldByName("TOTAL").AsInteger > 0) Then 
    	Err.Raise(vbsUserException, "","Não é permitido cadastrar duas versões de um mesmo fluxo. Remova as associações de outras versões do mesmo fluxo antes de poder continuar.") 
    End If 
  End If 
  q.Active = False 
  Set q = Nothing 
End Sub 
