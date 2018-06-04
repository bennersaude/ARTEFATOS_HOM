'HASH: C46895543424969011B41EA67EFA899E
 
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean) 
  If (CommandID = "EXECUTE") Then 
    Dim Q As BPesquisa 
    Set Q = NewQuery 
      Q.Add("UPDATE Z_WFTAREFAITENS SET EXECUTOR = :USUARIO, TERMINOUSER=:NOW WHERE EXECUTOR IS NULL AND HANDLE=:HANDLE") 
      Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
      Q.ParamByName("USUARIO").AsInteger = CurrentUser 
      Q.ParamByName("NOW").AsDateTime = ServerNow 
 
      StartTransaction 
      Q.ExecSQL 
      Commit 
      InfoDescription = "Execução do item da tarefa registrado com sucesso!" 
  End If 
End Sub 
