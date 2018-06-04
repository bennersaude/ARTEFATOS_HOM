'HASH: BB26D42D08F6E4F2DEB485727D7FA1FE
Option Explicit 
 
Dim LastLevelBefore As Boolean 
 
Public Sub TABLE_AfterInsert() 
  LastLevelBefore = False 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  LastLevelBefore = CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "S" 
End Sub 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  Dim q As Object 
  If (CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "N") And (LastLevelBefore) Then 
    Set q = NewQuery 
    q.Add("SELECT COUNT(HANDLE) NRECS FROM R_RELATORIOIMPRESSOES WHERE CLASSIFICACAO = :CLASSIFICACAO") 
    q.ParamByName("CLASSIFICACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
    q.Active = True 
    If q.FieldByName("NRECS").AsInteger > 0 Then 
      MsgBox "Registro deve ser último nível pois já existem impressos cadastrados." 
      CanContinue = False 
    End If 
    Set q = Nothing 
  End If 
End Sub 
