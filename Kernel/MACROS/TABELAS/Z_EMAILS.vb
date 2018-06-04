'HASH: 75E419359A918AB0516F60D087D50A3B
 
Option Explicit 
Dim StrFrom As String 
 
Sub NovoStatus(st As Integer ) 
Dim q As Object 
  Set q = NewQuery 
  q.Add("UPDATE Z_EMAILS SET STATUS = :STATUS, DATASTATUS = :DATASTATUS WHERE HANDLE = :HANDLE") 
  q.ParamByName("STATUS").AsInteger = st 
  q.ParamByName("DATASTATUS").AsDateTime = ServerNow 
  q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
  StartTransaction 
  q.ExecSQL 
  Commit 
  Set q = Nothing 
  RefreshNodesWithTable("Z_EMAILS")  
End Sub 
 
Public Sub BOTAOALTERAR_OnClick() 
  If CurrentQuery.State <> 1 Then 
    MsgBox "O registro está em edição!" 
  Else 
    NovoStatus 1 
  End If 
End Sub 
 
Public Sub BOTAOCANCELARENVIO_OnClick() 
  If MsgBox("Confirma cancelamento do envio do e-mail?" , vbYesNo) = vbYes Then 
    NovoStatus 5 
    RefreshNodesWithTable("Z_EMAILS") 
  End If 
End Sub 
 
Public Sub BOTAOENVIAR_OnClick() 
Dim m As Object 
  Set m = CreateBennerObject("CSCOMMON.ScheduledMails") 
  On Error GoTo ErroEnvio 
  m.SendWhere(CurrentSystem, "HANDLE = "+ CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)) 
  GoTo Sair 
ErroEnvio: 
 MsgBox Err.Description 
Sair: 
  Set m = Nothing 
  RefreshNodesWithTable("Z_EMAILS") 
End Sub 
 
Public Sub BOTAOLIBERAR_OnClick() 
  If CurrentQuery.State <> 1 Then 
    CurrentQuery.FieldByName("STATUS").AsInteger = 2 
  Else 
    NovoStatus 2 
  End If 
End Sub 
 
Public Sub TABLE_AfterInsert() 
  CurrentQuery.FieldByName("DE").AsString = StrFrom 
End Sub 
 
Public Sub TABLE_BeforeDelete(CanContinue As Boolean) 
Dim q As Object 
  Set q = NewQuery 
  q.Add("SELECT HANDLE FROM Z_EMAILANEXOS WHERE EMAIL = " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)) 
  q.Active = True 
  Do While Not q.EOF 
    ClearFieldDocument("Z_EMAILANEXOS", "ANEXO", q.FieldByName("HANDLE").AsInteger, True) 
    q.Next 
  Loop 
  q.Active = False 
  q.Clear 
  q.Add("DELETE FROM Z_EMAILANEXOS WHERE EMAIL = " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)) 
  q.ExecSQL 
  Set q = Nothing 
End Sub 
 
Public Sub TABLE_BeforeInsert(CanContinue As Boolean) 
Dim q As Object 
Set q = NewQuery 
  q.Add("SELECT NOME, EMAIL FROM Z_GRUPOUSUARIOS WHERE HANDLE = " + CStr(CurrentUser)) 
  q.Active = True 
  CanContinue = Trim(q.FieldByName("EMAIL").AsString) <> "" 
  StrFrom =  q.FieldByName("NOME").AsString + "<" +  q.FieldByName("EMAIL").AsString + ">" 
  q.Active = False 
  If Not CanContinue Then MsgBox "Usuário sem permissão para envio de e-mails." 
Set q = Nothing 
End Sub 
 
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean) 
  If CommandID = "EnviarEmail" Then 
	BOTAOENVIAR_OnClick 
  End If 
End Sub 
