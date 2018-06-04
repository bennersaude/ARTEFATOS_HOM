'HASH: D51410A1860D594468A8FF08E1CF396A
Dim UsuarioDelegar As Long 
 
Public Sub BOTAOASSUMIR_OnClick() 
  Dim CanContinue As Boolean 
  CanContinue = True 
  TABLE_OnCommandClick "ASSUMIR", CanContinue 
End Sub 
 
Public Sub BOTAODELEGAR_OnClick() 
  Dim CanContinue As Boolean  
  Dim frm As Object 
  Set frm = NewVirtualForm 
  frm.TableName = "Z_WFTAREFADELEGAR" 
  frm.Caption = "Delegar tarefa" 
  frm.Physical = False 
  frm.Width = 300 
  frm.Height = 200 
  SessionVar("WFUSERDELEGATE") = "Z_GRUPOUSUARIOS.HANDLE=0" 
  If (SessionVar("WORKFLOWSELECTION") = "3" And SessionVar("WORKFLOWPOOL") <> "0") Then ' Item da fila de Coordenacao selecionado 
    SessionVar("WFUSERDELEGATE") = "Z_GRUPOUSUARIOS.HANDLE IN(SELECT USUARIO FROM Z_WFPAPELUSUARIOS WHERE PAPEL = "+ SessionVar("WORKFLOWPOOL") + ")" 
 
  End If 
  If frm.Show = 0 Then 
    CanContinue = True 
    UsuarioDelegar = frm.FormQuery.FieldByName("DESTINATARIO").AsInteger 
    TABLE_OnCommandClick "DELEGAR", CanContinue 
  End If 
  SessionVar("WFUSERDELEGATE") = "" 
  Set frm = Nothing 
End Sub 
 
Public Sub BOTAOREATIVAR_OnClick() 
  Dim CanContinue As Boolean 
  CanContinue = True 
  TABLE_OnCommandClick "UNSUSPEND", CanContinue 
End Sub 
 
Public Sub BOTAOSUSPENDER_OnClick() 
  Dim CanContinue As Boolean 
  CanContinue = True 
  TABLE_OnCommandClick "SUSPEND", CanContinue 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  'WORKFLOW SELECTION 
  '0-Minhas Tarefas 
  '1-Realizadas por mim 
  '2-Suspensas 
  '3-Coordenação 
  BOTAOASSUMIR.Visible = False 
  BOTAODELEGAR.Visible = False 
  BOTAOREATIVAR.Visible = False 
  BOTAOSUSPENDER.Visible = False 
 
  If (SessionVar("WORKFLOWVISIBLE") = "S") Then 
    Dim Q As BPesquisa 
    If (SessionVar("WORKFLOWSELECTION") = "0" Or SessionVar("WORKFLOWSELECTION") = "3") Then 
      BOTAOASSUMIR.Visible = True 
      If (CurrentQuery.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser) Then 
        BOTAOASSUMIR.Caption = "Assumir" 
        BOTAOASSUMIR.Enabled = True 
      ElseIf (CurrentQuery.FieldByName("RESPONSAVEL").AsInteger = CurrentUser) Then 
        BOTAOASSUMIR.Caption = "Devolver" 
        BOTAOASSUMIR.Enabled = True 
      Else 
        BOTAOASSUMIR.Enabled = False 
      End If 
    End If 
 
    If (SessionVar("WORKFLOWSELECTION") = "0" Or SessionVar("WORKFLOWSELECTION") = "2") Then 
      Set Q = NewQuery 
      'Q.Add("(SELECT COUNT( * ) FROM Z_WFMODELOSUSPENDERGRUPOS WS, Z_GRUPOUSUARIOS WG WHERE WG.HANDLE=:USUARIO AND WS.MODELO = (SELECT WI.MODELO FROM Z_WFMODELOINSTANCIAS WI WHERE WI.HANDLE=:MODELOINSTANCIA) AND WS.GRUPO = WG.GRUPO) > 0") 
      Q.Add("SELECT * FROM (SELECT COUNT(WS.HANDLE) TOTAL FROM Z_WFMODELOSUSPENDERGRUPOS WS, Z_GRUPOUSUARIOS WG WHERE WG.HANDLE=:USUARIO AND WS.MODELO=(SELECT WI.MODELO FROM Z_WFMODELOINSTANCIAS WI WHERE WI.HANDLE=:MODELOINSTANCIA) AND WS.GRUPO = WG.GRUPO) TABELA WHERE TOTAL > 0") 
      Q.ParamByName("USUARIO").AsInteger = CurrentUser 
      Q.ParamByName("MODELOINSTANCIA").AsInteger = CurrentQuery.FieldByName("MODELOINSTANCIA").AsInteger 
      Q.Active = True 
      If (Not Q.EOF) Then 
        Dim qry As BPesquisa 
        Set qry = NewQuery 
        qry.Add("SELECT SITUACAO FROM Z_WFMODELOINSTANCIAS WHERE HANDLE=:MODELOINSTANCIA") 
        qry.ParamByName("MODELOINSTANCIA").AsInteger = CurrentQuery.FieldByName("MODELOINSTANCIA").AsInteger 
        qry.Active = True 
        If (SessionVar("WORKFLOWSELECTION") = "0") Then 
          BOTAOSUSPENDER.Visible = True 
          If (qry.FieldByName("SITUACAO").AsInteger = 3) Then 
            BOTAOSUSPENDER.Enabled = True ' Minhas Tarefas 
          Else 
            BOTAOSUSPENDER.Enabled = False 
          End If 
        Else 
          BOTAOREATIVAR.Visible = True 
          If (qry.FieldByName("SITUACAO").AsInteger = 6) Then 
            BOTAOREATIVAR.Enabled = True ' Suspensas 
          Else 
            BOTAOREATIVAR.Enabled = False 
          End If 
        End If 
        qry.Active = False 
        Set qry = Nothing 
      End If 
      Q.Active = False 
      Set Q = Nothing 
    ElseIf (SessionVar("WORKFLOWSELECTION") = "3") Then ' Coordenacao 
      BOTAODELEGAR.Enabled = True 
      BOTAODELEGAR.Visible = True 
    End If 
  End If 
End Sub 
 
Function GetGuidInstance(ByVal handleModeloInstancia As Integer) As String 
  Dim Q As BPesquisa 
  Set Q = NewQuery 
    Q.Clear 
    Q.Add("SELECT GUID FROM Z_WFMODELOINSTANCIAS WHERE HANDLE = :HANDLE") 
    Q.ParamByName("HANDLE").AsInteger = handleModeloInstancia 
    Q.Active = True 
    GetGuidInstance = Q.FieldByName("GUID").AsString 
    Q.Active = False 
    Set Q = Nothing 
End Function 
 
Public Sub ShowCancelMessage(message) 
      If (VisibleMode = True) Then 
        MsgBox(message) 
      Else 
        CancelDescription = message 
      End If 
End Sub 
 
 
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean) 
  Dim Q As BPesquisa 
  Dim guid As String 
 
  Dim wfl As Workflow 
  Set wfl = NewWorkflow 
 
  ' <<<<<<<<<< EXECUTE >>>>>>>>>>> 
  If (CommandID = "EXECUTE") Then 
    On Error GoTo TryExceptExecute 
      wfl.Services.Instance.Resume(CurrentQuery.FieldByName("MENSAGEM").AsString, UserNickName) 
 
    GoTo FinallyExecute 
    TryExceptExecute: 
      ShowCancelMessage(Err.Description) 
      CanContinue = False 
      Set wfl = Nothing 
      Exit Sub 
    FinallyExecute: 
      If (VisibleMode = True) Then 
        ' Atualiza a CurrentQuery 
        CurrentQuery.Active = False 
        CurrentQuery.Active = True 
      End If 
      InfoDescription = "Tarefa realizada com sucesso!" 
  End If 
 
  ' <<<<<<<<<< ASSUMIR / DEVOLVER >>>>>>>>>>> 
  If (CommandID = "ASSUMIR_AUTO" Or CommandID = "ASSUMIR" Or CommandID = "DEVOLVER") Then 
    If (CurrentQuery.EOF) Then 
      Exit Sub 
    End If 
    On Error GoTo TryExceptAssumir 
	  If (CurrentQuery.FieldByName("RESPONSAVEL").AsInteger = CurrentUser And CommandID = "ASSUMIR") Then 
	  	GoTo FinallyAssumir 
	  End If 
 
      Dim assumir As Boolean 
      assumir = (CurrentQuery.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser) 
 
      If (assumir = True) Then 
        wfl.Services.Task.Associate(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentUser) 
      ElseIf (CommandID <> "ASSUMIR_AUTO") Then 
        wfl.Services.Task.Disassociate(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentUser) 
        CommandID = "DEVOLVER" 
      End If 
 
    GoTo FinallyAssumir 
    TryExceptAssumir: 
      ShowCancelMessage(Err.Description) 
      CanContinue = False 
      Set wfl = Nothing 
      Exit Sub 
    FinallyAssumir: 
      If (assumir And CommandID = "ASSUMIR") Then 
            BOTAOASSUMIR.Caption = "Devolver" 
            InfoDescription = "A tarefa foi assumida com sucesso!" 
        ElseIf (CommandID = "DEVOLVER") Then 
            BOTAOASSUMIR.Caption = "Assumir" 
            InfoDescription = "A tarefa foi devolvida com sucesso!" 
        End If 
 
      ' Atualiza a CurrentQuery 
      If (VisibleMode = True) Then 
        CurrentQuery.Active = False 
        CurrentQuery.Active = True 
      End If 
  End If 
  Set wfl = Nothing 
 
 
  ' <<<<<<<<<< SUSPENDER >>>>>>>>>>> 
  If (CommandID = "SUSPEND") Then 
      If (CurrentQuery.EOF) Then 
          Set wfl = Nothing 
          Exit Sub 
      End If 
 
    guid = GetGuidInstance(CurrentQuery.FieldByName("MODELOINSTANCIA").AsInteger) 
 
    Set wfl = NewWorkflow 
    wfl.Services.Instance.Parameters.Add("$WF_USUARIO$", UserNickName) 
 
    On Error GoTo RollbackPoint 
        wfl.Services.Instance.Suspend(guid) 
        InfoDescription = "A solicitação de suspensão foi enviada com sucesso!" 
        GoTo AfterRollbackPoint 
      RollBackPoint: 
	    ShowCancelMessage(Err.Description) 
        CanContinue = False 
      AfterRollbackPoint: 
        Set wfl = Nothing 
        BOTAOSUSPENDER.Enabled = False 
    End If 
 
 
  ' <<<<<<<<<< REATIVAR >>>>>>>>>>> 
  If (CommandID = "UNSUSPEND") Then 
      If (CurrentQuery.EOF) Then 
          Exit Sub 
      End If 
 
    guid = GetGuidInstance(CurrentQuery.FieldByName("MODELOINSTANCIA").AsInteger) 
 
    Set wfl = NewWorkflow 
    wfl.Services.Instance.Parameters.AddParam("$WF_USUARIO$", UserNickName) 
 
    On Error GoTo RollbackPoint1 
        wfl.Services.Instance.Unsuspend(guid) 
        InfoDescription = "A solicitação de reativação foi enviada com sucesso!" 
        GoTo AfterRollbackPoint1 
      RollBackPoint1: 
        ShowCancelMessage(Err.Description) 
        CanContinue = False 
      AfterRollbackPoint1: 
        Set wfl = Nothing 
        BOTAOREATIVAR.Enabled = False 
    End If 
 
 
 
  ' <<<<<<<<<< DELEGAR >>>>>>>>>>> 
  If (CommandID = "DELEGAR") Then 
      If (CurrentQuery.EOF) Then 
          Set wfl = Nothing 
          Exit Sub 
      End If 
      If (VisibleMode = False) Then 
        UsuarioDelegar = CurrentVirtualQuery.FieldByName("DESTINATARIO").AsInteger 
    End If 
 
    Set Q = NewQuery 
    Dim du As String 
    Q.Add("SELECT APELIDO FROM Z_GRUPOUSUARIOS WHERE HANDLE=:HANDLE") 
    Q.ParamByName("HANDLE").AsInteger = UsuarioDelegar 
    Q.Active = True 
    du = Q.FieldByName("APELIDO").AsString 
    Q.Active = False 
 
    Q.Clear 
    Q.Add("SELECT DETALHES FROM Z_WFMODELOINSTANCIAATIVIDADES WHERE HANDLE = :HANDLE") 
    Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ATIVIDADE").AsInteger 
    Q.Active = True 
    Dim d As String 
    d = Q.FieldByName("DETALHES").AsString + Chr(13) + Chr(10) + "+ " + CStr(Now) + ": tarefa delegada a '"+ du +"' por '"+ UserNickName +"'." 
    Q.Active = False 
 
    Q.Clear 
    Q.Add("UPDATE Z_WFMODELOINSTANCIAATIVIDADES SET DETALHES = :DETALHES WHERE HANDLE=:HANDLE") 
    Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ATIVIDADE").AsInteger 
    Q.ParamByName("DETALHES").AsMemo = d 
 
	Dim qAudit  As BPesquisa 
	Set qAudit = NewQuery 
    qAudit.Add("INSERT INTO Z_WFTAREFAAUDITORIAS (HANDLE,TAREFA,TIPO,USUARIO,DATAHORA) VALUES (:HANDLE,:TAREFA,:TIPO,:USUARIO,:DATAHORA)") 
    qAudit.ParamByName("HANDLE").AsInteger = NewHandle("Z_WFTAREFAAUDITORIAS") 
    qAudit.ParamByName("TAREFA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
    qAudit.ParamByName("TIPO").AsInteger = 1 ' 1-assumir; 2-Devolver 
    qAudit.ParamByName("USUARIO").AsInteger = UsuarioDelegar 
	   qAudit.ParamByName("DATAHORA").AsDateTime = ServerNow 
 
    StartTransaction 
 
    CurrentQuery.Edit 
    CurrentQuery.FieldByName("RESPONSAVEL").AsInteger = UsuarioDelegar 
    CurrentQuery.Post 
    Q.ExecSQL 
    qAudit.ExecSQL 
 
    Commit 
 
    Set Q = Nothing 
    Set qAudit = Nothing 
 
    If (VisibleMode = True) Then 
      CurrentQuery.Active = False 
      CurrentQuery.Active = True 
    End If 
  End If 
  Set wfl = Nothing 
End Sub 
