'HASH: CC2240348A6E584DA21D082AF65C0196
 
 
Public Sub BOTAOREATIVAR_OnClick() 
  Dim CanContinue As Boolean 
  CanContinue = True 
  TABLE_OnCommandClick "REACTIVATE", CanContinue 
End Sub 
 
Public Sub ShowCancelMessage(message) 
      If (VisibleMode = True) Then 
        MsgBox(message) 
      Else 
        CancelDescription = message 
      End If 
End Sub 
 
Public Sub BOTAOTERMINAR_OnClick() 
  Dim CanContinue As Boolean 
  CanContinue = True 
  TABLE_OnCommandClick "TERMINATE", CanContinue 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  BOTAOREATIVAR.Enabled = False 
  BOTAOREATIVAR.Visible = False 
  BOTAOTERMINAR.Enabled = False 
  BOTAOTERMINAR.Visible = False 
 
  If (WebMode) Then 
    IMAGEMWEB.Text = "@<div align='center'><a onclick='dlg_showWindow(this); return false;' href='workflowimage.aspx?t=u&guid="+ CurrentQuery.FieldByName("GUID").AsString +"'><img src='workflowimage.aspx?guid="+ CurrentQuery.FieldByName("GUID").AsString +"' border=0 title='Clique na imagem para ampliar'/></a></div>" 
  Else 
    IMAGEMWEB.Visible = False 
    Dim ws As WorkflowWebService 
    Set ws = NewWorkflowWebService 
    IMAGEMDESKTOP.Image = ws.InternalObject.GetWorkflowImage(CurrentSystem, CurrentQuery.FieldByName("GUID").AsString) 
    Set ws = Nothing 
  End If 
 
  If (CurrentQuery.FieldByName("SITUACAO").AsInteger <> 1 And CurrentQuery.FieldByName("SITUACAO").AsInteger <> 4 And CurrentQuery.FieldByName("SITUACAO").AsInteger <> 8 And CurrentQuery.FieldByName("SUPERIOR").IsNull) Then 
    If (CurrentQuery.FieldByName("SITUACAO").AsInteger = 5 Or CurrentQuery.FieldByName("SITUACAO").AsInteger = 7) Then 
      BOTAOREATIVAR.Enabled = True 
      BOTAOREATIVAR.Visible = True 
    End If 
    BOTAOTERMINAR.Enabled = True 
    BOTAOTERMINAR.Visible = True 
  End If 
End Sub 
 
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean) 
  Dim wfl As Workflow 
  Dim Q As BPesquisa 
 
  If (CommandID = "TERMINATE") Then 
      If (CurrentQuery.EOF) Then 
        Exit Sub 
      End If 
 
    Dim confirma As Boolean 
    confirma = False 
 
    If (WebMode = True) Then 
      confirma = RequestConfirmation("Deseja realmente encerrar a execução do fluxo? Isto irá parar todas as atividades do fluxo e não será possível reativar o fluxo.") 
    Else 
      If (MsgBox("Deseja realmente encerrar a execução do fluxo? Isto irá parar todas as atividades do fluxo e não será possível reativar o fluxo.", vbYesNo) = vbYes) Then 
        confirma = True 
      End If 
    End If 
 
      If (confirma = True) Then 
        Set wfl = NewWorkflow 
        wfl.Services.Instance.Parameters.Add("$WF_USUARIO$", UserNickName) 
        wfl.Services.Instance.Parameters.Add("$WF_CURRENTSTATUS$", CurrentQuery.FieldByName("SITUACAO").AsInteger) 
 
      On Error GoTo TerminateRollBack 
        wfl.Services.Instance.Terminate(CurrentQuery.FieldByName("GUID").AsString) 
        InfoDescription = "A solicitação de encerrar a instância foi enviada com sucesso!" 
      GoTo TerminateFim 
      TerminateRollBack: 
        ShowCancelMessage(Err.Description) 
        CanContinue = False 
      TerminateFim: 
        Set wfl = Nothing 
    End If 
  End If 
  If (CommandID = "REACTIVATE") Then 
      If (CurrentQuery.EOF) Then 
        Exit Sub 
      End If 
      Set wfl = NewWorkflow 
      wfl.Services.Instance.Parameters.Add("$WF_USUARIO$", UserNickName) 
      wfl.Services.Instance.Parameters.Add("$WF_CURRENTSTATUS$", CurrentQuery.FieldByName("SITUACAO").AsInteger) 
 
    On Error GoTo ReactivateRollBack 
      wfl.Services.Instance.Unsuspend(CurrentQuery.FieldByName("GUID").AsString) 
      InfoDescription = "A solicitação de reativar a instância foi enviada com sucesso!" 
    GoTo ReactivateFim 
    ReactivateRollBack: 
      ShowCancelMessage(Err.Description) 
      CanContinue = False 
    ReactivateFim: 
      Set wfl = Nothing 
  End If 
End Sub 
