'HASH: 296B599786FE4BF620AC5060D1A3EED8
Public Sub CAMPOS_OnClick() 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishWFFieldsVisibilityByModel(CurrentSystem, -1, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
Public Sub EDITAR_OnClick() 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Benner.Tecnologia.Workflow.Application.DesignerForm") 
  Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("VALIDO").AsString <> "S") 
  ' caso haja inserção de modelo(s) devemos atualizar os registros na árvore 
  If ( Obj.ShouldRefreshNodesWithTable) Then 
    RefreshNodesWithTable ( "Z_WFMODELOS" ) 
  Else 
    SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False) 
  End If 
  Set Obj = Nothing 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  If EXCECAO.ResultFields = "NOME|" Then 
    EXCECAO.ResultFields = "NOME|VERSAO|" 
  End If 
End Sub 
 
Public Sub TABLE_BeforeDelete(CanContinue As Boolean) 
  Dim q As BPesquisa 
  Set q = NewQuery 
 
  Dim handle As Long 
  handle = CurrentQuery.FieldByName("HANDLE").AsInteger 
 
  q.Add("SELECT HANDLE FROM Z_WFMODELOINSTANCIAS WHERE MODELO=:MODELO") 
  q.ParamByName("MODELO").AsInteger = handle 
  q.Active = True 
  If (Not q.EOF) Then 
    q.Active = False 
    MsgBox("Não é possível excluir este fluxo pois o mesmo já foi utilizado para executar processos.") 
    GoTo CANCELAR 
  End If 
  q.Active = False 
 
  q.Clear 
  q.Add("SELECT HANDLE FROM Z_WFPROCESSOMODELOS WHERE MODELO=:MODELO") 
  q.ParamByName("MODELO").AsInteger = handle 
  q.Active = True 
  If (Not q.EOF) Then 
    q.Active = False 
    MsgBox("Não é possível excluir este fluxo pois o mesmo está relacionado com processos. Veja processos relacionados.") 
    GoTo CANCELAR 
  End If 
  q.Active = False 
 
 
  q.Clear 
  q.Add("DELETE FROM Z_WFMODELOTAREFAS WHERE MODELO=:HANDLE") 
  q.ParamByName("HANDLE").AsInteger = handle 
  q.ExecSQL 
 
  q.Clear 
  q.Add("DELETE FROM Z_WFMODELOPAPELPARTICIPANTES WHERE MODELOPAPEL IN(SELECT HANDLE FROM Z_WFMODELOPAPEIS WHERE MODELO=:HANDLE)") 
  q.ParamByName("HANDLE").AsInteger = handle 
  q.ExecSQL 
 
  q.Clear 
  q.Add("DELETE FROM Z_WFMODELOPAPEIS WHERE MODELO=:HANDLE") 
  q.ParamByName("HANDLE").AsInteger = handle 
  q.ExecSQL 
 
  q.Clear 
  q.Add("DELETE FROM Z_WFMODELOMACROS WHERE MODELO=:HANDLE") 
  q.ParamByName("HANDLE").AsInteger = handle 
  q.ExecSQL 
 
  q.Clear 
  q.Add("DELETE FROM Z_WFMODELOSUBMODELOS WHERE MODELO=:HANDLE") 
  q.ParamByName("HANDLE").AsInteger = handle 
  q.ExecSQL 
 
  GoTo FIM 
 
  CANCELAR: 
  CanContinue  = False 
 
  FIM: 
  Set q = Nothing 
End Sub 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  If (CurrentQuery.FieldByName("ATIVO").AsBoolean = True And CurrentQuery.FieldByName("VALIDO").AsBoolean = False) Then 
    CanContinue = False 
    If (VisibleMode) Then 
      MsgBox("Este fluxo não pode ser ativado pois contém erros na sua diagramação. Edite o fluxo para verificar os erros.") 
    Else 
      CancelDescription = "Este fluxo não pode ser ativado pois contém erros na sua diagramação. Edite o fluxo para verificar os erros." 
    End If 
  End If 
  If (CurrentQuery.FieldByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EXCECAO").AsInteger) Then 
    CanContinue = False 
    If (VisibleMode) Then 
      MsgBox("Não é permitido configurar o próprio fluxo como fluxo de tratamento de exceção. Altere o campo 'Em caso de erro executar o fluxo:'.") 
    Else 
      CancelDescription = "Não é permitido configurar o próprio fluxo como fluxo de tratamento de exceção. Altere o campo 'Em caso de erro executar o fluxo:'." 
    End If 
  End If 
 
  If (CurrentQuery.FieldByName("ATIVO").AsBoolean = True) Then 
    Dim q As BPesquisa 
    Set q = NewQuery 
    q.Add("SELECT COUNT(1) TOTAL FROM Z_WFMODELOSUBMODELOS A, Z_WFMODELOS B WHERE A.MODELO=:MODELO AND B.HANDLE = A.SUBMODELO AND B.ATIVO = 'N'") 
    q.ParamByName("MODELO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
    q.Active = True 
    If Not q.EOF Then 
      If (q.FieldByName("TOTAL").AsInteger > 0) Then 
        Err.Raise(vbsUserException, "", "Este fluxo contém subfluxos que não estão ativos. Para ativar este fluxo é necessário primeiro ativar seus subfluxos.") 
      End If 
    End If 
    q.Active = False 
    Set q = Nothing 
  Else 
    Set q = NewQuery 
    q.Add("SELECT B.NOME, B.VERSAO FROM Z_WFMODELOSUBMODELOS A, Z_WFMODELOS B WHERE A.SUBMODELO = :SUBMODELO AND B.HANDLE = A.MODELO AND B.ATIVO = 'S'") 
    q.ParamByName("SUBMODELO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
    q.Active = True 
    If Not q.EOF Then 
      If (q.FieldByName("NOME").AsString <> "") Then 
        Err.Raise(vbsUserException, "", "Não é possível desativar este fluxo pois ele está associado como subfluxo do fluxo '" + q.FieldByName("NOME").AsString + "' v." + q.FieldByName("VERSAO").AsString + " que está ativo.") 
      End If 
    End If 
    q.Active = False 
    Set q = Nothing 
  End If 
 
End Sub 
 
Public Sub TABLE_OnInsertBtnClick(CanContinue As Boolean) 
 
	Dim Obj As Object 
	Set Obj = CreateBennerObject("Benner.Tecnologia.Workflow.Application.DesignerForm") 
	Obj.Exec(CurrentSystem, -1, False) 
	' caso haja inserção de modelo(s) devemos atualizar os registros na árvore 
	If ( Obj.ShouldRefreshNodesWithTable ) Then 
	  RefreshNodesWithTable ( "Z_WFMODELOS" ) 
	End If 
	Set Obj = Nothing 
End Sub 
