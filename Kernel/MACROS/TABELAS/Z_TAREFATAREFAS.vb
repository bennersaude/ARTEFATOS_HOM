'HASH: 953531502444CDD5C45917A3B1177CE4
 
Function ObterCamadaCorrente As Long 
  If (IsServerDeveloper) Then 
    If (CustomSystem) Then 
      ObterCamadaCorrente = 40 
    Else 
      ObterCamadaCorrente = 20 
    End If 
  Else 
    ObterCamadaCorrente = 50 
  End If 
End Function 
 
 
Public Sub TABLE_AfterInsert() 
  CurrentQuery.FieldByName("TAREFAINCLUIDA").Clear 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  ' So permite editar itens da camada corrente 
  If (CurrentQuery.State = 1) Then 
    RecordReadOnly = (CurrentQuery.FieldByName("CAMADA").AsInteger <> ObterCamadaCorrente) 
  Else 
    RecordReadOnly = False 
  End If 
End Sub 
 
Public Sub TABLE_BeforeDelete(CanContinue As Boolean) 
  If (TAREFAINCLUIDA.ReadOnly) Then 
    CanContinue = False 
    CancelDescription = "Registro não pode ser excluído" 
    If (VisibleMode) Then 
      MsgBox(CancelDescription) 
    End If 
  End If 
End Sub 
 
Public Sub TABLE_BeforeEdit(CanContinue As Boolean) 
  If (TAREFAINCLUIDA.ReadOnly) Then 
    CanContinue = False 
    CancelDescription = "Registro não pode ser alterado" 
    If (VisibleMode) Then 
      MsgBox(CancelDescription) 
    End If 
  End If 
End Sub 
 
Function TemFilhasEmLoop(iTarefa As Integer, sTarefasVerificar As String) As Boolean 
Dim qLoop As Object 
Dim Filhas As String 
Dim EmLoop As Boolean 
 
  Filhas = "" 
  EmLoop = (CStr(iTarefa) = sTarefasVerificar) 
 
  If (Not EmLoop) Then 
    Set qLoop = NewQuery 
    qLoop.Add("SELECT A.TAREFAINCLUIDA FROM Z_TAREFATAREFAS A WHERE A.TAREFA IN (" + sTarefasVerificar + ")") 
    qLoop.Active = True 
    While (Not qLoop.EOF) 
      If (qLoop.FieldByName("TAREFAINCLUIDA").AsInteger = iTarefa) Then 
        EmLoop = True 
        Exit While 
      End If 
      If (Filhas = "") Then 
        Filhas = qLoop.FieldByName("TAREFAINCLUIDA").AsString 
      Else 
        Filhas = Filhas + "," + qLoop.FieldByName("TAREFAINCLUIDA").AsString 
      End If 
      qLoop.Next 
    Wend 
    qLoop.Active = False 
    Set qLoop = Nothing 
  End If 
 
  If (Not EmLoop) And (Filhas <> "") Then 
    EmLoop = TemFilhasEmLoop(iTarefa, Filhas) 
  End If 
 
  TemFilhasEmLoop = EmLoop 
End Function 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  'Valida se tem problema de loop 
  If (TemFilhasEmLoop(CurrentQuery.FieldByName("TAREFA").AsInteger, CurrentQuery.FieldByName("TAREFAINCLUIDA").AsString)) Then 
    CanContinue = False 
    CancelDescription = "Esta associação entre as tarefas gera uma referência circular" 
    If (VisibleMode) Then 
      MsgBox(CancelDescription) 
    End If 
  End If 
End Sub 
 
Public Sub TABLE_UpdateRequired() 
  ' Grava a camada corrente no registro 
  If (CurrentQuery.FieldByName("CAMADA").AsInteger = 0) Then 
    CurrentQuery.FieldByName("CAMADA").AsInteger = ObterCamadaCorrente 
  End If 
End Sub 
