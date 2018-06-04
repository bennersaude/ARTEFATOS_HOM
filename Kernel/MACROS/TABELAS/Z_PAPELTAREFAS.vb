'HASH: C868BF64E8BD9E27EB2BFBDAF5B4B6B7
 
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
 
Public Sub TABLE_AfterScroll() 
  ' So permite editar itens da camada corrente 
  If (CurrentQuery.State = 1) Then 
    RecordReadOnly = (CurrentQuery.FieldByName("CAMADA").AsInteger <> ObterCamadaCorrente) 
  Else 
    RecordReadOnly = False 
  End If 
End Sub 
 
Public Sub TABLE_BeforeDelete(CanContinue As Boolean) 
  If (TAREFA.ReadOnly) Then 
    CanContinue = False 
    CancelDescription = "Registro não pode ser excluído" 
    If (VisibleMode) Then 
      MsgBox(CancelDescription) 
    End If 
  End If 
End Sub 
 
Public Sub TABLE_BeforeEdit(CanContinue As Boolean) 
  If (TAREFA.ReadOnly) Then 
    CanContinue = False 
    CancelDescription = "Registro não pode ser alterado" 
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
