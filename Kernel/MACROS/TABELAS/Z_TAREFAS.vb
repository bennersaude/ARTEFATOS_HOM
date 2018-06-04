'HASH: C80467E7941894F06EB784E6698E9BFC
 
Public Function LimpaStr(Value As String) 
  Caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_" 
  Value = UCase(Value) 
  LimpaStr = "" 
 
  For i = 1 To Len(Value) 
    If InStr(Caracteres, Mid(Value, i, 1)) Then 
      LimpaStr = LimpaStr + Mid(Value, i, 1) 
    End If 
  Next i 
End Function 
 
Public Function AdicionaPrefixo(Value As String, Prefixo As String) 
  If (Mid(Value, 1, Len(Prefixo)) <> Prefixo) Then 
    AdicionaPrefixo = Prefixo + RemovePrefixo(Value) 
  Else 
    AdicionaPrefixo = Value 
  End If 
End Function 
 
Public Function RemovePrefixo(Value As String) 
  If (Mid(Value, 1, 2) = "K_") Then 
    RemovePrefixo = Mid(Value, 3) 
  ElseIf (Mid(Value, 1, 3) = "K9_") Then 
    RemovePrefixo = Mid(Value, 4) 
  Else 
    RemovePrefixo = Value 
  End If 
End Function 
 
Public Function AdicionaPrefixoCamada(Value As String) 
  If IsServerDeveloper Then 
    If CustomSystem Then 
      AdicionaPrefixoCamada = AdicionaPrefixo(Value, "K9_") 
    Else 
      AdicionaPrefixoCamada = RemovePrefixo(Value) 
    End If 
  Else 
    AdicionaPrefixoCamada = AdicionaPrefixo(Value, "K_") 
  End If 
End Function 
 
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
 
Public Sub NOME_OnExit() 
  ' Ajusta identificador e adiciona prefixo 
  If (CurrentQuery.State <> 1) And (CStr(CurrentQuery.FieldByName("NOME").OldValue) <> CStr(CurrentQuery.FieldByName("NOME").NewValue)) Then 
    CurrentQuery.FieldByName("NOME").AsString = AdicionaPrefixoCamada(LimpaStr(CurrentQuery.FieldByName("NOME").AsString)) 
  End If 
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
  If (TITULO.ReadOnly) Then 
    CanContinue = False 
    CancelDescription = "Registro não pode ser excluído" 
    If (VisibleMode) Then 
      MsgBox(CancelDescription) 
    End If 
  End If 
End Sub 
 
Public Sub TABLE_BeforeEdit(CanContinue As Boolean) 
  If (TITULO.ReadOnly) Then 
    CanContinue = False 
    CancelDescription = "Registro não pode ser alterado" 
    If (VisibleMode) Then 
      MsgBox(CancelDescription) 
    End If 
  End If 
End Sub 
 
Public Sub TABLE_UpdateRequired() 
  NOME_OnExit 
  ' Grava a camada corrente no registro 
  If (CurrentQuery.FieldByName("CAMADA").AsInteger = 0) Then 
    CurrentQuery.FieldByName("CAMADA").AsInteger = ObterCamadaCorrente 
  End If 
End Sub 
