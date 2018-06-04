'HASH: C5BBDC5EDF61060B0C7B738701D7A278
 
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
  CurrentQuery.FieldByName("PAPELINCLUIDO").Clear 
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
  If (PAPELINCLUIDO.ReadOnly) Then 
    CanContinue = False 
    CancelDescription = "Registro não pode ser excluído" 
    If (VisibleMode) Then 
      MsgBox(CancelDescription) 
    End If 
  End If 
End Sub 
 
Public Sub TABLE_BeforeEdit(CanContinue As Boolean) 
  If (PAPELINCLUIDO.ReadOnly) Then 
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
 
Function TemFilhosEmLoop(iPapel As Integer, sPapeisVerificar As String) As Boolean 
Dim qLoop As Object 
Dim Filhos As String 
Dim EmLoop As Boolean 
 
  Filhos = "" 
  EmLoop = (CStr(iPapel) = sPapeisVerificar) 
 
  If (Not EmLoop) Then 
    Set qLoop = NewQuery 
    qLoop.Add("SELECT A.PAPELINCLUIDO FROM Z_PAPELPAPEIS A WHERE A.PAPEL IN (" + sPapeisVerificar + ")") 
    qLoop.Active = True 
    While (Not qLoop.EOF) 
      If (qLoop.FieldByName("PAPELINCLUIDO").AsInteger = iPapel) Then 
        EmLoop = True 
        Exit While 
      End If 
      If (Filhos = "") Then 
        Filhos = qLoop.FieldByName("PAPELINCLUIDO").AsString 
      Else 
        Filhos = Filhos + "," + qLoop.FieldByName("PAPELINCLUIDO").AsString 
      End If 
      qLoop.Next 
    Wend 
    qLoop.Active = False 
    Set qLoop = Nothing 
  End If 
 
  If (Not EmLoop) And (Filhos <> "") Then 
    EmLoop = TemFilhosEmLoop(iPapel, Filhos) 
  End If 
 
  TemFilhosEmLoop = EmLoop 
End Function 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  'Valida se tem problema de loop 
  If (TemFilhosEmLoop(CurrentQuery.FieldByName("PAPEL").AsInteger, CurrentQuery.FieldByName("PAPELINCLUIDO").AsString)) Then 
    CanContinue = False 
    CancelDescription = "Esta associação entre os papéis gera uma referência circular" 
    If (VisibleMode) Then 
      MsgBox(CancelDescription) 
    End If 
  End If 
End Sub 
